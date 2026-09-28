// api/e.ts — Analytics forwarder.
//
// The browser posts events to this same-origin path; this function relays them
// to our Umami instance. It replaces a static rewrite in vercel.json, for one
// reason: a rewrite has to name the destination in a committed file, and this
// repo is mirrored publicly. Here the address is an environment variable, so it
// is not in the repository at all and the mirror cannot leak it by omission.
//
// The client contract is unchanged — src/lib/analytics.ts still posts to
// /api/e, so the CSP still sees a same-origin request and there is still no
// third-party script anywhere in the page.
//
// Edge runtime: this runs near the visitor rather than in one fixed region,
// which keeps European traffic in Europe on its way to an instance in fra1.
// It also needs no Node APIs — everything below is Web-standard.
//
// No imports, on purpose. The allowlist below mirrors the client's instead of
// importing it, so nothing the SPA pulls in can break this function's bundle.
// Keep tests out of this folder: Vercel deploys every file in api/ as a
// function. They live in src/lib/__tests__/analyticsRelay.test.ts.

export const config = { runtime: 'edge' };

// No @types/node in this project, and an edge function has no Node APIs beyond
// this. Declaring the shape locally beats pulling in the whole Node surface.
declare const process: { env: Record<string, string | undefined> };

/** Ten enum-only events; none comes close. Anything larger is not ours. */
const MAX_BODY_BYTES = 2048;

/**
 * Not a secret — see UMAMI_WEBSITE_ID in src/lib/analytics.ts. Pinned here so
 * the relay writes to our site and to no other site on the instance.
 */
const WEBSITE_ID = 'c1726bec-39be-4544-97a1-ba25209fd0e8';

const MODES = ['live', 'sample'] as const;
const TRIGGERS = ['auto', 'manual'] as const;

/**
 * Baseline check ids are the one open set: the catalog grows, and this file
 * cannot import it. A kebab-case slug is what every id in it looks like.
 */
const CHECK_ID = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const MAX_CHECK_ID_LENGTH = 64;

/**
 * Mirror of ALLOWED in src/lib/analytics.ts. That one stops the app leaking;
 * this one stops everything that is not the app, because a public POST path
 * is reachable without ever loading the page. A test sends every event the
 * client can emit through this function, so an event added there and not here
 * fails rather than vanishing.
 */
const EVENTS: Readonly<Record<string, Readonly<Record<string, readonly string[] | RegExp>>>> = {
  view_opened: { view: ['grid', 'matrix', 'sankey', 'gaps', 'impact', 'baseline'], mode: MODES },
  source_chosen: { mode: MODES },
  load_outcome: { outcome: ['ok', 'consent_required', 'error'] },
  sandbox: { step: ['entered', 'changed', 'diffed', 'reset'] },
  baseline_fix: { checkId: CHECK_ID },
  export: { format: ['summary', 'powershell', 'json'], action: ['copy', 'download'] },
  evaluated: { mode: MODES, trigger: TRIGGERS },
  analysis_run: {
    view: ['gaps', 'impact'],
    trigger: TRIGGERS,
    personas: ['generic', 'selected', 'resolved', 'not_applicable'],
  },
  app_error: { source: ['uncaught', 'promise'] },
  tour: { step: ['started', 'completed', 'dismissed'] },
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** Exact in both directions — an extra key is rejected, not ignored. */
function hasExactKeys(value: Record<string, unknown>, expected: readonly string[]): boolean {
  return (
    Object.keys(value).length === expected.length &&
    expected.every((key) => Object.hasOwn(value, key))
  );
}

function accepts(allowed: readonly string[] | RegExp, value: string): boolean {
  if (allowed instanceof RegExp) {
    return value.length <= MAX_CHECK_ID_LENGTH && allowed.test(value);
  }
  return allowed.includes(value);
}

/**
 * The event a request carries, or null when it is not one of ours. Returns a
 * fresh object built from validated values, so nothing the caller sent is
 * forwarded by reference.
 */
function readEvent(parsed: unknown): { name: string; data: Record<string, string> } | null {
  if (!isRecord(parsed) || !hasExactKeys(parsed, ['type', 'payload'])) return null;
  if (parsed.type !== 'event') return null;

  const payload = parsed.payload;
  if (!isRecord(payload) || !hasExactKeys(payload, ['website', 'name', 'data', 'url'])) return null;
  if (payload.website !== WEBSITE_ID || payload.url !== '/') return null;

  // hasOwn, not a lookup: 'constructor' must not resolve to Object's.
  const { name, data: sent } = payload;
  if (typeof name !== 'string' || !Object.hasOwn(EVENTS, name)) return null;

  const spec = EVENTS[name];
  if (!isRecord(sent) || !hasExactKeys(sent, Object.keys(spec))) return null;

  const data: Record<string, string> = {};
  for (const [key, allowed] of Object.entries(spec)) {
    const value = sent[key];
    if (typeof value !== 'string' || !accepts(allowed, value)) return null;
    data[key] = value;
  }
  return { name, data };
}

export default async function handler(request: Request): Promise<Response> {
  if (request.method !== 'POST') {
    return new Response(null, { status: 405, headers: { Allow: 'POST' } });
  }

  const upstream = process.env.UMAMI_URL;
  if (!upstream) {
    // Misconfigured rather than misused. The client swallows every failure, so
    // this is invisible to users but plain in the network tab when looking.
    return new Response(null, { status: 503 });
  }

  // Bytes, not characters: a string's length undercounts anything non-ASCII.
  const raw = await request.arrayBuffer();
  if (raw.byteLength > MAX_BODY_BYTES) return new Response(null, { status: 413 });

  let event: ReturnType<typeof readEvent>;
  try {
    event = readEvent(JSON.parse(new TextDecoder().decode(raw)));
  } catch {
    return new Response(null, { status: 400 });
  }
  if (!event) return new Response(null, { status: 400 });

  try {
    const res = await fetch(upstream, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        // Umami derives browser and device from the user-agent, and needs a
        // real one at all — it drops requests without it.
        'User-Agent': request.headers.get('user-agent') ?? '',
        // Without this the visitor hash would be computed from this function's
        // address, collapsing every visitor into one and quietly making the
        // visitor counts wrong rather than absent.
        'X-Forwarded-For': request.headers.get('x-forwarded-for') ?? '',
      },
      // Rebuilt, never relayed. Umami also accepts referrer, screen, language
      // and title; the request body is the caller's, so forwarding it would
      // let anyone who skips the page put those fields in the dataset.
      body: JSON.stringify({
        type: 'event',
        payload: { website: WEBSITE_ID, name: event.name, data: event.data, url: '/' },
      }),
    });
    return new Response(null, { status: res.ok ? 204 : 502 });
  } catch {
    // Upstream unreachable — a dropped analytics event is not worth an error.
    return new Response(null, { status: 502 });
  }
}
