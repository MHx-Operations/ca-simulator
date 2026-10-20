// lib/__tests__/analyticsRelay.test.ts — The /api/e edge function.
//
// The client's allowlist stops the app leaking. It does nothing about a caller
// that never loads the app: /api/e is a public POST path, and whatever it
// forwards lands in the dataset. These tests pin both halves of the relay's
// job — every event the client can emit gets through, and nothing else does.
//
// This file lives here rather than beside the function because Vercel deploys
// every file under api/ as a function.

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import handler from '../../../api/e';
import { ALLOWED, trackEvent, resetSandboxChangeTracking } from '../analytics';
import { BASELINE_CHECKS } from '@/data/baselineChecks';

const UPSTREAM = 'https://analytics.example.test/api/send';
const WEBSITE_ID = 'c1726bec-39be-4544-97a1-ba25209fd0e8';

// The network boundary is mocked; our own logic never is.
const mockFetch = vi.fn((_url: string, _init?: RequestInit) =>
  Promise.resolve(new Response(null, { status: 200 })),
);

function event(name: string, data: Record<string, unknown>) {
  return { type: 'event', payload: { website: WEBSITE_ID, name, data, url: '/' } };
}

function post(body: unknown, headers: Record<string, string> = {}): Promise<Response> {
  return handler(
    new Request('https://ca.example.test/api/e', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...headers },
      body: typeof body === 'string' ? body : JSON.stringify(body),
    }),
  );
}

/** What the relay put on the wire to Umami. */
function forwarded(call = 0) {
  const [url, init] = mockFetch.mock.calls[call];
  return { url, init, body: JSON.parse(init?.body as string) };
}

/** Every combination of allowed values for one event. */
function combinations(spec: Readonly<Record<string, readonly string[]>>): Record<string, string>[] {
  return Object.entries(spec).reduce<Record<string, string>[]>(
    (rows, [key, values]) => rows.flatMap((row) => values.map((value) => ({ ...row, [key]: value }))),
    [{}],
  );
}

beforeEach(() => {
  mockFetch.mockClear();
  vi.stubGlobal('fetch', mockFetch);
  vi.stubEnv('UMAMI_URL', UPSTREAM);
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.unstubAllEnvs();
});

describe('relay — stays in step with the client', () => {
  it.each(Object.keys(ALLOWED))('accepts every %s the client can emit', async (name) => {
    // Guards drift: an event or value added to ALLOWED and not to the relay
    // would be dropped in production with nothing to show for it.
    for (const data of combinations(ALLOWED[name])) {
      mockFetch.mockClear();
      const res = await post(event(name, data));
      expect(res.status, `${name} ${JSON.stringify(data)}`).toBe(204);
      expect(forwarded().body).toEqual(event(name, data));
    }
  });

  it('accepts every id in the baseline catalog', async () => {
    for (const check of BASELINE_CHECKS) {
      const res = await post(event('baseline_fix', { checkId: check.id }));
      expect(res.status, check.id).toBe(204);
    }
  });

  it('forwards what the client actually sends, unchanged', async () => {
    // The real output of trackEvent() goes in, so the website id, the constant
    // url and the payload shape cannot drift between the two files.
    const clientFetch = vi.fn((_url: string, _init?: RequestInit) =>
      Promise.resolve(new Response(null, { status: 204 })),
    );
    vi.stubGlobal('window', { doNotTrack: undefined });
    vi.stubGlobal('navigator', { doNotTrack: undefined });
    vi.stubGlobal('localStorage', { getItem: () => null, setItem: () => {}, removeItem: () => {} });
    vi.stubGlobal('fetch', clientFetch);
    resetSandboxChangeTracking();
    trackEvent({ name: 'analysis_run', props: { view: 'gaps', trigger: 'manual', personas: 'resolved' } });
    const clientBody = clientFetch.mock.calls[0][1]?.body as string;

    vi.stubGlobal('fetch', mockFetch);
    const res = await post(clientBody);

    expect(res.status).toBe(204);
    expect(forwarded().body).toEqual(JSON.parse(clientBody));
  });
});

describe('relay — rejects anything the client could not have sent', () => {
  async function expectRejected(body: unknown) {
    const res = await post(body);
    expect(res.status, JSON.stringify(body)).toBe(400);
    expect(mockFetch).not.toHaveBeenCalled();
  }

  it('rejects an unknown event name', async () => {
    await expectRejected(event('policy_viewed', { name: 'CA001 - Require MFA' }));
    await expectRejected(event('', {}));
  });

  it('rejects names that exist on every object', async () => {
    // A plain lookup would find Object.prototype members and an empty spec.
    await expectRejected(event('constructor', {}));
    await expectRejected(event('toString', {}));
    await expectRejected(event('__proto__', {}));
  });

  it('rejects free-form values', async () => {
    await expectRejected(event('view_opened', { view: 'Finance Policies', mode: 'live' }));
    await expectRejected(event('source_chosen', { mode: 'contoso.onmicrosoft.com' }));
    await expectRejected(event('load_outcome', { outcome: 'Insufficient privileges' }));
    await expectRejected(event('baseline_fix', { checkId: 'Require MFA for Finance' }));
    await expectRejected(event('baseline_fix', { checkId: 'a1b2c3d4-0000-0000-0000-00000000000G' }));
    await expectRejected(event('baseline_fix', { checkId: `a${'-a'.repeat(40)}` }));
    await expectRejected(event('baseline_fix', { checkId: '' }));
  });

  it('rejects extra, missing and non-string props', async () => {
    await expectRejected(event('view_opened', { view: 'grid', mode: 'live', tenantName: 'Contoso' }));
    await expectRejected(event('view_opened', { view: 'grid' }));
    await expectRejected(event('load_outcome', { outcome: 42 }));
    await expectRejected(event('load_outcome', { outcome: ['ok'] }));
    await expectRejected(event('sandbox', { step: null }));
    await expectRejected({ type: 'event', payload: { website: WEBSITE_ID, name: 'tour', data: ['started'], url: '/' } });
  });

  it.each(['referrer', 'hostname', 'screen', 'language', 'title', 'id', 'tag'])(
    'rejects a payload carrying %s',
    async (field) => {
      // Umami accepts all of these. The client sends none, so a request that
      // has one did not come from the client.
      const body = event('load_outcome', { outcome: 'ok' });
      await expectRejected({ ...body, payload: { ...body.payload, [field]: 'https://contoso.sharepoint.com' } });
    },
  );

  it('rejects another website id — the relay writes to one site only', async () => {
    const body = event('load_outcome', { outcome: 'ok' });
    await expectRejected({ ...body, payload: { ...body.payload, website: '00000000-0000-4000-8000-000000000000' } });
  });

  it('rejects a real page url', async () => {
    const body = event('load_outcome', { outcome: 'ok' });
    await expectRejected({ ...body, payload: { ...body.payload, url: '/?tenant=contoso' } });
  });

  it('rejects other Umami request types and extra top-level keys', async () => {
    const body = event('load_outcome', { outcome: 'ok' });
    await expectRejected({ ...body, type: 'identify' });
    await expectRejected({ ...body, extra: true });
  });

  it('rejects bodies that are not a JSON object', async () => {
    await expectRejected('not json');
    await expectRejected('');
    await expectRejected('null');
    await expectRejected('[]');
    await expectRejected('"event"');
  });
});

describe('relay — transport', () => {
  it('answers only POST', async () => {
    const res = await handler(new Request('https://ca.example.test/api/e'));
    expect(res.status).toBe(405);
    expect(res.headers.get('Allow')).toBe('POST');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('answers 503 when the destination is not configured', async () => {
    vi.stubEnv('UMAMI_URL', '');
    const res = await post(event('load_outcome', { outcome: 'ok' }));
    expect(res.status).toBe(503);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('measures the body in bytes, not characters', async () => {
    // 1,500 characters, 3,000 bytes: under the limit by one count, over by the other.
    const res = await post(event('baseline_fix', { checkId: 'é'.repeat(1500) }));
    expect(res.status).toBe(413);
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('sends to the configured destination with three headers and no others', async () => {
    await post(event('load_outcome', { outcome: 'ok' }), {
      'User-Agent': 'Mozilla/5.0 (test)',
      'X-Forwarded-For': '203.0.113.7',
      Referer: 'https://contoso.sharepoint.com/sites/security',
      Cookie: 'session=abc',
      Authorization: 'Bearer not-ours',
    });

    const { url, init } = forwarded();
    expect(url).toBe(UPSTREAM);
    expect(init?.method).toBe('POST');
    expect(init?.headers).toEqual({
      'Content-Type': 'application/json',
      'User-Agent': 'Mozilla/5.0 (test)',
      'X-Forwarded-For': '203.0.113.7',
    });
    expect(init?.body as string).not.toContain('contoso');
  });

  it('answers 502 when the destination refuses or cannot be reached', async () => {
    mockFetch.mockResolvedValueOnce(new Response(null, { status: 500 }));
    expect((await post(event('load_outcome', { outcome: 'ok' }))).status).toBe(502);

    mockFetch.mockRejectedValueOnce(new Error('unreachable'));
    expect((await post(event('load_outcome', { outcome: 'ok' }))).status).toBe(502);
  });
});
