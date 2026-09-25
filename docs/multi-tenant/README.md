# Deploying the tenant list

The tenant picker (see [Multiple tenants](../../README.md#multiple-tenants))
reads its list of tenants from a JSON file. How you serve that file decides who
can read it. This folder covers the two ways to deploy it.

| | Option A: plain file | Option B: behind a sign-in |
|---|---|---|
| Setup | One volume mount | oauth2-proxy sidecar, an extra app registration, a TLS reverse proxy |
| Who can read the list | **Anyone who knows the URL** | Only accounts that can sign in to your own tenant |
| Who can open the app | Anyone | Only accounts that can sign in to your own tenant |
| Suits | Local use, private networks | A host reachable from the internet |

## Where the app looks for the list

The app fetches the list from `VITE_TENANTS_URL`, or from `/tenants.json` when
that variable is unset or empty. Set it to a path on the app's own origin
(such as `/private/tenants.json`): the request carries no credentials of its
own, so protection has to come from the browser's session cookie for that
origin. The image runs the Vite dev server, which reads `VITE_*` variables when
the container is created, so after changing the variable recreate the
container with `docker compose ... up -d`. A plain `restart` keeps the old
value.

If the fetch fails for any reason (missing file, 401, 403, a page instead of
JSON), the app behaves exactly as it does with no tenants configured: no
picker and no error message. The browser's network log still shows the failed
request.

## Option A: plain file

Mount the list into the app's `public/` folder:

```yaml
services:
  ca-simulator:
    volumes:
      - ./tenants.json:/app/public/tenants.json:ro
```

The list is then a public static asset at `https://<host>/tenants.json`. It
contains your customers' names and tenant IDs, and **anyone who can reach the
host can download it** without signing in. Signing in to the app does not
protect it: the file is served before, and independently of, any sign-in.

Use this only where the host itself is not reachable by people who should not
see the list.

## Option B: behind a sign-in (oauth2-proxy)

[oauth2-proxy](https://oauth2-proxy.github.io/oauth2-proxy/) sits in front of
the app and lets through only browsers that have signed in to your own Entra
tenant. It serves both the app and the list, so nothing is reachable without
that sign-in:

```
browser ── HTTPS ──▶ your reverse proxy (TLS) ── HTTP ──▶ oauth2-proxy
                                                            ├─ /private/tenants.json ─▶ list file (read-only mount)
                                                            └─ everything else ───────▶ ca-simulator
```

Users then sign in twice in a row: once at the sidecar, then as usual in the
app. The second one usually completes without a prompt, because Entra already
has a session.

### Prerequisites

1. **Docker Compose 2.24 or later.** The overlay uses `!reset` to remove the
   app's published port; older versions would leave it published, which
   bypasses the sign-in.
2. **A TLS-terminating reverse proxy** (nginx, Caddy, Traefik, ...) for your
   host name, forwarding to the sidecar and setting `X-Forwarded-Proto` and
   `X-Forwarded-For`. The sidecar speaks plain HTTP and must not be exposed
   directly.
3. **A separate app registration for the sidecar.** Don't reuse the app's own
   registration: the sidecar needs a client secret, and the app's registration
   is multi-tenant and consented in your customers' tenants.
   - Supported account types: *Accounts in this organizational directory only*
   - Platform **Web** (not SPA), redirect URI `https://<host>/oauth2/callback`
   - Certificates & secrets: create a client secret
   - Leave *ID tokens (used for implicit and hybrid flows)* unticked; the
     sidecar uses the authorization code flow.
   - Guest (B2B) accounts in your tenant can sign in too. To limit access to
     specific people, set *Assignment required* on the registration's
     enterprise application and assign users or groups, or narrow
     `OAUTH2_PROXY_EMAIL_DOMAINS`.
4. **The app's own registration** gets `https://<host>` as an SPA redirect URI,
   as for any deployment.

### Steps

1. In the repository checkout, copy [`.env.example`](.env.example) to `.env` in
   the repository root and fill in every empty value. `.env` is gitignored,
   and `.dockerignore` keeps it and any tenant list out of the image built from
   the checkout.
2. Put your list outside the repository (format:
   [`public/tenants.example.json`](../../public/tenants.example.json)) and set
   `TENANTS_FILE` to its path. It must exist before the first start, or Docker
   creates an empty directory in its place.
3. Start, from the repository root:
   ```bash
   docker compose -f docker-compose.yml -f docs/multi-tenant/docker-compose.hardened.yml up -d
   ```
   Compose refuses to start while a required value is missing from `.env`.
4. Point your reverse proxy at the sidecar (`OAUTH2_PROXY_BIND`, default
   `127.0.0.1:4180`).

To run a prebuilt image instead of building from the checkout, add
`image: <your image>` to the `ca-simulator` service in `docker-compose.yml` and
start with `up -d --no-build`.

The sidecar session lasts `OAUTH2_PROXY_COOKIE_EXPIRE` (default 12 hours) and
is not re-checked against Entra in between, so that is how long a disabled
account keeps access to the list.

### Check it

Without a session cookie (a private window, or `curl`):

```bash
curl -sS -o /dev/null -w '%{http_code}\n' https://<host>/
curl -sS -o /dev/null -w '%{http_code}\n' https://<host>/private/tenants.json
curl -sS -o /dev/null -w '%{http_code}\n' https://<host>/tenants.json
```

All three must answer `302` (a redirect to the Entra sign-in). A `200` means
the list or the app is reachable without signing in. The app's own port must
not answer at all from outside the host:

```bash
curl -sS -m 5 http://<host>:5173/   # expect a connection failure
```

Then sign in with an account of your tenant: the tenant picker appears. An
account from any other tenant must be refused.

### A misconfiguration looks like "no tenants"

If the sidecar refuses the list request (wrong path, expired session, sidecar
down), the app cannot tell that apart from "no list configured" and quietly
shows no picker. That is deliberate: the app must never surface an error that
reveals the list exists. So a missing picker is the symptom to check the
sidecar's logs for:

```bash
docker compose -f docker-compose.yml -f docs/multi-tenant/docker-compose.hardened.yml logs oauth2-proxy
```

### Troubleshooting

| Symptom | Cause |
|---|---|
| `docker compose` stops with `required variable ... is missing a value` | That value is empty in `.env` |
| The list path answers `http: attempting to traverse a non-directory` | A `file://` upstream points at a file instead of a directory; keep the upstream as shipped |
| Entra shows `AADSTS50011` (redirect URI mismatch) | `https://<host>/oauth2/callback` is not registered on the sidecar's registration under platform **Web**, or `PUBLIC_HOST` differs from the host name users open |
| `Blocked request. This host ("…") is not allowed.` | `PUBLIC_HOST` does not match the host name in the browser |
| Signed in, but no tenant picker | See [A misconfiguration looks like "no tenants"](#a-misconfiguration-looks-like-no-tenants) |
