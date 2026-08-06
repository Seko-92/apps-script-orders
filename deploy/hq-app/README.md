# Floor Board → VPS  (slice 1 of the web-app move)

**Pure transport change. The board looks and behaves exactly the same.**
Only how it reaches the server moves:

```
before:   FloorBoard.html on /exec  ──google.script.run──→  Apps Script
after:    FloorBoard.html on Caddy  ──fetch /api/board──→  n8n  ──→  Apps Script
```

`google.script.run` exists only inside Google's HtmlService frame, so it cannot
survive the move. Building that HTTP pipe **is the point** — it's the same
infrastructure kit expansion needs next, proven first against a board whose
correct behaviour we already know by heart.

## It still works in BOTH places

`hqCall()` in `FloorBoard.html` detects its host: native `google.script.run`
when Apps Script serves it, `fetch()` otherwise. **The move is reversible** —
point the tablet back at `/exec` any time, same file, no edits. Keep both live
for a few days and compare.

## What you get

* No more **New Version** cut to change the board — edit, upload, refresh
* `hq.yassinqurabi.com` instead of the `AKfycb…` URL (Add-to-Home-Screen
  becomes trivial, which was the original tablet friction)
* The data pipe every later page reuses

## What gets worse — honestly

n8n joins the board's critical path. n8n down = board dark. Today the
equivalent is Apps Script being down, so it's comparable rather than free.

---

# Steps

### 1. DNS
A record `hq.yassinqurabi.com` → server IP, **DNS-only (grey cloud)**.
Orange-cloud proxying breaks Caddy's HTTP-01 challenge — the same trap from the
original n8n setup.

**It's ONE app, not one page per site.** `/` is the board, `/kits` will be kit
expansion, `/api/*` is n8n. Named for the app because the board is simply what
lives at the root — decided before DNS on purpose, since the tablet's
home-screen icon points at whatever hostname we pick.

### 2. Apps Script — cut ONE New Version
Three new `doPost` actions (`boardTick`, `boardStatus`, `boardRadio`) are
already pushed. They live on `/exec`, so they need:

> Manage Deployments → ⋮ → Edit → **New version** → Deploy

**Never "New Deployment"** — that mints a fresh URL and orphans all 11
hardcoded workflow nodes.

This is the **last** version cut the board will ever need.

### ⚡ The proxy CACHES reads — and that is what makes multiple screens safe

Every device polls independently, and each poll is its own Apps Script execution.
A consumer Google account has **~90 minutes of script runtime per day**, so at
~4s a tick one board burns roughly **12 minutes per hour**:

| Devices | Daily quota runs dry after |
|---|---|
| 1 | ~7 hours |
| 3 | ~2.5 hours |
| 10 | **~45 minutes** |

And when it runs dry, *everything* Apps Script touches fails — n8n inserts,
Telegram commands, not just the board. (This ceiling is not new and was not
caused by the move; `google.script.run` spent exactly the same runtime. The
board simply never had enough viewers to reach it.)

Node **2. Cache + Forward** holds the last tick in workflow static data for 15s
and serves every device from that copy, so **Apps Script cost is independent of
how many screens are open** — one board or fifty, the same few calls per minute.
Most polls also come back instantly from n8n instead of waiting on a sheet read.

Two rules baked into that node, both load-bearing:
* **`CACHEABLE` is reads only.** `boardStatus` (the ✓ Pick write) is never served
  from — or stored in — the cache.
* **Only a response that actually looks like a tick gets cached** (`res.cockpit`
  present). Caching an error would pin every board to it for the whole window.

Debugging aid: a cached response carries `_cached: true` and `_ageMs`.

### 3. n8n — import `hq-board-proxy.n8n.json`
Then replace the two placeholders in node **2. Forward to Apps Script**:

| placeholder | value |
|---|---|
| `__APPS_SCRIPT_EXEC_URL__` | `WEB_APP_URL` from `Secrets.js` |
| `__APP_SECRET_TOKEN__` | `APP_SECRET_TOKEN` from `Secrets.js` |

**Activate the workflow.** A production webhook only answers when active — a
404 here is an inactive workflow, not an auth problem.

Three nodes: **Webhook** (POST, path `hq-board`, Respond = *Using Respond to
Webhook node*) → **Code** (`2. Cache + Forward`, the caching + fetch logic) →
**Respond to Webhook** (JSON, `={{ JSON.stringify($json) }}`).

⚠ **The placeholders now live inside the Code node**, near the top — not in a
URL field. Open `2. Cache + Forward` and replace both `EXEC_URL` and `TOKEN`.

⚠ Static data only persists on **production** executions (a real webhook call),
not on manual "Execute workflow" runs — so test by loading the board or curling
the endpoint, not with the editor's run button.

Cold-cache note: several devices arriving simultaneously on an empty cache will
each miss and call Apps Script once. Harmless — they desynchronise immediately
and it costs one extra call, not a storm.

### 4. Caddy — config + the mount
Paste `Caddyfile.snippet` into `/opt/caddy/Caddyfile`.

The static files also need to be **mounted into the Caddy container** — Caddy
only sees what's mounted, not the whole host filesystem. In
`/opt/caddy/docker-compose.yml`, add the bind mount to the **service's**
`volumes:` list (NOT the top-level `volumes:` block, which only declares named
volumes and will error):

```yaml
    volumes:
      - ./Caddyfile:/etc/caddy/Caddyfile
      - ./site:/srv:ro
      - /opt/hq-app:/opt/hq-app:ro      # <- add this line
      - caddy_data:/data
      - caddy_config:/config
```

Then:

```
cd /opt/caddy
docker compose exec caddy caddy validate --config /etc/caddy/Caddyfile
docker compose up -d caddy                       # recreates: brief blip on ALL sites
docker compose exec caddy ls /opt/hq-app         # must list index.html
```

Do the recreate **outside Houston hours** — it briefly drops n8n and tracker too.

### 5. Upload the page
```
mkdir -p /opt/hq-app
# from the project folder:
scp FloorBoard.html hetzner:/opt/hq-app/index.html
```

### 6. Test, in this order
1. `https://hq.yassinqurabi.com` loads and paints numbers
   → the whole pipe works
2. Watch it for one minute — the clock ticks, the poll refreshes
   → `boardTick` is flowing every 15s
3. Tap **✓ Pick** on a PENDING order → it moves to PREP, undo toast appears
   → `boardStatus` writes correctly through the chain
4. Turn the radio on → a track name appears under the station
   → `boardRadio` works (this one is easy to forget)

If step 1 fails, check n8n's Executions panel first — that tells you whether
the request reached n8n at all, which splits "Caddy problem" from
"Apps Script problem" immediately.

---

## Security note

**Exposure is unchanged.** The board was already public on `/exec`; it's public
on the new domain too. `boardSetStatus` remains narrowed **server-side** to
PENDING / PREPARING only — it cannot ship, cancel or delete regardless of who
calls it. That allow-list is the security boundary, exactly as before.

The shared token lives in n8n and never reaches the browser.

The one real change is **discoverability**: a short memorable hostname is easier
to stumble on than an `AKfycb…` URL. If that ever matters, the fix is a Caddy
`basic_auth` on `/api/board` — but it would also have to be added to the
picker's tablet, so it's a deliberate trade, not an obvious win.
