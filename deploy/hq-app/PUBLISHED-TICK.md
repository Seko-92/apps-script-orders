# Published Tick — design note

**Problem.** The board recomputes everything on every poll, on every device, and
Apps Script runs one execution at a time. Cost scales with **viewers**, not with
**activity**. Two devices already queue behind each other; the quota is the wall.

**Move.** Compute once, write the answer down, let everyone read it.
This is not a new pattern here — Kit Health, Out of Stock and Price Audit are
already computed views. The tick is the one hot read that never got the
treatment.

```
Apps Script  ──computes──►  __Published!A1 (JSON)
                                   │
n8n  ──reads that ONE cell─────────┘   (no logic, nothing to drift)
                                   │
Board  ◄───────────────────────────┘   instant
```

---

## 1. What goes in the payload

Same shape `getDashboardTick()` returns today, with **one rule that changes**:

> **Publish FACTS, not elapsed time.**

Anything time-derived is stored as a **timestamp** and the age is computed by the
board's own clock. `oldestReceivedAt`, not `"59m"`. `lastSyncAt`, not
`lastSyncMinutes`.

Why it matters: elapsed values go stale on their own, with no event to react to.
Storing the fact makes the payload correct until the underlying data changes —
which is what lets publishing be event-driven at all. The board already ticks a
clock every second; let it do the subtraction.

**Size budget.** A cell holds 50,000 characters (Gotcha #13 — the limit that
broke the Zoho payload cache). Estimate: 60 pick rows x ~120 chars = ~7KB, plus
cockpit/alerts/timeline ~5KB = **~15KB**, comfortably inside.

Guard anyway: if the JSON exceeds **45,000** chars, drop the timeline first, then
trim the pick list, and set `trimmed:true` so the board can say so. Never write a
payload that throws — a failed write leaves the board on the last good copy.

## 2. Where it lives

A hidden **`__Published`** sheet:

| cell | holds |
|---|---|
| `A1` | the tick JSON |
| `A2` | publishedAt (a real Date) |
| `A3` | last publish error, if any |

Its own sheet rather than a corner of `__SparkData`, because the purpose should
be obvious to anyone who unhides it, and because a second published payload
(reports, a future surface) gets a second row instead of a new home.

**Only Apps Script writes it.** n8n and the board are readers, always.

## 3. When it is recomputed

**Chokepoints mark it dirty; a timer publishes.**

* `updateOrderStatus()` and the `doPost` insert path set a dirty flag
  (a Script Property write, ~50ms). These already bust the n8n tick cache today,
  so they are known-correct chokepoints.
* A **2-minute trigger** publishes *only if dirty*, then clears the flag.

Publishing inline at the chokepoint was rejected: it would add ~5s to every
status flip, and a Telegram PREP tap must stay instant.

Cost, roughly:

| | today | published |
|---|---|---|
| scales with | **viewers x poll rate** | **order activity** |
| 1 board, all day | ~4,300 executions | ~0 for readers |
| 10 boards | 10x | unchanged |
| quiet evening | still polling | flag check only |

## 4. When it is missing or stale — the board must never go dark

Ordered fallback, and this is what makes the cutover zero-risk:

1. payload present and `publishedAt` fresh -> serve it
2. payload stale beyond a threshold, or missing -> **n8n calls the existing
   `boardTick` action**, exactly as today
3. that fails too -> the board keeps its last good tick and shows the live
   indicator as stale (already implemented)

The old path is not deleted. It becomes the safety net, which means this can ship
without a flag day and be reverted by pointing n8n back.

## 5. What n8n does

Read one cell. Return it. Fall back if empty or stale.

**No business logic, ever.** That is the whole point of choosing "publish the
result" over "reimplement the computation in n8n": the tick's logic (boundary
detection, natural aisle sort, kit detection, aging, paid-shipping counts,
Activity Log tallies) is 200-300 lines that would otherwise live in a
browser-edited Code node on the VPS — not version-controlled, not diffable, and
not editable by Claude. That is a maintenance tax forever, for a latency win we
can get without it.

---

## Explicitly NOT in this change

* **The drawer stays on Apps Script**, behind progressive loading. A dossier must
  be current — someone reads "on hand 3" and walks to a shelf — and 3,500 SKUs
  will not fit the publish-to-a-cell pattern anyway.
* **No Postgres.** Revisit when we want per-SKU published data or a query Sheets
  genuinely cannot answer (the ripple as an indexed lookup). Not because one
  evening was slow.
* **No change to how the sheet works.** It stays the system of record and the
  operator UI.

## Order of work

1. `publishBoardTick()` + the `__Published` sheet + the dirty flag  (Apps Script, testable)
2. 2-minute trigger, gated on the dirty flag
3. n8n reads the cell, falls back to `boardTick`
4. Board consumes timestamps instead of pre-computed ages
5. Watch for a day with the fallback live, then consider raising the poll interval

Steps 1-2 are useful alone: even with n8n unchanged, the payload exists and can
be inspected.


---

# Step 3 — n8n reads the cell  (SHIPPED 2026-08-07)

`hq-board-proxy.n8n.json` now has **four** nodes:

```
1. Board Request  ->  2. Read Published Tick  ->  3. Serve or Forward  ->  4. Respond
   (webhook)          (Sheets API, one cell)     (three tiers)
```

**Three tiers, cheapest first:**

| tier | source | when |
|---|---|---|
| 1 | n8n static cache (15s) | repeat polls inside the window |
| 2 | **the published cell** | payload present and under 10 min old |
| 3 | Apps Script live | missing / stale / malformed, or any non-tick action |

Tier 3 is a **safety net, not dead code**. The board cannot go dark because of
this change, and reverting is pointing the flow back at it.

**Only `boardTick` is ever served from tiers 1-2.** `boardStatus` is the ✓ Pick
write; `boardPart` / `boardPartLite` / `boardOrder` are drawer lookups that must
be current. They always go live.

**Trust window is 10 minutes** — twice the 5-minute publish trigger, so one
missed run is tolerated but a dead trigger is not.

## Import steps

1. Import the workflow (replacing the existing HQ Board Proxy).
2. Node **2. Read Published Tick** -> Credential -> pick the existing
   **Google Sheets OAuth2** account. It only needs READ on the spreadsheet.
3. Node **3. Serve or Forward** -> replace `__APPS_SCRIPT_EXEC_URL__` and
   `__APP_SECRET_TOKEN__` (same two values as before).
4. **Activate.** If n8n created a SECOND workflow on import, delete the old one
   — two active workflows on the same path is how the board went dark once.

## Verifying which tier answered

The response carries its own provenance:

| field | meaning |
|---|---|
| `_published: true` + `_publishedAgeMs` | tier 2 — the published cell |
| `_cached: true` + `_ageMs` | tier 1 — n8n's static cache |
| `_liveFallback: true` | tier 3 — Apps Script had to compute it |
| none of the above | tier 3 on a non-tick action |

```
curl -s -X POST https://hq.yassinqurabi.com/api/board \
  -H 'Content-Type: application/json' -d '{"action":"boardTick"}' \
  | head -c 400
```

Seeing `_liveFallback: true` repeatedly means tier 2 is not working — check that
the publish trigger is armed and `__Published!A1` has content.
