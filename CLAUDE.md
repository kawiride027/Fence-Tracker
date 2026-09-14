# TGD Fence Tracker

Mobile web app for The Green Dumpster fence drivers to record material movements
(deliveries, pickups, swaps, sales, yard audits) from their phones. Logs to a
Google Sheet via Apps Script and emails the sales rep when there are extras to
invoice or a material sale.

**Live:** https://kawiride027.github.io/Fence-Tracker/ — PIN **9909**

## Stack

Single-file app — `index.html` only. React 18 UMD from cdnjs, **no JSX and no
build step**: the UI is written as plain `h(...)` calls (`var h=React.createElement`).
Inline style objects, no Tailwind. Edit `index.html` directly and push; GitHub
Pages serves it.

Note it mounts with the legacy `ReactDOM.render`, not `createRoot`. Works on
React 18 with a console warning. Left alone deliberately — not worth the churn.

## Backend

Google Sheet **"Fence Inventory"** (`1a-R3sQ_M2Z6q3v198xVPHZen4uMTWPomteb7HCutYwU`)
with a container-bound Apps Script ("Backend v6", 356 lines), deployed as a web
app: **Version 3** (Feb 18 2026), execute as dustin@thegreendumpster.com, access
Anyone. That `/exec` URL is baked into `index.html` as `DEFAULT_SCRIPT_URL`.

Four tabs, created on demand by the script:

| tab | contents |
|---|---|
| `Yard Inventory` | live on-hand dashboard, per item × new/used, plus Out on Rent |
| `Inventory Log` | one row per job line item |
| `Extras Log` | extras needing invoicing |
| `Sales Log` | material sales |

Three POST actions, all to the same URL: `inventory` (every submit),
`email_extras` (extras present), `email_sale` (job type = sale).

Payload contract, verified against the script on 2026-09-14: `jobDate`, `jobType`,
`driver`, `customer`, `jobSite`, `poNumber`, `items`, `swapOut`, `swapIn`,
`extras`, `extraNotes`, `notes`, `paymentReceived`, `paymentType`, `paymentNotes`,
`salesEmail`; line items as `{item, quantity, unit, condition}`.

**`updateYardInventory()` matches items on the exact name string**, multiplication
sign (×) included. Renaming an item in `FENCE_ITEMS` without renaming it in the
script's `ITEMS` array silently stops the yard count from updating. Change both.

## Item model

17 SKUs across Panels / Hardware / Windscreen / Gates / Other. Each carries one
of three tags that drives what the driver is asked:

- `hasBoth` (7) — stocked new *and* used, so quantities are tracked under
  `<id>_new` and `<id>_used` keys and the driver enters both. Panels, ped-gate
  panels, bases.
- `alwaysNew` (8) — posts, clamps, wheels, all four windscreens, sandbags.
- `alwaysUsed` (2) — pedestrian gates, barricades.

Material sales always force `new` regardless of tag.

## How submits report

`postToScript()` tries a normal `fetch` with `Content-Type: text/plain` — Apps
Script accepts that without a CORS preflight, so the response is usually readable
and we can tell the driver whether it landed. If CORS blocks it, it falls back to
a blind `no-cors` post: the row still writes, we just can't confirm. The outcome
drives `syncStatus`, shown on the success screen.

This replaced a fire-and-forget `try{fetch(...)}catch(x){}` that was never awaited,
so a job that never reached the sheet still showed the green success screen.

## Recent Changes

- **2026-09-14** — Baked the deployed `/exec` URL in as `DEFAULT_SCRIPT_URL`
  (it was `""`, so every driver had to paste it by hand or nothing logged), and
  replaced the silent fire-and-forget submit with an awaited one that reports the
  real outcome. Re-authorized the Apps Script and verified the whole path
  end to end.

## History worth knowing

Built as a Claude artifact in Feb 2026; an export of an early version still sits
at `~/Downloads/Other/fence-inventory-app.jsx` (Feb 18) and is **older** than this
repo — it predates the `hasBoth`/`alwaysNew`/`alwaysUsed` work. Don't build from it.

The app appeared unused for months. Two reasons, both found on 2026-09-14: the
script URL was blank by default, and the Apps Script's authorization had lapsed,
which made the endpoint return 403 to everyone. Running `initYardInventory()` once
and granting the scopes fixed the second. The 403 looked exactly like a Workspace
admin policy blocking public access but was not — **check authorization before
concluding an Apps Script web app is blocked by policy.**

## Open items

- [ ] **Enter real starting yard counts** in `Yard Inventory` columns B (New) and
      C (Used) — still all 0, so on-hand numbers are meaningless until seeded.
      Movement logging is unaffected
- [ ] Delete leftover test rows: the "Adrian" rows from Feb 18–19 and the
      "ZZZ TEST - DELETE ME" row from the 2026-09-14 connection test
- [ ] Deployment is pinned at **Version 3**. Editing the script does not go live
      until a new deployment version is pushed
- [ ] No offline queue. A submit with no signal is saved locally and flagged, but
      the driver must redo it. Worth adding retry if drivers hit dead zones
