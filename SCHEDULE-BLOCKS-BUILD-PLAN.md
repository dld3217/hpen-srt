# SRT Schedule Blocks — Build Plan

Flexible per-engagement scheduling + lightweight time accounting. Decided with Dennis/Charlie 2026-09-16.

## Why
Charlie is making SRT his system of record. Real engagements (esp. strategic) run for months with
**multiple visits of the same type**. He needs to schedule — and later account for — his own time:
Remote (with the SE/customer), **Prep** (his own home-office prep), and On-Site, in any number,
contiguous or not. The resulting planned-vs-actual ledger is also the artifact that wins the
managers (Bruno/Watkins): committed vs utilized time per SSE.

## Model — `IScheduleBlock` (src/models/ScheduleBlock.ts)
An engagement (CSERequests row) carries `scheduleBlocks: IScheduleBlock[]`, stored as JSON in a new
`ScheduleBlocks` multi-line text column. Each block:
- `id` — stable, for edit/confirm/log
- `type` — `Remote | Prep | On-Site`
- `start` / `end` — `YYYY-MM-DD` (independent ranges → contiguous-or-not is free); `tbd` when undated
- `location` — On-Site only (Remote/Prep = home office, blank)
- `label` — optional note ("Kickoff visit", "Exec readout prep")
- **Confirmation** (customer-facing only): `proposedBy` ('SE'|'SSE') + `confirmed`. **Prep is self-booked → skips the handshake.**
- **Accounting**: `logged` (Planned→Actual, the one-click switch) + `actualHours` (optional override; when logged & null, actual = planned)

Units: plan in **days** (from the date range), account in **hours** (`HOURS_PER_DAY = 8`, override on log). Reporting:
- committed = Σ plannedHours of non-TBD blocks
- utilized  = Σ actualHours of logged blocks

## Availability
`getSseCommitments` iterates **every non-TBD block** (all types, prep included) as busy → simpler than
today's Remote+On-Site special-casing, and multi-visit double-book protection falls out.

## Phasing (de-risk the scheduling core — the most-iterated part, see date-negotiation v1.0.94–116)
- **Phase 1 (this build):** data model + service; shared block-builder UI on all 3 forms
  (SSE Request, Strategic, New Special Project); dashboard drawer display + inline add/edit + the
  Planned→Actual switch; feed all blocks into availability. Delivers multi-visit planning, one-click
  accounting, and accurate double-booking immediately.
- **Phase 2:** per-block SE↔SSE confirmation; one PA calendar event per confirmed block; retire the
  legacy scalar Remote*/Onsite* columns (kept read-only during Phase 1 for back-compat).

## SP schema (Dennis)
Add one column to **CSERequests**: `ScheduleBlocks` — **Multiple lines of text** (plain).

## Phase 1 step order
1. Model `ScheduleBlock.ts` + `ICseRequest.scheduleBlocks` + service (SP_SELECT/map/create/update) ← START
2. Shared `ScheduleBlockEditor` component (list + add/edit/remove, per-type)
3. Wire into the 3 forms
4. Dashboard drawer: display + inline edit + Planned→Actual toggle
5. Availability: iterate blocks
