# SRT — SSE Support Request Tracker (SPFx) — Claude Instructions

Standalone, Dennis-owned SPFx solution for tracking SSE strategic engagements / support requests.
**Multi web-part:** srtDashboard, srtInsights, sseAvailability, sseRequestForm, strategicEngagement.
Org model: Charlie Clemmer / Mike Bruno / Rick Watkins. Deployed to the hpen-poc-manager site catalog.

## Stack & environment
- SPFx 1.20 · React 17 · TypeScript · Fluent UI v8 · PnPjs v4
- Node **18.20.8** (nvm): `export PATH="/c/Users/doddd/AppData/Local/nvm/v18.20.8:$PATH"` — or `nvm use 18`
- GitHub (Dennis's): https://github.com/dld3217/hpen-srt — **work direct on `master`, no branches** (Dennis's preference).
- Deploy target: hpen-poc-manager site App Catalog.

## Build & deploy (ship) — ALWAYS clean first
There's an npm script that does the whole clean+ship:
```
npm run ship          # = gulp clean && gulp bundle --ship && gulp package-solution --ship
```
(or run those three gulp steps manually with `npx`).
Output `sharepoint/solution/hpen-srt.sppkg` → upload to **Apps for SharePoint** → **Replace** →
**Deploy** → Ctrl+F5. Deploy the .sppkg freely (dev-stage cadence).

## Version bumping — bump BOTH, every build (not linked)
1. `src/appVersion.ts` → `APP_VERSION`  (most web parts import this)
2. `config/package-solution.json` → `solution.version`
- ⚠️ Known drift: `src/webparts/sseRequestForm/components/SseRequestForm.tsx` has its OWN hardcoded
  `const VERSION = '1.0.63'` that does NOT use `APP_VERSION`. If that web part's badge should match,
  update it too — better, switch it to import `APP_VERSION`.

## Pages
Dashboard / Insights / Availability render as app pages (some as single-part app pages created via
console REST). `SSE-Availability.aspx` exists.

## SharePoint / Power Automate notes
- **PnP.PowerShell blocked** on the HPE tenant → browser-console REST for data ops.
- PA Create/Update to a **multi-choice Choice** column needs an ARRAY of `{Value}` OBJECTS, not bare
  strings (a generic 400 otherwise) — reshape with a Select action first.
- Same SP gotchas as POC Manager (encoded internal names, `ExternalClass` rich text, fixed-panel close buttons).

## Where "current state" lives (this file is conventions only)
Live state — current version, Clemmer write-back, SSE nudge, deploy status — is in the auto-memory:
`~/.claude/projects/C--Users-doddd/memory/project_srt.md` and `reminder_active.md`.
