// SRT SED / Admin Guide — in-app slide-in help panel (mirrors POC Manager's guide pattern).
// Self-contained HTML (inline CSS, no scripts) so it renders in an <iframe srcDoc> and a pop-out window.
export const SED_GUIDE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SRT — SED / Admin Guide</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; background:#f0f2f5; color:#1a2332; font-size:14px; line-height:1.6; }
  .top-bar { background:#1a2332; padding:14px 24px; display:flex; align-items:center; gap:12px; }
  .top-bar .logo { font-size:17px; font-weight:700; color:#fff; }
  .top-bar .sub { font-size:12px; color:#8fb9ac; margin-left:6px; }
  .top-bar .badge { margin-left:auto; font-size:11px; color:#fff; background:#01a982; padding:4px 10px; border-radius:3px; font-weight:600; }
  .content { padding:28px 36px 80px; max-width:860px; margin:0 auto; }
  h1 { font-size:22px; font-weight:700; margin-bottom:4px; }
  .lead { color:#5a6b7a; font-size:13px; margin-bottom:18px; }
  h2 { font-size:16px; font-weight:700; margin:34px 0 10px; padding-bottom:8px; border-bottom:2px solid #1a2332; }
  h3 { font-size:13px; font-weight:700; margin:18px 0 6px; }
  p { margin-bottom:10px; color:#2a3a4a; font-size:13px; }
  ul,ol { margin:0 0 12px 20px; color:#2a3a4a; font-size:13px; }
  li { margin-bottom:5px; }
  code { font-family:'Courier New',monospace; font-size:11px; background:#eef1f4; border:1px solid #d1d9e0; border-radius:3px; padding:1px 5px; }
  .intro { background:#1a2332; color:#c3d0dc; border-radius:6px; padding:14px 18px; margin-bottom:26px; font-size:13px; line-height:1.7; }
  .intro strong { color:#01a982; }
  .note { background:#fff8e0; border-left:4px solid #e0c000; border-radius:4px; padding:10px 14px; margin:12px 0; font-size:12.5px; }
  .tip { background:#e8faf3; border-left:4px solid #01a982; border-radius:4px; padding:10px 14px; margin:12px 0; font-size:12.5px; }
  .warn { background:#fde7e9; border-left:4px solid #a4262c; border-radius:4px; padding:10px 14px; margin:12px 0; font-size:12.5px; }
  .step { display:flex; gap:12px; margin-bottom:10px; align-items:flex-start; }
  .step-n { min-width:24px; height:24px; background:#1a2332; color:#fff; border-radius:50%; font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; flex-shrink:0; margin-top:1px; }
  table { border-collapse:collapse; width:100%; margin:10px 0 18px; font-size:12.5px; }
  th,td { border:1px solid #d1d9e0; padding:6px 10px; text-align:left; vertical-align:top; }
  th { background:#1a2332; color:#fff; font-weight:600; }
  tr:nth-child(even) td { background:#f7f9fb; }
</style>
</head>
<body>
  <div class="top-bar">
    <span class="logo">SRT — SED / Admin Guide</span>
    <span class="sub">SSE Support Request Tracker</span>
    <span class="badge">Leaders &amp; Admins</span>
  </div>
  <div class="content">
    <h1>Triage, sign-off, demand &amp; capacity</h1>
    <p class="lead">For SE Directors and administrators managing SSE demand across the team.</p>

    <div class="intro">
      As a leader you own the <strong>front and back of the queue</strong>: you triage incoming demand, make sure the
      right SSE is assigned, and <strong>sign off</strong> completed work. SRT also gives you the <strong>demand &amp; capacity</strong>
      picture &mdash; who's asking, who's committed, and where there's room.
    </div>

    <h2>1. Triage incoming demand</h2>
    <p>New requests land under <strong>Awaiting SED</strong>. For each:</p>
    <ul>
      <li>Assign / confirm the <strong>SSE</strong>, or send it back with <strong>Needs Info</strong> if it's underspecified.</li>
      <li>You can <strong>propose dates</strong> on the SSE's behalf &mdash; the SE then confirms or counters (the standard handshake).</li>
      <li><strong>Solutions</strong> on a request are click-to-edit for admins if they were miscoded.</li>
    </ul>
    <div class="tip">Everyone sees the board read-only; approve/assign/date actions and edits are gated to SEDs and admins (and the assigned SSE for their own rows).</div>

    <h2>2. The two-stage close &amp; sign-off</h2>
    <p>"Done" is a deliberate handshake, not a single flip:</p>
    <div class="step"><div class="step-n">1</div><div>SSE marks the work <strong>Complete</strong> &rarr; it moves to <strong>&#128390;&#65039; Needs Sign-off &mdash; Awaiting SE</strong> (amber).</div></div>
    <div class="step"><div class="step-n">2</div><div>The Primary SE clicks <strong>&#10003; Verify Complete</strong> &rarr; it moves to <strong>&#10003; Completed</strong> (green). The <strong>Complete</strong> KPI tile counts only signed-off work.</div></div>
    <div class="step"><div class="step-n">3</div><div>As an admin you can <strong>Sign Off</strong> a row directly if the SE is unavailable.</div></div>
    <div class="note">Historical completes that predate the two-stage flow were grandfathered into Completed with a <code>SignedOffBy = Historical</code> stamp, so the Needs Sign-off queue reflects only real, current work.</div>

    <h2>3. SSE Demand Insights</h2>
    <p>The <strong>&#128202; SSE Demand Insights</strong> link opens the analytics view &mdash; demand by SSE, request mix, and
       <strong>committed vs utilized hours</strong> (from the Planned&rarr;Actual schedule blocks). Use it to see where demand is
       concentrated and whether capacity is keeping up. Personal / PTO time is excluded from these numbers.</p>

    <h2>4. Availability &amp; team capacity</h2>
    <p>The <strong>&#128197; SSE Availability</strong> view and the <strong>&#128293; Team Availability heatmap</strong> (all SSEs &times; ~13 weeks)
       are your capacity-planning tools &mdash; who's free, who's slammed, and where to steer the next request. Firm commitments show solid;
       tentative / on-hold requests show amber-hatched so you can see pending load too.</p>

    <h2>5. Admin panel</h2>
    <p>Admins get an <strong>Admin</strong> button (header) opening the admin panel to manage:</p>
    <ul>
      <li><strong>SEDs</strong> and <strong>Strategic SSEs</strong> membership (who has which role).</li>
      <li>SSEs are also auto-detected from the <strong>Contact Directory</strong> (category "Strategic Systems Engineers").</li>
      <li>Canonical BU / region values &mdash; keep these aligned to the gold-standard nomenclature shared with POC Manager.</li>
    </ul>

    <h2>6. Special Projects &amp; demo data</h2>
    <ul>
      <li><strong>&#11088; New Special Project</strong> creates an engagement that starts Active (no SE request needed) &mdash; for proactive / internal work.</li>
      <li><strong>&#129514; Load Demo</strong> pre-loads a realistic 3-month sample calendar for walk-throughs; <strong>&#129529; Clear [SAMPLE]</strong> removes it.</li>
    </ul>

    <h3>View-as (admins)</h3>
    <p>Admins can <strong>view the board as another user</strong> to check what an SE or SSE sees. Note that <em>writes still happen as your real identity unless</em> the action explicitly carries the viewed-as user (e.g. Schedule Myself books for the viewed-as SSE) &mdash; so use view-as for verification, and double-check whose name a booking lands under.</p>
    <div class="warn">Role spoofing is admin-only and server-checked &mdash; a non-admin can't unlock leader actions from the browser console.</div>
  </div>
</body>
</html>`;
