// SRT SSE Guide — in-app slide-in help panel (mirrors POC Manager's guide pattern).
// Self-contained HTML (inline CSS, no scripts) so it renders in an <iframe srcDoc> and a pop-out window.
export const SSE_GUIDE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SRT — SSE Guide</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; background:#f0f2f5; color:#1a2332; font-size:14px; line-height:1.6; }
  .top-bar { background:#01a982; padding:14px 24px; display:flex; align-items:center; gap:12px; }
  .top-bar .logo { font-size:17px; font-weight:700; color:#fff; }
  .top-bar .sub { font-size:12px; color:rgba(255,255,255,0.85); margin-left:6px; }
  .top-bar .badge { margin-left:auto; font-size:11px; color:#fff; background:rgba(0,0,0,0.18); padding:4px 10px; border-radius:3px; font-weight:600; }
  .content { padding:28px 36px 80px; max-width:860px; margin:0 auto; }
  h1 { font-size:22px; font-weight:700; margin-bottom:4px; }
  .lead { color:#5a6b7a; font-size:13px; margin-bottom:18px; }
  h2 { font-size:16px; font-weight:700; margin:34px 0 10px; padding-bottom:8px; border-bottom:2px solid #01a982; }
  h3 { font-size:13px; font-weight:700; margin:18px 0 6px; }
  p { margin-bottom:10px; color:#2a3a4a; font-size:13px; }
  ul,ol { margin:0 0 12px 20px; color:#2a3a4a; font-size:13px; }
  li { margin-bottom:5px; }
  code { font-family:'Courier New',monospace; font-size:11px; background:#eef1f4; border:1px solid #d1d9e0; border-radius:3px; padding:1px 5px; }
  .intro { background:#1a2332; color:#c3d0dc; border-radius:6px; padding:14px 18px; margin-bottom:26px; font-size:13px; line-height:1.7; }
  .intro strong { color:#01a982; }
  .note { background:#fff8e0; border-left:4px solid #e0c000; border-radius:4px; padding:10px 14px; margin:12px 0; font-size:12.5px; }
  .tip { background:#e8faf3; border-left:4px solid #01a982; border-radius:4px; padding:10px 14px; margin:12px 0; font-size:12.5px; }
  .step { display:flex; gap:12px; margin-bottom:10px; align-items:flex-start; }
  .step-n { min-width:24px; height:24px; background:#01a982; color:#fff; border-radius:50%; font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; flex-shrink:0; margin-top:1px; }
  table { border-collapse:collapse; width:100%; margin:10px 0 18px; font-size:12.5px; }
  th,td { border:1px solid #d1d9e0; padding:6px 10px; text-align:left; vertical-align:top; }
  th { background:#1a2332; color:#fff; font-weight:600; }
  tr:nth-child(even) td { background:#f7f9fb; }
  .legend span { display:inline-block; font-size:11px; font-weight:700; padding:2px 9px; border-radius:10px; margin-right:6px; }
</style>
</head>
<body>
  <div class="top-bar">
    <span class="logo">SRT — SSE Guide</span>
    <span class="sub">SSE Support Request Tracker</span>
    <span class="badge">HPE Networking</span>
  </div>
  <div class="content">
    <h1>Managing your queue &amp; availability</h1>
    <p class="lead">For Strategic SSEs being requested onto POCs and engagements.</p>

    <div class="intro">
      SRT is your <strong>queue and calendar</strong>. Requests for your help land here; you <strong>schedule</strong> them,
      track <strong>planned vs actual</strong> time, keep your <strong>availability</strong> honest so nobody double-books you,
      and hand work back to the SE to close. This guide covers all of it.
    </div>

    <h2>1. Pick up a request</h2>
    <p>New requests assigned to you appear under <strong>Awaiting SSE</strong>. Open one and either:</p>
    <ul>
      <li><strong>Accept a TBD</strong> &mdash; if the SE left dates open, propose your dates.</li>
      <li><strong>Propose dates</strong> &mdash; offer on-site / remote dates. The SE gets the confirm turn (see the handshake below).</li>
    </ul>

    <h2>2. The date handshake</h2>
    <p>Dates aren't final until both sides agree:</p>
    <div class="step"><div class="step-n">1</div><div>You propose dates &mdash; the request shows <strong>Dates proposed</strong> and the turn is the SE's.</div></div>
    <div class="step"><div class="step-n">2</div><div>The SE <strong>Confirms</strong> (locked in) or <strong>Proposes different</strong> (bounces back to you). Repeat until agreed.</div></div>
    <div class="note">A special-project or SED-driven request may need both the SED and you to accept before it goes firm &mdash; until then it shows as <strong>&#9906; Tentative / on&#8209;hold</strong> (amber) so another SE won't book over it.</div>

    <h2>3. Schedule blocks — Planned &rarr; Actual</h2>
    <p>Each engagement is built from <strong>schedule blocks</strong> so a 2-hour planning call isn't tracked like a full day on site:</p>
    <ul>
      <li>Add any number of <strong>Remote</strong>, <strong>Prep</strong>, or <strong>On-Site</strong> blocks.</li>
      <li>Each block toggles between <strong>Days</strong> and <strong>Hours</strong> &mdash; capture effort at the right granularity.</li>
      <li>Book it as <strong>Planned</strong>; when it happens, one click flips it to <strong>Actual</strong> so committed-vs-utilized time is real.</li>
      <li>The editor flags <strong>conflicts / free-busy</strong> against your existing blocks as you add them.</li>
    </ul>

    <h2>4. Your availability calendar &amp; the team heatmap</h2>
    <p>"Availability at a Glance" shows <strong>three months side by side</strong>. Colors reflect your firm commitments:</p>
    <p class="legend">
      <span style="background:#e8faf3;color:#107c10;">Green = free</span>
      <span style="background:#fff4ce;color:#8a6000;">Yellow = partial</span>
      <span style="background:#fde7e9;color:#a4262c;">Red = full</span>
      <span style="background:#f3ecff;color:#6b2faf;">Amber-hatch = tentative / on-hold</span>
    </p>
    <ul>
      <li>Hover a booked day to see the <strong>customer / engagement</strong>; click it to jump straight to that request on the board.</li>
      <li>The <strong>&#128293; Team Availability heatmap</strong> shows every SSE across ~13 weeks &mdash; useful when load needs to be shared.</li>
      <li><strong>Firm</strong> time (Scheduled / In&nbsp;Progress) counts as busy; <strong>Parked / Declined / Cancelled / Complete</strong> free the calendar back up.</li>
    </ul>

    <h2>5. Block your personal time — &#128197; Schedule Myself</h2>
    <p>Use the <strong>&#128197; Schedule Myself</strong> button (top of the dashboard) to block <strong>PTO, holidays, or appointments</strong>:</p>
    <div class="step"><div class="step-n">1</div><div>Pick the type and dates. It's saved as a hidden <em>Personal</em> entry &mdash; it blocks your availability but never shows on the board or in Insights.</div></div>
    <div class="step"><div class="step-n">2</div><div>To change or cancel an appointment, reopen <strong>Schedule Myself</strong> &mdash; it lists your upcoming personal entries with an <strong>&#10005; Remove</strong> on each (the only place to manage them, since they're hidden from the board).</div></div>

    <h2>6. Finish &amp; hand off</h2>
    <div class="step"><div class="step-n">1</div><div>When the work is done, mark the request <strong>Complete</strong>. It moves to <strong>Needs Sign-off</strong> and waits on the SE.</div></div>
    <div class="step"><div class="step-n">2</div><div>The SE clicks <strong>&#10003; Verify Complete</strong> to close it. (That handshake is theirs &mdash; it's how "done" gets confirmed on both sides.)</div></div>
    <div class="tip">Need to pause an engagement? <strong>Park</strong> it &mdash; that frees your calendar for those dates until it resumes.</div>

    <h3>Demo data (admins &amp; SSEs)</h3>
    <p>The <strong>&#129514; Load Demo</strong> control pre-loads sample engagements across a 3-month calendar so you can practice or demo the full availability picture; <strong>&#129529; Clear [SAMPLE]</strong> backs them out.</p>
  </div>
</body>
</html>`;
