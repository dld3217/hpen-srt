// SRT SE Guide — in-app slide-in help panel (mirrors POC Manager's guide pattern).
// Self-contained HTML (inline CSS, no scripts) so it renders in an <iframe srcDoc> and a pop-out window.
export const SE_GUIDE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>SRT — SE Guide</title>
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
</style>
</head>
<body>
  <div class="top-bar">
    <span class="logo">SRT — SE Guide</span>
    <span class="sub">SSE Support Request Tracker</span>
    <span class="badge">HPE Networking</span>
  </div>
  <div class="content">
    <h1>Requesting &amp; tracking SSE support</h1>
    <p class="lead">For Sales Engineers who need a Strategic SSE on a POC or engagement.</p>

    <div class="intro">
      SRT is where you <strong>ask for Strategic SSE help</strong>, agree on <strong>when</strong> it happens,
      and <strong>close the loop</strong> when it's done. You raise a request, the SSE (with their SED) picks it up
      and proposes dates, you confirm, the work happens, and you sign it off. This guide walks that full path.
    </div>

    <h2>1. Raise a request</h2>
    <p>There are two ways in &mdash; both land in the same place:</p>
    <ul>
      <li><strong>From POC Manager</strong> &mdash; on a POC's Customer Info tab, check <strong>&#127919; This POC needs SSE Support</strong> and fill in the panel. This links the request to the POC automatically.</li>
      <li><strong>From the SSE Request form</strong> &mdash; the <strong>+ New SSE Request</strong> button in the SRT header (for engagements that aren't a POC).</li>
    </ul>
    <p>Give the SSE enough to plan around:</p>
    <table>
      <tr><th>Field</th><th>Why it matters</th></tr>
      <tr><td>Primary SSE</td><td>Who you want (auto-fills from your territory's Strategic SSE; you can change it).</td></tr>
      <tr><td>Priority + reason</td><td>High / Medium / Low. The reason ("exec review next week") helps the SED triage.</td></tr>
      <tr><td>Support type</td><td><strong>Remote</strong>, <strong>On-Site</strong>, or <strong>Both</strong> &mdash; drives which schedule fields appear.</td></tr>
      <tr><td>Dates (or TBD)</td><td>Propose dates if you know them, or mark <strong>TBD</strong> and let the SSE propose.</td></tr>
      <tr><td>Solutions focus + description</td><td>What skills are needed so the right SSE is matched.</td></tr>
    </table>
    <div class="tip">When you submit, Outlook opens pre-filled and the request is filed to SRT. The SSE and their SED see it on the board.</div>

    <h2>2. Read the board</h2>
    <p>The dashboard groups every request by where it is in its life. The KPI tiles at the top are click-to-filter:</p>
    <table>
      <tr><th>Section</th><th>Meaning</th></tr>
      <tr><td>Awaiting SED</td><td>New &mdash; the SED hasn't triaged/assigned it yet.</td></tr>
      <tr><td>Needs Info</td><td>The SED or SSE needs more from you before proceeding.</td></tr>
      <tr><td>Awaiting SSE</td><td>Assigned &mdash; waiting on the SSE to accept or propose dates.</td></tr>
      <tr><td>Active</td><td>Dates confirmed, work in flight.</td></tr>
      <tr><td>Needs Sign-off</td><td>SSE marked it complete &mdash; <strong>waiting on you to verify</strong> (see step 4).</td></tr>
      <tr><td>Completed</td><td>Verified and closed.</td></tr>
      <tr><td>Parked</td><td>On hold; frees the SSE's calendar until it resumes.</td></tr>
    </table>

    <h2>3. The date handshake (confirm or counter)</h2>
    <p>Scheduling is a two-way handshake so nobody is double-booked:</p>
    <div class="step"><div class="step-n">1</div><div>The SSE (or SED) proposes on-site / remote dates. Your request shows a <strong>&#128197; Dates proposed</strong> badge, and in POC Manager you'll see a yellow "awaiting confirmation" banner.</div></div>
    <div class="step"><div class="step-n">2</div><div>Open the request. Click <strong>Confirm</strong> to accept the dates &mdash; they're locked in and (for a POC) mirrored onto the POC's dates automatically.</div></div>
    <div class="step"><div class="step-n">3</div><div>Or click <strong>Propose different</strong> to counter with new dates &mdash; the turn bounces back to the SSE. This ping-pongs until you both agree.</div></div>
    <div class="note">Confirmed dates block that time on the SSE's availability calendar, so other SEs can see they're committed.</div>

    <h2>4. Close it out &mdash; Verify Complete</h2>
    <p>Closing is deliberately two-stage so "done" is a mutual handshake, not a one-sided flip:</p>
    <div class="step"><div class="step-n">1</div><div>When the work is finished, the SSE marks it <strong>Complete</strong>. It moves to the amber <strong>&#128390;&#65039; Needs Sign-off &mdash; Awaiting SE</strong> section.</div></div>
    <div class="step"><div class="step-n">2</div><div>You (the Primary SE) click <strong>&#10003; Verify Complete</strong>. That's the handshake &mdash; the request moves to <strong>&#10003; Completed</strong> and counts as truly done.</div></div>
    <div class="tip">Until you verify, it sits in Needs Sign-off. If a request is done but lingering there, it's waiting on your click.</div>

    <h2>Tips</h2>
    <ul>
      <li>Everyone can <em>see</em> the whole board (read-only); actions are limited to the people on each request.</li>
      <li>Set <strong>Priority</strong> honestly &mdash; it's how SEDs sequence limited SSE capacity.</li>
      <li>Raising the request from the POC keeps everything linked; you'll get the schedule updates back on the POC.</li>
    </ul>
  </div>
</body>
</html>`;
