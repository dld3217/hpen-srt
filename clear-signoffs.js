/* ============================================================================
 * SRT one-time cleanup: clear the Sign-off flag on all COMPLETED requests
 * ----------------------------------------------------------------------------
 * Wipes SignedOffBy + SignOffDate on every RequestStatus='Complete' row so the
 * slate is clean and going forward every completion runs through the SE-verified
 * sign-off flow. NOTE: a cleared Complete row satisfies "Complete && no sign-off"
 * = Needs Sign-off, so these WILL re-surface in the Needs Sign-off tile/section
 * for the SE to verify fresh. That is the intent.
 *
 * HOW TO RUN
 *   1. Open the site in the browser:  https://hpe.sharepoint.com/teams/hpen-poc-manager
 *   2. F12 -> Console tab.
 *   3. Paste this whole file, press Enter.  It runs a DRY RUN and prints a table
 *      of every row it WOULD clear (flagging demo [SAMPLE] vs real). Nothing is written.
 *   4. Review the table. When satisfied, change  APPLY = false  to  APPLY = true
 *      below, paste again, and it will write the changes.
 *
 * Options:
 *   SAMPLE_ONLY = true   -> only clear demo [SAMPLE] rows (leave real sign-offs intact)
 *
 * Safe: dry-run first, per-row error logging, only touches Completed rows that
 * currently HAVE a sign-off stamp.
 * ========================================================================== */
(async () => {
  const SITE        = 'https://hpe.sharepoint.com/teams/hpen-poc-manager';
  const APPLY       = false;   // <-- leave false to preview; set true to write changes
  const SAMPLE_ONLY = false;   // <-- set true to clear ONLY demo [SAMPLE] rows

  const LIST   = "getbytitle('CSERequests')";
  const H      = (extra = {}) => ({ Accept: 'application/json;odata=nometadata', ...extra });
  // Choice fields can read back as a plain string, a { Value } object, or a JSON string.
  const unwrap = v => {
    if (!v) return '';
    if (typeof v === 'object' && 'Value' in v) return v.Value || '';
    if (typeof v === 'string' && v.charAt(0) === '{') { try { return JSON.parse(v).Value || ''; } catch (e) { /* fall through */ } }
    return String(v);
  };
  const isSample = it => /^\s*\[sample\]/i.test(it.CustomerName || '') || /^\s*\[sample\]/i.test(it.Title || '');

  // 1) Page through all CSERequests.
  const items = [];
  let url = `${SITE}/_api/web/lists/${LIST}/items?$select=Id,Title,CustomerName,RequestStatus,SignedOffBy,SignOffDate&$top=2000`;
  while (url) {
    const page = await fetch(url, { headers: H() }).then(r => r.json());
    items.push(...(page.value || []));
    url = page['odata.nextLink'] || page['@odata.nextLink'] || null;
  }

  // 2) Completed rows that currently carry a sign-off stamp.
  let targets = items.filter(it => unwrap(it.RequestStatus) === 'Complete' && (it.SignedOffBy || it.SignOffDate));
  if (SAMPLE_ONLY) targets = targets.filter(isSample);

  const sampleCount = targets.filter(isSample).length;
  console.log(`Scanned ${items.length} rows — ${targets.length} Completed row(s) with a sign-off to clear (${sampleCount} demo [SAMPLE], ${targets.length - sampleCount} real):`);
  console.table(targets.map(t => ({
    Id: t.Id, Customer: t.CustomerName || t.Title, Demo: isSample(t) ? '[SAMPLE]' : '',
    'SignedOffBy (clearing)': t.SignedOffBy || '(blank)', 'SignOffDate (clearing)': t.SignOffDate || '(blank)',
  })));

  if (!targets.length) { console.log('Nothing to clear.'); return; }
  if (!APPLY) { console.log('%cDRY RUN — no changes written. Set APPLY = true to apply.', 'color:#b45309;font-weight:bold'); return; }

  // 3) Apply. MERGE-update each row, clearing both sign-off fields.
  const digest = document.getElementById('__REQUESTDIGEST')?.value
    || (await fetch(`${SITE}/_api/contextinfo`, { method: 'POST', headers: H() }).then(r => r.json())).FormDigestValue;

  let ok = 0, fail = 0;
  for (const t of targets) {
    try {
      const res = await fetch(`${SITE}/_api/web/lists/${LIST}/items(${t.Id})`, {
        method: 'POST',
        headers: H({ 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
        body: JSON.stringify({ SignedOffBy: null, SignOffDate: null }),
      });
      if (res.ok) { ok++; console.log(`✓ ${t.Id} ${t.CustomerName || t.Title}`); }
      else { fail++; console.error(`✗ ${t.Id} (${res.status})`, await res.text()); }
    } catch (e) { fail++; console.error(`✗ ${t.Id}`, e); }
  }
  console.log(`%cDone. Cleared ${ok}, failed ${fail}.`, 'color:#107c10;font-weight:bold');
})();
