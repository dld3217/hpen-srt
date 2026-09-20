/* ============================================================================
 * SRT one-time cleanup: grandfather historical Completed rows into true Completed
 * ----------------------------------------------------------------------------
 * For every RequestStatus='Complete' row that has NO sign-off yet, stamp
 * SignedOffBy (the SE's name, parsed from SEPrimary; else "Historical") +
 * SignOffDate = now. That moves the historical backlog straight into the true
 * "Completed" section, so the SE-verify handshake only governs NEW completions.
 *
 * HOW TO RUN
 *   1. Open the site:  https://hpe.sharepoint.com/teams/hpen-poc-manager  -> F12 -> Console.
 *   2. Paste this whole file, Enter. DRY RUN: prints what it WOULD stamp. Nothing written.
 *   3. Review. Set  APPLY = false -> true , paste again to write.
 *
 * Options:
 *   STAMP = 'Historical'  -> use a fixed label instead of the parsed SE name.
 * ========================================================================== */
(async () => {
  const SITE  = 'https://hpe.sharepoint.com/teams/hpen-poc-manager';
  const APPLY = false;          // <-- leave false to preview; set true to write
  const STAMP = '';             // <-- '' = use the SE's name (fallback "Historical"); or set e.g. 'Historical'

  const LIST   = "getbytitle('CSERequests')";
  const H      = (extra = {}) => ({ Accept: 'application/json;odata=nometadata', ...extra });
  const unwrap = v => {
    if (!v) return '';
    if (typeof v === 'object' && 'Value' in v) return v.Value || '';
    if (typeof v === 'string' && v.charAt(0) === '{') { try { return JSON.parse(v).Value || ''; } catch (e) { /* fall through */ } }
    return String(v);
  };
  // Parse "Name / email" or AD "Last, First (title)" -> "First Last".
  const seName = raw => {
    let n = (raw || '').split('/')[0].replace(/\s*\(.*\)\s*$/, '').trim();
    if (n.indexOf(',') !== -1) { const p = n.split(','); n = `${(p[1] || '').trim()} ${(p[0] || '').trim()}`.trim(); }
    return n;
  };

  // 1) Page through all CSERequests.
  const items = [];
  let url = `${SITE}/_api/web/lists/${LIST}/items?$select=Id,Title,CustomerName,RequestStatus,SignedOffBy,SignOffDate,SEPrimary&$top=2000`;
  while (url) {
    const page = await fetch(url, { headers: H() }).then(r => r.json());
    items.push(...(page.value || []));
    url = page['odata.nextLink'] || page['@odata.nextLink'] || null;
  }

  // 2) Completed rows with NO sign-off yet.
  const targets = items
    .filter(it => unwrap(it.RequestStatus) === 'Complete' && !it.SignedOffBy && !it.SignOffDate)
    .map(it => ({ id: it.Id, customer: it.CustomerName || it.Title, by: STAMP || seName(it.SEPrimary) || 'Historical' }));

  console.log(`Scanned ${items.length} rows — ${targets.length} Completed row(s) to grandfather:`);
  console.table(targets.map(t => ({ Id: t.id, Customer: t.customer, '-> SignedOffBy': t.by })));

  if (!targets.length) { console.log('Nothing to sign off.'); return; }
  if (!APPLY) { console.log('%cDRY RUN — no changes written. Set APPLY = true to apply.', 'color:#b45309;font-weight:bold'); return; }

  // 3) Apply. MERGE-update each row.
  const now    = new Date().toISOString();
  const digest = document.getElementById('__REQUESTDIGEST')?.value
    || (await fetch(`${SITE}/_api/contextinfo`, { method: 'POST', headers: H() }).then(r => r.json())).FormDigestValue;

  let ok = 0, fail = 0;
  for (const t of targets) {
    try {
      const res = await fetch(`${SITE}/_api/web/lists/${LIST}/items(${t.id})`, {
        method: 'POST',
        headers: H({ 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
        body: JSON.stringify({ SignedOffBy: t.by, SignOffDate: now }),
      });
      if (res.ok) { ok++; console.log(`✓ ${t.id} ${t.customer} -> ${t.by}`); }
      else { fail++; console.error(`✗ ${t.id} (${res.status})`, await res.text()); }
    } catch (e) { fail++; console.error(`✗ ${t.id}`, e); }
  }
  console.log(`%cDone. Signed off ${ok}, failed ${fail}.`, 'color:#107c10;font-weight:bold');
})();
