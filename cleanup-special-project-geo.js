/* ============================================================================
 * SRT one-time cleanup: retag mis-filed Special-Project rows
 * ----------------------------------------------------------------------------
 * Charlie's spreadsheet import put CIC / Marketing program names into the geo
 * fields (HPENBusinessUnit / BURegion) of ordinary rows. This finds those rows
 * and reclassifies them as real Special Projects:
 *     EngagementType        -> 'Special Project'
 *     SpecialProjectCategory-> the matched category (e.g. 'Houston CIC')
 *     SpecialProjectInitiative-> matched initiative, if the dirty value was one
 *     HPENBusinessUnit      -> cleared
 *     BURegion              -> cleared
 *
 * HOW TO RUN
 *   1. Open the site in the browser:  https://hpe.sharepoint.com/teams/hpen-poc-manager
 *   2. F12 -> Console tab.
 *   3. Paste this whole file, press Enter.  It runs a DRY RUN and prints a table
 *      of every row it WOULD change. Nothing is written.
 *   4. Review the table. When satisfied, change  APPLY = false  to  APPLY = true
 *      below, paste again, and it will write the changes.
 *
 * Safe: dry-run first, per-row error logging, and it only touches rows whose
 * geo field matches a known Special-Project category/initiative name.
 * ========================================================================== */
(async () => {
  const SITE  = 'https://hpe.sharepoint.com/teams/hpen-poc-manager';
  const APPLY = false;   // <-- leave false to preview; set true to write changes

  const LIST  = "getbytitle('CSERequests')";
  const norm  = s => (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
  const H     = (extra = {}) => ({ Accept: 'application/json;odata=nometadata', ...extra });
  // Choice fields can read back three ways: a plain string, a wrapped { Value } object, OR a
  // STRING containing SPListExpandedReference JSON (Power Automate writes this). Handle all three.
  const unwrap = v => {
    if (!v) return '';
    if (typeof v === 'object' && 'Value' in v) return v.Value || '';
    if (typeof v === 'string' && v.charAt(0) === '{') { try { return JSON.parse(v).Value || ''; } catch (e) { /* fall through */ } }
    return String(v);
  };

  // 1) Load the Special-Project taxonomy from AppConfig (fall back to the app defaults).
  let spMap = {
    'Houston CIC':  ['Bee Counting', 'Hardware Demo'],
    'San Jose CIC': [],
    'Marketing':    ['HPE Marketing Days', 'Discover', 'Sales Kick Off', 'TechJam'],
  };
  try {
    const cfg = await fetch(
      `${SITE}/_api/web/lists/getbytitle('AppConfig')/items?$filter=Title eq 'SRTSpecialProjects'&$select=Value`,
      { headers: H() }).then(r => r.json());
    if (cfg.value && cfg.value[0] && cfg.value[0].Value) spMap = JSON.parse(cfg.value[0].Value);
  } catch (e) { console.warn('Could not read SRTSpecialProjects config — using defaults.', e); }

  // Build normKey -> { category, initiative } lookup (categories AND their initiatives).
  const lookup = {};
  for (const [cat, inits] of Object.entries(spMap)) {
    if (cat) lookup[norm(cat)] = { category: cat, initiative: '' };
    (inits || []).forEach(i => { if (i && !lookup[norm(i)]) lookup[norm(i)] = { category: cat, initiative: i }; });
  }
  console.log('Special-Project names to match:', Object.keys(lookup));

  // 2) Page through all CSERequests.
  const items = [];
  let url = `${SITE}/_api/web/lists/${LIST}/items?$select=Id,Title,HPENBusinessUnit,BURegion,EngagementType,SpecialProjectCategory,SpecialProjectInitiative&$top=2000`;
  while (url) {
    const page = await fetch(url, { headers: H() }).then(r => r.json());
    items.push(...(page.value || []));
    url = page['odata.nextLink'] || page['@odata.nextLink'] || null;
  }

  // 3) Find rows whose BU or Region field matches a Special-Project name.
  const dirty = [], review = [];
  const looksSpecial = s => /cic|innovation|market/.test(norm(s));
  for (const it of items) {
    const bu = unwrap(it.HPENBusinessUnit);
    const rg = unwrap(it.BURegion);
    const hit = lookup[norm(bu)] || lookup[norm(rg)];
    if (hit) dirty.push({ id: it.Id, title: it.Title, bu, rg, hit });
    else if (looksSpecial(bu) || looksSpecial(rg)) review.push({ Id: it.Id, Title: it.Title, BU: bu, Region: rg });
  }

  if (review.length) {
    console.log(`%c${review.length} row(s) look special but matched NO config name — assign these manually or tell me:`, 'color:#b45309;font-weight:bold');
    console.table(review);
  }
  console.log(`Scanned ${items.length} rows — ${dirty.length} to retag:`);
  console.table(dirty.map(d => ({
    Id: d.id, Title: d.title, 'BU (old)': d.bu, 'Region (old)': d.rg,
    '-> Category': d.hit.category, '-> Initiative': d.hit.initiative || '(none)',
  })));

  if (!dirty.length) { console.log('Nothing to clean up.'); return; }
  if (!APPLY) { console.log('%cDRY RUN — no changes written. Set APPLY = true to apply.', 'color:#b45309;font-weight:bold'); return; }

  // 4) Apply. MERGE-update each row.
  const digest = document.getElementById('__REQUESTDIGEST')?.value
    || (await fetch(`${SITE}/_api/contextinfo`, { method: 'POST', headers: H() }).then(r => r.json())).FormDigestValue;

  let ok = 0, fail = 0;
  for (const d of dirty) {
    const body = {
      EngagementType: 'Special Project',
      SpecialProjectCategory: d.hit.category,
      HPENBusinessUnit: null,   // clear the geo fields — Special Projects are non-geo
      BURegion: null,
    };
    if (d.hit.initiative) body.SpecialProjectInitiative = d.hit.initiative;
    try {
      const res = await fetch(`${SITE}/_api/web/lists/${LIST}/items(${d.id})`, {
        method: 'POST',
        headers: H({ 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
        body: JSON.stringify(body),
      });
      if (res.ok) { ok++; console.log(`✓ ${d.id} ${d.title} -> ${d.hit.category}`); }
      else { fail++; console.error(`✗ ${d.id} (${res.status})`, await res.text()); }
    } catch (e) { fail++; console.error(`✗ ${d.id}`, e); }
  }
  console.log(`%cDone. Updated ${ok}, failed ${fail}.`, 'color:#107c10;font-weight:bold');
})();
