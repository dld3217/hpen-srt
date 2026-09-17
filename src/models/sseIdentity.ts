// Reconcile the same SSE across the different spellings that land in `requestedCse` /
// commitment records: AD "Last, First (title)" with a drifting title suffix, a "First Last"
// hand-typed variant, some with an email after "/", some without. Used to key aggregations
// (Insights SSE Workload, SSE Availability roster) so one person is one row.

// Clean an SSE display name: strip a trailing "(title…)" and convert "Last, First" → "First Last".
export const parseSseDisplay = (raw: string): string => {
  let name = (raw || '').replace(/\s*\(.*\)\s*$/, '').trim();
  if (name.indexOf(',') !== -1) { const p = name.split(','); name = `${p[1].trim()} ${p[0].trim()}`; }
  return name;
};

// Fold a (already display-cleaned) name to a comparison key: lowercase, strip non-alphanumerics.
export const sseNameKey = (name: string): string => (name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
