import * as React from 'react';
import { useState, useEffect } from 'react';
import { APP_VERSION } from '../../../appVersion';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
import { ContactDirectoryService, IContact, GENERALIST_CATEGORY } from '../../../services/ContactDirectoryService';
import { ISseCommitment } from '../../../models/ICseRequest';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISseAvailabilityProps {
  sp: SPFI;
  context: WebPartContext;
}

const VERSION = APP_VERSION;
const SRT_DASHBOARD_URL = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SRT-Resource-Dashboard.aspx';

const dstr = (iso: string): string => (iso || '').substring(0, 10);
const fmtD = (iso: string): string => {
  const p = dstr(iso); if (!p) return '';
  const [y, m, d] = p.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
};
const fmtRange = (a: string, b: string): string => {
  const s = fmtD(a), e = fmtD(b);
  return e && e !== s ? `${s} – ${e}` : s;
};

const CARD: React.CSSProperties = { background: '#fff', border: '1px solid #e0e0e0', borderRadius: 8, padding: '14px 18px', marginBottom: 14 };

// A merged availability row: one SSE (from the directory roster) plus any confirmed commitments.
interface IEntry {
  name: string;
  email: string;
  phone: string;
  bu: string;
  inDir: boolean;               // came from the Contact Directory roster
  items: ISseCommitment[];      // confirmed upcoming on-site / remote blocks
}

const norm = (s: string): string => (s || '').toLowerCase().trim();

export const SseAvailability: React.FC<ISseAvailabilityProps> = ({ sp, context }) => {
  const [loading, setLoading] = useState(true);
  const [commitments, setCommitments] = useState<ISseCommitment[]>([]);
  const [sses, setSses] = useState<IContact[]>([]);
  const [filter, setFilter] = useState('');

  useEffect(() => {
    Promise.all([
      new CseRequestService(sp).getSseCommitments().catch(() => [] as ISseCommitment[]),
      new ContactDirectoryService(context).getAll().catch(() => [] as IContact[])
    ]).then(([c, contacts]) => {
      setCommitments(c);
      setSses(contacts.filter(x => x.category === GENERALIST_CATEGORY));
      setLoading(false);
    }).catch(() => setLoading(false));
  }, []);

  if (loading) {
    return <div style={{ padding: 40, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading SSE availability…" /></div>;
  }

  // Build one entry per directory SSE, attaching commitments matched by email (fallback: name).
  const used = new Set<ISseCommitment>();
  const entries: IEntry[] = [];
  sses.forEach(s => {
    const semail = norm(s.email), sname = norm(s.name);
    const items = commitments.filter(c => {
      const ce = norm(c.sseEmail);
      if (ce && semail) return ce === semail;   // prefer email match
      return norm(c.sseName) === sname;          // fallback to name
    });
    items.forEach(i => used.add(i));
    entries.push({ name: s.name, email: s.email, phone: s.phone, bu: s.businessUnit, inDir: true, items });
  });

  // Any commitments whose SSE isn't in the directory (e.g. free-typed name) — keep them, grouped by name.
  const adhoc: { [name: string]: ISseCommitment[] } = {};
  commitments.forEach(c => {
    if (used.has(c)) return;
    const k = c.sseName || '(unassigned)';
    if (!adhoc[k]) adhoc[k] = [];
    adhoc[k].push(c);
  });
  Object.keys(adhoc).forEach(k => entries.push({
    name: k, email: adhoc[k][0].sseEmail || '', phone: '', bu: '', inDir: false, items: adhoc[k]
  }));

  const bookedCount = entries.filter(e => e.items.length > 0).length;

  // Filter by name; booked SSEs first, then alphabetical.
  const shown = entries
    .filter(e => !filter || e.name.toLowerCase().indexOf(filter.toLowerCase()) !== -1)
    .sort((a, b) => {
      const ab = a.items.length > 0 ? 0 : 1, bb = b.items.length > 0 ? 0 : 1;
      if (ab !== bb) return ab - bb;
      return norm(a.name) < norm(b.name) ? -1 : 1;
    });

  return (
    <div style={{ maxWidth: 980, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 18px', borderRadius: '6px 6px 0 0', marginBottom: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>🗓️ SSE Availability</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — Strategic SSE roster &amp; upcoming commitments</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <a href={SRT_DASHBOARD_URL}
            style={{ padding: '5px 14px', background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.35)', borderRadius: 4, color: '#fff', fontSize: 12, fontWeight: 600, textDecoration: 'none', whiteSpace: 'nowrap' }}>
            ← SRT Dashboard
          </a>
          <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)' }}>v{VERSION}</div>
        </div>
      </div>

      <div style={{ fontSize: 12, color: '#605e5c', marginBottom: 14 }}>
        Every Strategic SSE (from the Contact Directory) with their confirmed upcoming <strong>on-site</strong> and <strong>remote</strong> commitments — schedule new engagements around these to avoid double-booking. (Only locked-in dates show; tentative / proposed dates do not.)
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 14, flexWrap: 'wrap' }}>
        <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Filter by SSE name…"
          style={{ width: 260, maxWidth: '100%', fontSize: 13, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
        <div style={{ fontSize: 12, color: '#888' }}>
          {entries.length} SSE{entries.length === 1 ? '' : 's'} · <strong style={{ color: HPE_NAVY }}>{bookedCount}</strong> with upcoming commitments
        </div>
      </div>

      {entries.length === 0 && (
        <div style={CARD}>Couldn’t load the SSE roster from the Contact Directory, and there are no confirmed commitments to show.</div>
      )}

      {entries.length > 0 && shown.length === 0 && (
        <div style={CARD}>No SSE matches “{filter}”.</div>
      )}

      {shown.map(e => {
        const items = e.items.slice().sort((a, b) => (dstr(a.start) < dstr(b.start) ? -1 : 1));
        const busy = items.length > 0;
        return (
          <div key={(e.email || e.name) + (e.inDir ? '' : '~')} style={CARD}>
            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', marginBottom: busy ? 8 : 0, paddingBottom: busy ? 6 : 0, borderBottom: busy ? `2px solid ${HPE_GREEN}` : 'none' }}>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap' }}>
                <span style={{ fontSize: 14, fontWeight: 700, color: HPE_NAVY }}>{e.name}</span>
                {e.bu && <span style={{ fontSize: 10, color: '#605e5c', background: '#f3f2f1', borderRadius: 3, padding: '1px 6px' }}>{e.bu}</span>}
                {e.phone && <a href={`tel:${e.phone}`} style={{ fontSize: 12, color: HPE_GREEN, textDecoration: 'none' }}>📞 {e.phone}</a>}
                {!e.inDir && <span style={{ fontSize: 10, color: '#a4262c' }}>not in directory</span>}
              </div>
              <div style={{ fontSize: 11, color: busy ? '#888' : '#107c10', fontWeight: busy ? 400 : 600 }}>
                {busy ? `${items.length} upcoming` : 'Available'}
              </div>
            </div>
            {busy ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                {items.map((c, i) => (
                  <div key={i} style={{ fontSize: 13, color: '#323130' }}>
                    {c.type === 'On-site' ? '🏢' : '💻'} <strong>{fmtRange(c.start, c.end)}</strong>
                    <span style={{ color: '#605e5c' }}> · {c.type}{c.location ? ` · ${c.location}` : ''}</span>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ fontSize: 12, color: '#a19f9d', fontStyle: 'italic' }}>No upcoming confirmed commitments.</div>
            )}
          </div>
        );
      })}
    </div>
  );
};
