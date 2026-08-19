import * as React from 'react';
import { useState, useEffect } from 'react';
import { APP_VERSION } from '../../../appVersion';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
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

export const SseAvailability: React.FC<ISseAvailabilityProps> = ({ sp }) => {
  const [loading, setLoading] = useState(true);
  const [commitments, setCommitments] = useState<ISseCommitment[]>([]);
  const [filter, setFilter] = useState('');

  useEffect(() => {
    new CseRequestService(sp).getSseCommitments()
      .then(c => setCommitments(c))
      .catch(() => setCommitments([]))
      .finally(() => setLoading(false));
  }, []);

  if (loading) {
    return <div style={{ padding: 40, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading SSE availability…" /></div>;
  }

  // Group commitments by SSE name.
  const groups: { [name: string]: ISseCommitment[] } = {};
  commitments.forEach(c => {
    const key = c.sseName || '(unassigned)';
    if (!groups[key]) groups[key] = [];
    groups[key].push(c);
  });
  const names = Object.keys(groups)
    .sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : 1))
    .filter(n => !filter || n.toLowerCase().indexOf(filter.toLowerCase()) !== -1);

  return (
    <div style={{ maxWidth: 980, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 18px', borderRadius: '6px 6px 0 0', marginBottom: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>🗓️ SSE Availability</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — confirmed upcoming SSE commitments</div>
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
        Confirmed upcoming <strong>on-site</strong> and <strong>remote</strong> commitments per SSE — schedule new engagements around these to avoid double-booking. (Tentative / proposed dates are not shown; only locked-in dates.)
      </div>

      <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Filter by SSE name…"
        style={{ width: 260, maxWidth: '100%', fontSize: 13, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4, marginBottom: 14, boxSizing: 'border-box' }} />

      {commitments.length === 0 && (
        <div style={CARD}>No confirmed upcoming commitments across any SSE right now. 👍</div>
      )}

      {commitments.length > 0 && names.length === 0 && (
        <div style={CARD}>No SSE matches “{filter}”.</div>
      )}

      {names.map(name => {
        const items = groups[name].slice().sort((a, b) => (dstr(a.start) < dstr(b.start) ? -1 : 1));
        return (
          <div key={name} style={CARD}>
            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', marginBottom: 8, paddingBottom: 6, borderBottom: `2px solid ${HPE_GREEN}` }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: HPE_NAVY }}>{name}</div>
              <div style={{ fontSize: 11, color: '#888' }}>{items.length} upcoming</div>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              {items.map((c, i) => (
                <div key={i} style={{ fontSize: 13, color: '#323130' }}>
                  {c.type === 'On-site' ? '🏢' : '💻'} <strong>{fmtRange(c.start, c.end)}</strong>
                  <span style={{ color: '#605e5c' }}> · {c.type}{c.location ? ` · ${c.location}` : ''}</span>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
};
