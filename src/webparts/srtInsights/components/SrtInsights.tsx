import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
import { ICseRequest } from '../../../models/ICseRequest';
import { IEnvironmentRow } from '../../../models/StrategicEngagement';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISrtInsightsProps {
  sp: SPFI;
  context: WebPartContext;
}

const VERSION = '1.0.3';
const SRT_DASHBOARD_URL = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SRT-Resource-Dashboard.aspx';

const parseEnv = (r: ICseRequest): IEnvironmentRow[] => {
  try { const a = JSON.parse(r.currentEnvironment || '[]'); return Array.isArray(a) ? a : []; } catch { return []; }
};
const sseName = (r: ICseRequest): string => {
  const v = r.requestedCse || '';
  const name = v.indexOf('/') !== -1 ? v.split('/')[0].trim() : v.trim();
  return name || '(unassigned)';
};
const isStrategic = (r: ICseRequest): boolean => r.engagementType === 'Strategic Engagement';
const isActive = (r: ICseRequest): boolean => ['Complete', 'Declined', 'Cancelled'].indexOf(r.requestStatus) === -1;

// ── Styles ──
const CARD: React.CSSProperties = { background: '#fff', border: '1px solid #e0e0e0', borderRadius: 8, padding: '16px 20px', marginBottom: 16 };
const SECTION_TITLE: React.CSSProperties = { fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 12, paddingBottom: 6, borderBottom: `2px solid ${HPE_GREEN}` };
const TH: React.CSSProperties = { textAlign: 'left', fontSize: 11, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.4px', padding: '6px 10px', borderBottom: '1px solid #eee' };
const TD: React.CSSProperties = { fontSize: 13, padding: '7px 10px', borderBottom: '1px solid #f3f3f3' };

const Tile: React.FC<{ label: string; value: number | string; color?: string }> = ({ label, value, color }) => (
  <div style={{ flex: '1 1 130px', background: '#fff', border: '1px solid #e0e0e0', borderRadius: 8, padding: '14px 10px', textAlign: 'center' }}>
    <div style={{ fontSize: 26, fontWeight: 700, color: color || HPE_NAVY }}>{value}</div>
    <div style={{ fontSize: 11, color: '#605e5c', marginTop: 2 }}>{label}</div>
  </div>
);

const Bar: React.FC<{ label: string; count: number; max: number; color: string }> = ({ label, count, max, color }) => (
  <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr 36px', gap: 10, alignItems: 'center', marginBottom: 6 }}>
    <div style={{ fontSize: 12.5, color: '#323130', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</div>
    <div style={{ background: '#f0f0f0', borderRadius: 4, height: 16, overflow: 'hidden' }}>
      <div style={{ width: `${max > 0 ? (count / max) * 100 : 0}%`, background: color, height: '100%', borderRadius: 4 }} />
    </div>
    <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textAlign: 'right' }}>{count}</div>
  </div>
);

export const SrtInsights: React.FC<ISrtInsightsProps> = ({ sp }) => {
  const [loading, setLoading] = useState(true);
  const [requests, setRequests] = useState<ICseRequest[]>([]);
  const [sseSort, setSseSort] = useState<{ field: 'name' | 'active' | 'strategic' | 'poc'; dir: 'asc' | 'desc' }>({ field: 'active', dir: 'desc' });
  const [vendorSort, setVendorSort] = useState<{ field: 'vendor' | 'displace' | 'integrate'; dir: 'asc' | 'desc' }>({ field: 'displace', dir: 'desc' });

  useEffect(() => {
    new CseRequestService(sp).getAll()
      .then(r => { setRequests(r); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading insights…" style={{ marginTop: 32 }} />;

  const total = requests.length;
  const strategic = requests.filter(isStrategic);
  const strategicCount = strategic.length;
  const pocCount = total - strategicCount;
  const activeCount = requests.filter(isActive).length;

  // Displacement / integration pipeline (from environment rows)
  const vendorAgg: Record<string, { displace: number; integrate: number; areas: Record<string, boolean> }> = {};
  requests.forEach(r => parseEnv(r).forEach(row => {
    if (!row.vendor || (row.disposition !== 'Displace' && row.disposition !== 'Integrate')) return;
    const v = vendorAgg[row.vendor] || { displace: 0, integrate: 0, areas: {} };
    if (row.disposition === 'Displace') v.displace += 1; else v.integrate += 1;
    if (row.solution) v.areas[row.solution] = true;
    vendorAgg[row.vendor] = v;
  }));
  const vendorRows = Object.keys(vendorAgg)
    .map(v => ({ vendor: v, ...vendorAgg[v], areaList: Object.keys(vendorAgg[v].areas) }))
    .sort((a, b) => {
      const cmp = vendorSort.field === 'vendor' ? a.vendor.localeCompare(b.vendor) : (a[vendorSort.field] as number) - (b[vendorSort.field] as number);
      return vendorSort.dir === 'asc' ? cmp : -cmp;
    });
  const vendorSortClick = (field: 'vendor' | 'displace' | 'integrate'): void =>
    setVendorSort(prev => prev.field === field ? { field, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { field, dir: field === 'vendor' ? 'asc' : 'desc' });
  const vendorIcon = (field: string): string => vendorSort.field !== field ? ' ⇅' : vendorSort.dir === 'asc' ? ' ▲' : ' ▼';
  const displaceTargets = vendorRows.reduce((s, v) => s + v.displace, 0);

  // SSE workload (active engagements per SSE)
  const sseAgg: Record<string, { active: number; strategic: number; poc: number }> = {};
  requests.filter(isActive).forEach(r => {
    const n = sseName(r);
    const a = sseAgg[n] || { active: 0, strategic: 0, poc: 0 };
    a.active += 1; if (isStrategic(r)) a.strategic += 1; else a.poc += 1;
    sseAgg[n] = a;
  });
  const sseRows = Object.keys(sseAgg).map(n => ({ name: n, ...sseAgg[n] })).sort((a, b) => {
    const cmp = sseSort.field === 'name' ? a.name.localeCompare(b.name) : (a[sseSort.field] as number) - (b[sseSort.field] as number);
    return sseSort.dir === 'asc' ? cmp : -cmp;
  });
  const sseSortClick = (field: 'name' | 'active' | 'strategic' | 'poc'): void =>
    setSseSort(prev => prev.field === field ? { field, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { field, dir: field === 'name' ? 'asc' : 'desc' });
  const sseIcon = (field: string): string => sseSort.field !== field ? ' ⇅' : sseSort.dir === 'asc' ? ' ▲' : ' ▼';

  // Desired outcome distribution (strategic only)
  const outcomeAgg: Record<string, number> = {};
  strategic.forEach(r => (r.desiredOutcome || []).forEach(o => { outcomeAgg[o] = (outcomeAgg[o] || 0) + 1; }));
  const outcomeRows = Object.keys(outcomeAgg).map(o => ({ outcome: o, count: outcomeAgg[o] })).sort((a, b) => b.count - a.count);
  const outcomeMax = outcomeRows.reduce((m, o) => Math.max(m, o.count), 0);

  return (
    <div style={{ maxWidth: 980, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 18px', borderRadius: '6px 6px 0 0', marginBottom: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>SSE Demand Insights</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — SSE demand, capacity & competitive pipeline</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <a href={SRT_DASHBOARD_URL}
            style={{ padding: '5px 14px', background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.35)',
              borderRadius: 4, color: '#fff', fontSize: 12, fontWeight: 600, textDecoration: 'none', whiteSpace: 'nowrap' }}>
            ← SRT Dashboard
          </a>
          <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)' }}>v{VERSION}</div>
        </div>
      </div>

      {/* KPI tiles */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, marginBottom: 16 }}>
        <Tile label="Total Requests" value={total} />
        <Tile label="POC Support" value={pocCount} color="#0078d4" />
        <Tile label="Strategic Engagements" value={strategicCount} color={HPE_GREEN} />
        <Tile label="Active" value={activeCount} color="#8a6000" />
        <Tile label="Displacement Targets" value={displaceTargets} color="#a4262c" />
      </div>

      {/* Demand split */}
      <div style={CARD}>
        <div style={SECTION_TITLE}>Demand Split</div>
        <Bar label="POC Support" count={pocCount} max={total} color="#0078d4" />
        <Bar label="Strategic Engagement" count={strategicCount} max={total} color={HPE_GREEN} />
      </div>

      {/* SSE Workload */}
      <div style={CARD}>
        <div style={SECTION_TITLE}>SSE Workload — active engagements</div>
        {sseRows.length === 0 ? <div style={{ fontSize: 13, color: '#888' }}>No active engagements.</div> : (
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr>
              <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => sseSortClick('name')}>SSE{sseIcon('name')}</th>
              <th style={{ ...TH, textAlign: 'right', cursor: 'pointer', userSelect: 'none' }} onClick={() => sseSortClick('active')}>Active{sseIcon('active')}</th>
              <th style={{ ...TH, textAlign: 'right', cursor: 'pointer', userSelect: 'none' }} onClick={() => sseSortClick('strategic')}>Strategic{sseIcon('strategic')}</th>
              <th style={{ ...TH, textAlign: 'right', cursor: 'pointer', userSelect: 'none' }} onClick={() => sseSortClick('poc')}>POC Support{sseIcon('poc')}</th>
            </tr></thead>
            <tbody>
              {sseRows.map(s => (
                <tr key={s.name}>
                  <td style={{ ...TD, fontWeight: 600 }}>{s.name}</td>
                  <td style={{ ...TD, textAlign: 'right', fontWeight: 700 }}>{s.active}</td>
                  <td style={{ ...TD, textAlign: 'right', color: HPE_GREEN }}>{s.strategic}</td>
                  <td style={{ ...TD, textAlign: 'right', color: '#0078d4' }}>{s.poc}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Competitive Displacement Pipeline */}
      <div style={CARD}>
        <div style={SECTION_TITLE}>Competitive Pipeline — displace vs integrate by vendor</div>
        {vendorRows.length === 0 ? <div style={{ fontSize: 13, color: '#888' }}>No competitive landscape captured yet.</div> : (
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr>
              <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => vendorSortClick('vendor')}>Incumbent Vendor{vendorIcon('vendor')}</th>
              <th style={{ ...TH, textAlign: 'right', cursor: 'pointer', userSelect: 'none' }} onClick={() => vendorSortClick('displace')}>🎯 Displace{vendorIcon('displace')}</th>
              <th style={{ ...TH, textAlign: 'right', cursor: 'pointer', userSelect: 'none' }} onClick={() => vendorSortClick('integrate')}>🤝 Integrate{vendorIcon('integrate')}</th>
              <th style={TH}>Solution Areas</th>
            </tr></thead>
            <tbody>
              {vendorRows.map(v => (
                <tr key={v.vendor}>
                  <td style={{ ...TD, fontWeight: 600 }}>{v.vendor}</td>
                  <td style={{ ...TD, textAlign: 'right', fontWeight: 700, color: '#a4262c' }}>{v.displace || '—'}</td>
                  <td style={{ ...TD, textAlign: 'right', color: '#8a6000' }}>{v.integrate || '—'}</td>
                  <td style={{ ...TD, fontSize: 12, color: '#605e5c' }}>{v.areaList.join(', ')}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Desired Outcome distribution */}
      <div style={CARD}>
        <div style={SECTION_TITLE}>Desired Outcomes — what engagements drive toward</div>
        {outcomeRows.length === 0 ? <div style={{ fontSize: 13, color: '#888' }}>No strategic engagements captured yet.</div> :
          outcomeRows.map(o => <Bar key={o.outcome} label={o.outcome} count={o.count} max={outcomeMax} color={HPE_NAVY} />)}
      </div>
    </div>
  );
};
