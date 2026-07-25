import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
import { ICseRequest, CSE_STATUS_STYLE, CUST_TEMP_STYLE, SCHEDULE_STATUS_STYLE } from '../../../models/ICseRequest';
import { SOLUTIONS } from '../../../models/ISolution';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

const codeToName = (codes: string): string => {
  if (!codes) return '—';
  return codes.split(',').map(c => {
    const sol = SOLUTIONS.find(s => s.code === c.trim());
    return sol ? sol.name : c.trim();
  }).join(', ');
};

export interface ISrtDashboardProps {
  sp: SPFI;
  context: WebPartContext;
}

const VERSION = '1.0.1';

const TH: React.CSSProperties = {
  padding: '8px 10px', textAlign: 'left', fontSize: 11, fontWeight: 700,
  textTransform: 'uppercase', letterSpacing: '0.4px', whiteSpace: 'nowrap',
};
const TD: React.CSSProperties = {
  padding: '8px 10px', verticalAlign: 'top',
};

export const SrtDashboard: React.FC<ISrtDashboardProps> = ({ sp, context }) => {
  const [requests, setRequests] = useState<ICseRequest[]>([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState('');

  const userEmail = context.pageContext.user.email.toLowerCase();
  const userName  = context.pageContext.user.displayName || userEmail;

  useEffect(() => {
    const svc = new CseRequestService(sp);
    svc.getAll()
      .then(all => { setRequests(all); setLoading(false); })
      .catch(err => { setError(String(err)); setLoading(false); });
  }, []);

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading SRT Dashboard…" style={{ marginTop: 40 }} />;

  if (error) return (
    <div style={{ padding: 24, color: '#a4262c', background: '#fde7e9', borderRadius: 6, margin: 16 }}>
      <strong>Error loading requests:</strong> {error}
    </div>
  );

  const pending    = requests.filter(r => r.requestStatus === 'Pending');
  const active     = requests.filter(r => r.requestStatus === 'Accepted' || r.requestStatus === 'Scheduled' || r.requestStatus === 'In Progress');
  const complete   = requests.filter(r => r.requestStatus === 'Complete');
  const needsSignOff = complete.filter(r => !r.signedOffBy);

  return (
    <div style={{ fontFamily: 'inherit', minHeight: 400 }}>

      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>SRT Dashboard</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — SSE Support Request Tracker</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <span style={{ fontSize: 13, color: 'rgba(255,255,255,0.8)' }}>Welcome, {userName}</span>
          <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.5)' }}>v{VERSION}</span>
        </div>
      </div>

      {/* KPI tiles */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, padding: '16px 20px 0' }}>
        {[
          { label: 'Pending',       count: pending.length,         bg: '#edebe9', color: '#605e5c' },
          { label: 'Active',        count: active.length,          bg: '#eff6fc', color: '#0078d4' },
          { label: 'Complete',      count: complete.length,        bg: '#e8faf3', color: '#107c10' },
          { label: 'Needs Sign-off', count: needsSignOff.length,   bg: '#fff4ce', color: '#8a6000' },
        ].map(tile => (
          <div key={tile.label} style={{ background: tile.bg, borderRadius: 6, padding: '12px 16px', textAlign: 'center' }}>
            <div style={{ fontSize: 28, fontWeight: 700, color: tile.color }}>{tile.count}</div>
            <div style={{ fontSize: 11, fontWeight: 600, color: tile.color, textTransform: 'uppercase', letterSpacing: '0.5px' }}>{tile.label}</div>
          </div>
        ))}
      </div>

      {/* Request table */}
      <div style={{ padding: '16px 20px' }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 10, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 6 }}>
          All Requests ({requests.length})
        </div>

        {requests.length === 0 ? (
          <div style={{ textAlign: 'center', padding: 40, color: '#888', fontSize: 14 }}>
            No SSE requests yet. Requests submitted via the SSE Request Form or POC Manager will appear here.
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ background: HPE_NAVY, color: '#fff' }}>
                  <th style={TH}>Customer</th>
                  <th style={TH}>SSE</th>
                  <th style={TH}>BU / Region</th>
                  <th style={TH}>Solutions</th>
                  <th style={TH}>Priority</th>
                  <th style={TH}>Status</th>
                  <th style={TH}>Schedule</th>
                  <th style={TH}>Temp</th>
                  <th style={TH}>Sign-off</th>
                </tr>
              </thead>
              <tbody>
                {requests.map((req, i) => {
                  const statusStyle   = CSE_STATUS_STYLE[req.requestStatus];
                  const schedStyle    = SCHEDULE_STATUS_STYLE[req.scheduleStatus];
                  const tempStyle     = CUST_TEMP_STYLE[req.custTemp];
                  const sseName       = req.requestedCse.includes('/') ? req.requestedCse.split('/')[0].trim() : req.requestedCse;
                  return (
                    <tr key={req.id ?? i} style={{ background: i % 2 === 0 ? '#fff' : '#faf9f8', borderBottom: '1px solid #edebe9' }}>
                      <td style={TD}>
                        <div style={{ fontWeight: 600 }}>{req.customerName || '—'}</div>
                        {req.pocName && <div style={{ fontSize: 11, color: '#888' }}>{req.pocName}</div>}
                        <div style={{ fontSize: 11, color: '#888' }}>{req.sePrimary.split('/')[0]?.trim()}</div>
                      </td>
                      <td style={TD}>{sseName || '—'}</td>
                      <td style={TD}>
                        <div>{req.hpenBusinessUnit}</div>
                        <div style={{ fontSize: 11, color: '#888' }}>{req.buRegion}</div>
                      </td>
                      <td style={{ ...TD, fontSize: 11, maxWidth: 140 }}>
                        {codeToName(req.solutionsFocus)}
                      </td>
                      <td style={TD}>
                        <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 700,
                          background: req.csePriority === 'High' ? '#fde7e9' : req.csePriority === 'Medium' ? '#fff4ce' : '#e8faf3',
                          color: req.csePriority === 'High' ? '#a4262c' : req.csePriority === 'Medium' ? '#8a6000' : '#107c10' }}>
                          {req.csePriority || '—'}
                        </span>
                      </td>
                      <td style={TD}>
                        <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
                          background: statusStyle.bg, color: statusStyle.color }}>
                          {req.requestStatus}
                        </span>
                      </td>
                      <td style={TD}>
                        <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
                          background: schedStyle.bg, color: schedStyle.color }}>
                          {req.scheduleStatus}
                        </span>
                      </td>
                      <td style={TD}>
                        <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
                          background: tempStyle.bg, color: tempStyle.color }}>
                          {req.custTemp}
                        </span>
                      </td>
                      <td style={TD}>
                        {req.signedOffBy
                          ? <div style={{ fontSize: 11, color: '#107c10' }}>✓ {req.signedOffBy}</div>
                          : req.requestStatus === 'Complete'
                            ? <span style={{ fontSize: 11, color: '#8a6000' }}>Pending</span>
                            : <span style={{ fontSize: 11, color: '#aaa' }}>—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

    </div>
  );
};
