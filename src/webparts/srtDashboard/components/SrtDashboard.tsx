import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
import { ConfigService } from '../../../services/ConfigService';
import { ICseRequest, CSE_STATUS_STYLE, CUST_TEMP_STYLE, SCHEDULE_STATUS_STYLE, ScheduleStatus } from '../../../models/ICseRequest';
import { SOLUTIONS } from '../../../models/ISolution';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';
import { SrtAdminPanel } from './SrtAdminPanel';

const codeToName = (codes: string): string => {
  if (!codes) return '—';
  return codes.split(',').map(c => {
    const sol = SOLUTIONS.find(s => s.code === c.trim());
    return sol ? sol.name : c.trim();
  }).join(', ');
};

interface IDateEdit {
  remoteTbd: boolean; remoteStart: string; remoteEnd: string; remoteDuration: string;
  onsiteTbd: boolean; onsiteStart: string; onsiteEnd: string; onsiteDuration: string; onsiteDestination: string;
}

const CANCEL_REASONS = [
  'Customer Cancelled',
  'Partner Cancelled',
  'Self-Resolved',
  'Duplicate Request',
  'No Longer a Fit',
  'Other',
];

const calcOnsiteDays = (start: string, end: string): string => {
  if (!start || !end) return '';
  const days = Math.round((new Date(end).getTime() - new Date(start).getTime()) / 86400000);
  return days > 0 ? `${days} day${days !== 1 ? 's' : ''}` : '';
};

const toDateInput = (iso: string): string => iso ? iso.substring(0, 10) : '';

const fmtDate = (iso: string): string => {
  if (!iso) return '';
  const d = new Date(iso);
  return isNaN(d.getTime()) ? '' : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
};

const parseSseName = (raw: string): string => {
  // Strip parenthetical title, e.g. "(Distinguished Technologist - HPE Networking)"
  let name = raw.replace(/\s*\(.*\)\s*$/, '').trim();
  // Convert AD "Last, First" format to "First Last"
  if (name.includes(',')) {
    const parts = name.split(',');
    name = `${parts[1].trim()} ${parts[0].trim()}`;
  }
  return name || raw;
};

export interface ISrtDashboardProps {
  sp: SPFI;
  context: WebPartContext;
}

const VERSION = '1.0.25';

const TH: React.CSSProperties = {
  padding: '8px 10px', textAlign: 'left', fontSize: 11, fontWeight: 700,
  textTransform: 'uppercase', letterSpacing: '0.4px', whiteSpace: 'nowrap',
};
const TD: React.CSSProperties = {
  padding: '8px 10px', verticalAlign: 'top',
};

export const SrtDashboard: React.FC<ISrtDashboardProps> = ({ sp, context }) => {
  const [requests, setRequests]     = useState<ICseRequest[]>([]);
  const [loading, setLoading]       = useState(true);
  const [error, setError]           = useState('');
  const [isAdmin, setIsAdmin]           = useState(false);
  const [showAdmin, setShowAdmin]       = useState(false);
  const [savingId, setSavingId]         = useState<number | null>(null);
  const [decliningId, setDecliningId]   = useState<number | null>(null);
  const [declineNote, setDeclineNote]   = useState('');
  const [needsInfoId, setNeedsInfoId]           = useState<number | null>(null);
  const [needsInfoNote, setNeedsInfoNote]       = useState('');
  const [expandedId, setExpandedId]             = useState<number | null>(null);
  const [dateEdit, setDateEdit]                 = useState<IDateEdit | null>(null);
  const [savingDates, setSavingDates]           = useState(false);
  const [declineDatesId, setDeclineDatesId]     = useState<number | null>(null);
  const [declineDatesNote, setDeclineDatesNote] = useState('');
  const [cancelId, setCancelId]                 = useState<number | null>(null);
  const [cancelReason, setCancelReason]         = useState('');
  const [cancelNote, setCancelNote]             = useState('');
  const [signOffId, setSignOffId]               = useState<number | null>(null);
  const [signOffName, setSignOffName]           = useState('');
  const [filterStatus, setFilterStatus]         = useState('All');
  const [filterBU, setFilterBU]                 = useState('All');
  const [filterRegion, setFilterRegion]         = useState('All');
  const [filterPriority, setFilterPriority]     = useState('All');
  const [filterSearch, setFilterSearch]         = useState('');

  const userEmail    = context.pageContext.user.email.toLowerCase();
  const displayName  = context.pageContext.user.displayName || userEmail;
  const userName     = displayName.includes(',')
    ? displayName.split(',')[1].trim().split(' ')[0]
    : displayName.split(' ')[0];

  useEffect(() => {
    const svc       = new CseRequestService(sp);
    const configSvc = new ConfigService(sp);
    Promise.all([svc.getAll(), configSvc.isSuperUser(userEmail)])
      .then(([all, admin]) => { setRequests(all); setIsAdmin(admin); setLoading(false); })
      .catch(err => { setError(String(err)); setLoading(false); });
  }, []);

  const handleAccept = async (id: number): Promise<void> => {
    setSavingId(id);
    try {
      await new CseRequestService(sp).updateStatus(id, 'Accepted');
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: 'Accepted' } : r));
    } finally {
      setSavingId(null);
    }
  };

  const handleDeclineConfirm = async (id: number): Promise<void> => {
    setSavingId(id);
    try {
      await new CseRequestService(sp).updateStatus(id, 'Declined', declineNote || undefined);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: 'Declined', notes: declineNote || r.notes } : r));
      setDecliningId(null);
      setDeclineNote('');
    } finally {
      setSavingId(null);
    }
  };

  const handleNeedsInfoConfirm = async (id: number): Promise<void> => {
    setSavingId(id);
    try {
      await new CseRequestService(sp).updateStatus(id, 'Needs Info', needsInfoNote || undefined);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: 'Needs Info', notes: needsInfoNote || r.notes } : r));
      setNeedsInfoId(null);
      setNeedsInfoNote('');
    } finally {
      setSavingId(null);
    }
  };

  const handleExpand = (req: ICseRequest): void => {
    if (expandedId === req.id) {
      setExpandedId(null);
      setDateEdit(null);
      setDeclineDatesId(null);
      setDeclineDatesNote('');
    } else {
      setExpandedId(req.id!);
      setDateEdit({
        remoteTbd:        req.remoteTbd,
        remoteStart:      toDateInput(req.remoteStart),
        remoteEnd:        toDateInput(req.remoteEnd),
        remoteDuration:   req.remoteDuration,
        onsiteTbd:        req.onsiteTbd,
        onsiteStart:      toDateInput(req.onsiteStart),
        onsiteEnd:        toDateInput(req.onsiteEnd),
        onsiteDuration:   req.onsiteDuration,
        onsiteDestination: req.onsiteDestination,
      });
      setDeclineDatesId(null);
      setDeclineDatesNote('');
    }
  };

  const handleSaveDates = async (req: ICseRequest): Promise<void> => {
    if (!dateEdit) return;
    setSavingDates(true);
    const newStatus: ScheduleStatus = req.scheduleStatus === 'Dates Confirmed' ? 'Rescheduling' : 'Dates Proposed';
    try {
      await new CseRequestService(sp).updateDates(req.id!, { ...dateEdit, scheduleStatus: newStatus });
      setRequests(prev => prev.map(r => r.id === req.id
        ? { ...r, ...dateEdit, scheduleStatus: newStatus }
        : r
      ));
      setExpandedId(null);
      setDateEdit(null);
    } finally { setSavingDates(false); }
  };

  const handleConfirmDates = async (id: number): Promise<void> => {
    setSavingDates(true);
    try {
      await new CseRequestService(sp).confirmDates(id);
      setRequests(prev => prev.map(r => r.id === id
        ? { ...r, scheduleStatus: 'Dates Confirmed', requestStatus: 'Scheduled' }
        : r
      ));
      setExpandedId(null);
    } finally { setSavingDates(false); }
  };

  const handleDeclineDatesConfirm = async (id: number): Promise<void> => {
    setSavingDates(true);
    try {
      await new CseRequestService(sp).declineDates(id, declineDatesNote || undefined);
      setRequests(prev => prev.map(r => r.id === id
        ? { ...r, scheduleStatus: 'Rescheduling', notes: declineDatesNote || r.notes }
        : r
      ));
      setDeclineDatesId(null);
      setDeclineDatesNote('');
      setExpandedId(null);
    } finally { setSavingDates(false); }
  };

  const handleCancelConfirm = async (id: number): Promise<void> => {
    if (!cancelReason) return;
    setSavingDates(true);
    try {
      await new CseRequestService(sp).cancelRequest(id, cancelReason, cancelNote || undefined);
      setRequests(prev => prev.map(r => r.id === id
        ? { ...r, requestStatus: 'Cancelled', notes: [cancelReason, cancelNote].filter(Boolean).join(' — ') }
        : r
      ));
      setCancelId(null); setCancelReason(''); setCancelNote('');
      setExpandedId(null); setDateEdit(null);
    } finally { setSavingDates(false); }
  };

  const handleCustTemp = async (id: number, temp: string): Promise<void> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const t = temp as any;
    try {
      await new CseRequestService(sp).updateCustTemp(id, t);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, custTemp: t } : r));
    } catch { /* ignore */ }
  };

  const handleSignOff = async (id: number): Promise<void> => {
    if (!signOffName.trim()) return;
    setSavingId(id);
    try {
      await new CseRequestService(sp).signOff(id, signOffName.trim());
      setRequests(prev => prev.map(r => r.id === id ? { ...r, signedOffBy: signOffName.trim(), signOffDate: new Date().toISOString() } : r));
      setSignOffId(null);
      setSignOffName('');
    } finally { setSavingId(null); }
  };

  const canCancel = (req: ICseRequest): boolean => {
    if (['Complete', 'Cancelled', 'Declined'].includes(req.requestStatus)) return false;
    if (isAdmin) return true;
    return req.sePrimary.toLowerCase().includes(userEmail) && ['Pending', 'Accepted'].includes(req.requestStatus);
  };

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading SRT Dashboard…" style={{ marginTop: 40 }} />;

  if (error) return (
    <div style={{ padding: 24, color: '#a4262c', background: '#fde7e9', borderRadius: 6, margin: 16 }}>
      <strong>Error loading requests:</strong> {error}
    </div>
  );

  const visibleRequests = isAdmin
    ? requests
    : requests.filter(r =>
        r.sePrimary.toLowerCase().includes(userEmail) &&
        r.requestStatus !== 'Cancelled'
      );

  const buOptions     = Array.from(new Set(visibleRequests.map(r => r.hpenBusinessUnit).filter(Boolean))).sort();
  const regionOptions = Array.from(new Set(visibleRequests.map(r => r.buRegion).filter(Boolean))).sort();

  const filteredRequests = visibleRequests.filter(r => {
    if (filterStatus !== 'All' && r.requestStatus !== filterStatus) return false;
    if (filterBU !== 'All' && r.hpenBusinessUnit !== filterBU) return false;
    if (filterRegion !== 'All' && r.buRegion !== filterRegion) return false;
    if (filterPriority !== 'All' && r.csePriority !== filterPriority) return false;
    if (filterSearch && !r.customerName.toLowerCase().includes(filterSearch.toLowerCase())) return false;
    return true;
  });

  const pending      = visibleRequests.filter(r => r.requestStatus === 'Pending');
  const needsInfo    = visibleRequests.filter(r => r.requestStatus === 'Needs Info');
  const active       = visibleRequests.filter(r => r.requestStatus === 'Accepted' || r.requestStatus === 'Scheduled' || r.requestStatus === 'In Progress');
  const complete     = visibleRequests.filter(r => r.requestStatus === 'Complete');
  const needsSignOff = complete.filter(r => !r.signedOffBy);

  return (
    <div style={{ fontFamily: 'inherit', minHeight: 400 }}>

      {showAdmin && <SrtAdminPanel sp={sp} context={context} onClose={() => setShowAdmin(false)} />}

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
          <span style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)' }}>{(() => { const h = new Date().getHours(); return h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening'; })()}, {userName}</span>
          <a href="https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SSE-Request-Form.aspx"
            target="_blank" rel="noreferrer"
            style={{ padding: '4px 14px', background: HPE_GREEN, border: 'none', borderRadius: 4,
              color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer', textDecoration: 'none' }}>
            + New Request
          </a>
          <a href="https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/Home.aspx"
            target="_blank" rel="noreferrer"
            style={{ padding: '4px 14px', background: 'rgba(255,255,255,0.12)',
              border: '1px solid rgba(255,255,255,0.35)', borderRadius: 4,
              color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer', textDecoration: 'none' }}>
            POC Manager
          </a>
          {isAdmin && (
            <button
              onClick={() => setShowAdmin(true)}
              style={{
                padding: '4px 14px', background: 'rgba(255,255,255,0.12)',
                border: '1px solid rgba(255,255,255,0.35)', borderRadius: 4,
                color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer',
              }}>
              Admin
            </button>
          )}
          <span style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', letterSpacing: '0.5px' }}>v{VERSION}</span>
        </div>
      </div>

      {/* KPI tiles */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 12, padding: '16px 20px 0' }}>
        {[
          { label: 'Pending',        count: pending.length,       bg: '#edebe9', color: '#605e5c' },
          { label: 'Needs Info',     count: needsInfo.length,     bg: '#f0e6ff', color: '#6b2faf' },
          { label: 'Active',         count: active.length,        bg: '#eff6fc', color: '#0078d4' },
          { label: 'Complete',       count: complete.length,      bg: '#e8faf3', color: '#107c10' },
          { label: 'Needs Sign-off', count: needsSignOff.length,  bg: '#fff4ce', color: '#8a6000' },
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
          {isAdmin ? 'All Requests' : 'My Requests'} ({filteredRequests.length}{filteredRequests.length !== visibleRequests.length ? ` of ${visibleRequests.length}` : ''})
        </div>

        {/* Filter bar */}
        <div style={{ display: 'flex', gap: 8, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center' }}>
          <input
            type="text"
            value={filterSearch}
            onChange={e => setFilterSearch(e.target.value)}
            placeholder="Search customer…"
            style={{ fontSize: 12, padding: '5px 10px', border: '1px solid #ccc', borderRadius: 4, minWidth: 160 }}
          />
          <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
            style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
            <option value="All">All Status</option>
            {['Pending','Accepted','Scheduled','In Progress','Complete','Declined','Needs Info','Cancelled'].map(s =>
              <option key={s} value={s}>{s}</option>)}
          </select>
          <select value={filterBU} onChange={e => setFilterBU(e.target.value)}
            style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
            <option value="All">All BUs</option>
            {buOptions.map(bu => <option key={bu} value={bu}>{bu}</option>)}
          </select>
          <select value={filterRegion} onChange={e => setFilterRegion(e.target.value)}
            style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
            <option value="All">All Regions</option>
            {regionOptions.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
          <select value={filterPriority} onChange={e => setFilterPriority(e.target.value)}
            style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
            <option value="All">All Priority</option>
            {['Low','Medium','High','Critical'].map(p => <option key={p} value={p}>{p}</option>)}
          </select>
          {(filterStatus !== 'All' || filterBU !== 'All' || filterRegion !== 'All' || filterPriority !== 'All' || filterSearch) && (
            <button onClick={() => { setFilterStatus('All'); setFilterBU('All'); setFilterRegion('All'); setFilterPriority('All'); setFilterSearch(''); }}
              style={{ fontSize: 11, padding: '5px 10px', background: '#f3f2f1', border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#605e5c' }}>
              Clear Filters
            </button>
          )}
        </div>

        {visibleRequests.length === 0 ? (
          <div style={{ textAlign: 'center', padding: 40, color: '#888', fontSize: 14 }}>
            No SSE requests yet. Requests submitted via the SSE Request Form or POC Manager will appear here.
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ background: HPE_NAVY, color: '#fff' }}>
                  <th style={TH}>Customer</th>
                  <th style={TH}>SE</th>
                  <th style={TH}>SSE</th>
                  <th style={TH}>BU / Region</th>
                  <th style={TH}>Solutions</th>
                  <th style={TH}>Priority</th>
                  <th style={TH}>Status</th>
                  <th style={TH}>Schedule</th>
                  <th style={TH}>Temp</th>
                  <th style={TH}>Sign-off</th>
                  {isAdmin && <th style={TH}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {filteredRequests.map((req, i) => {
                  const statusStyle   = CSE_STATUS_STYLE[req.requestStatus]  || CSE_STATUS_STYLE.Pending;
                  const schedStyle    = SCHEDULE_STATUS_STYLE[req.scheduleStatus] || SCHEDULE_STATUS_STYLE.TBD;
                  const tempStyle     = CUST_TEMP_STYLE[req.custTemp]    || CUST_TEMP_STYLE.Normal;
                  const sseName       = parseSseName(req.requestedCse.includes('/') ? req.requestedCse.split('/')[0].trim() : req.requestedCse);
                  const isExpanded    = expandedId === req.id;
                  const isCancelled   = req.requestStatus === 'Cancelled';
                  const canEditDates  = (isAdmin || req.sePrimary.toLowerCase().includes(userEmail))
                                        && !['Declined', 'Complete', 'Cancelled'].includes(req.requestStatus);
                  const showDateActions = isAdmin && (req.scheduleStatus === 'Dates Proposed' || req.scheduleStatus === 'Rescheduling');
                  const colSpan       = isAdmin ? 11 : 10;
                  const rowBg         = isExpanded ? '#f0ebff' : isCancelled ? '#f8f8f8' : i % 2 === 0 ? '#fff' : '#faf9f8';
                  return (
                    <React.Fragment key={req.id ?? i}>
                    <tr style={{ background: rowBg, borderBottom: isExpanded ? 'none' : '1px solid #edebe9', cursor: 'pointer' }}
                        onClick={() => handleExpand(req)}>
                      <td style={{ ...TD, opacity: isCancelled ? 0.6 : 1 }}>
                        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 6 }}>
                          <span style={{ fontSize: 10, color: '#6b2faf', marginTop: 3, flexShrink: 0 }}>{isExpanded ? '▾' : '▸'}</span>
                          <div>
                            <div style={{ fontWeight: 600, textDecoration: isCancelled ? 'line-through' : 'none' }}>{req.customerName || '—'}</div>
                            {req.pocName && <div style={{ fontSize: 11, color: '#888' }}>{req.pocName}</div>}
                          </div>
                        </div>
                      </td>
                      <td style={{ ...TD, fontSize: 12, opacity: isCancelled ? 0.6 : 1 }}>
                        {parseSseName(req.sePrimary.split('/')[0]?.trim() || '') || '—'}
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
                        {!req.remoteTbd && req.remoteStart && (
                          <div style={{ fontSize: 10, color: '#555', marginTop: 3 }}>
                            Remote: {fmtDate(req.remoteStart)}{req.remoteEnd ? ` – ${fmtDate(req.remoteEnd)}` : ''}
                          </div>
                        )}
                        {!req.onsiteTbd && req.onsiteStart && (
                          <div style={{ fontSize: 10, color: '#555', marginTop: 1 }}>
                            Onsite: {fmtDate(req.onsiteStart)}{req.onsiteEnd ? ` – ${fmtDate(req.onsiteEnd)}` : ''}
                          </div>
                        )}
                      </td>
                      <td style={TD} onClick={e => e.stopPropagation()}>
                        {isAdmin ? (
                          <select
                            value={req.custTemp}
                            onChange={e => handleCustTemp(req.id!, e.target.value).catch(() => undefined)}
                            style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
                              background: tempStyle.bg, color: tempStyle.color,
                              border: 'none', cursor: 'pointer' }}>
                            {(['Low','Normal','High','Critical']).map(t => <option key={t} value={t}>{t}</option>)}
                          </select>
                        ) : (
                          <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
                            background: tempStyle.bg, color: tempStyle.color }}>
                            {req.custTemp}
                          </span>
                        )}
                      </td>
                      <td style={TD} onClick={e => e.stopPropagation()}>
                        {req.signedOffBy ? (
                          <div style={{ fontSize: 11, color: '#107c10' }}>✓ {req.signedOffBy}</div>
                        ) : req.requestStatus === 'Complete' ? (
                          isAdmin && signOffId === req.id ? (
                            <div>
                              <input
                                type="text"
                                value={signOffName}
                                onChange={e => setSignOffName(e.target.value)}
                                placeholder="Your name"
                                autoFocus
                                style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }}
                              />
                              <div style={{ display: 'flex', gap: 4 }}>
                                <button
                                  disabled={!signOffName.trim() || savingId === req.id}
                                  onClick={() => handleSignOff(req.id!).catch(() => undefined)}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                  {savingId === req.id ? '…' : '✓ Confirm'}
                                </button>
                                <button
                                  onClick={() => { setSignOffId(null); setSignOffName(''); }}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>
                                  Cancel
                                </button>
                              </div>
                            </div>
                          ) : isAdmin ? (
                            <button
                              onClick={() => { setSignOffId(req.id!); setSignOffName(''); }}
                              style={{ fontSize: 11, padding: '3px 10px', background: '#e8faf3', color: '#107c10', border: '1px solid #107c10', borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>
                              Sign Off
                            </button>
                          ) : (
                            <span style={{ fontSize: 11, color: '#8a6000' }}>Pending</span>
                          )
                        ) : (
                          <span style={{ fontSize: 11, color: '#aaa' }}>—</span>
                        )}
                      </td>
                      {isAdmin && <td style={{ ...TD, minWidth: 160 }} onClick={e => e.stopPropagation()}>
                        {req.requestStatus === 'Pending' && (
                          decliningId === req.id ? (
                            <div>
                              <input
                                type="text"
                                value={declineNote}
                                onChange={e => setDeclineNote(e.target.value)}
                                placeholder="Reason (optional)"
                                style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }}
                              />
                              <div style={{ display: 'flex', gap: 4 }}>
                                <button
                                  disabled={savingId === req.id}
                                  onClick={() => handleDeclineConfirm(req.id!)}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#a4262c', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                  {savingId === req.id ? '…' : 'Confirm'}
                                </button>
                                <button
                                  onClick={() => { setDecliningId(null); setDeclineNote(''); }}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>
                                  Cancel
                                </button>
                              </div>
                            </div>
                          ) : needsInfoId === req.id ? (
                            <div>
                              <input
                                type="text"
                                value={needsInfoNote}
                                onChange={e => setNeedsInfoNote(e.target.value)}
                                placeholder="What info is needed?"
                                style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }}
                              />
                              <div style={{ display: 'flex', gap: 4 }}>
                                <button
                                  disabled={savingId === req.id}
                                  onClick={() => handleNeedsInfoConfirm(req.id!)}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#6b2faf', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                  {savingId === req.id ? '…' : 'Send'}
                                </button>
                                <button
                                  onClick={() => { setNeedsInfoId(null); setNeedsInfoNote(''); }}
                                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>
                                  Cancel
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                              <button
                                disabled={savingId === req.id}
                                onClick={() => handleAccept(req.id!)}
                                style={{ fontSize: 11, padding: '4px 10px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                {savingId === req.id ? '…' : '✓ Accept'}
                              </button>
                              <button
                                onClick={() => { setNeedsInfoId(req.id!); setNeedsInfoNote(''); }}
                                style={{ fontSize: 11, padding: '4px 10px', background: '#f0e6ff', color: '#6b2faf', border: '1px solid #6b2faf', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                ? Info
                              </button>
                              <button
                                onClick={() => { setDecliningId(req.id!); setDeclineNote(''); }}
                                style={{ fontSize: 11, padding: '4px 10px', background: '#fde7e9', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                                ✕ Decline
                              </button>
                            </div>
                          )
                        )}
                        {req.requestStatus !== 'Pending' && (
                          <span style={{ fontSize: 11, color: '#aaa' }}>—</span>
                        )}
                      </td>}
                    </tr>

                    {/* ── Expanded date drawer ── */}
                    {isExpanded && dateEdit && (
                      <tr style={{ background: '#f8f5ff', borderBottom: '2px solid #6b2faf' }}>
                        <td colSpan={colSpan} style={{ padding: '16px 20px' }} onClick={e => e.stopPropagation()}>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>

                            {/* Remote section */}
                            <div>
                              <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>
                                Remote Support
                              </div>
                              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, marginBottom: 8, cursor: canEditDates ? 'pointer' : 'default' }}>
                                <input type="checkbox" checked={dateEdit.remoteTbd}
                                  disabled={!canEditDates}
                                  onChange={e => setDateEdit(prev => prev ? { ...prev, remoteTbd: e.target.checked, remoteStart: '', remoteEnd: '' } : prev)} />
                                TBD — dates not yet set
                              </label>
                              {!dateEdit.remoteTbd && (
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Start</div>
                                    <input type="date" value={dateEdit.remoteStart} disabled={!canEditDates}
                                      onChange={e => setDateEdit(prev => prev ? { ...prev, remoteStart: e.target.value } : prev)}
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>End</div>
                                    <input type="date" value={dateEdit.remoteEnd} disabled={!canEditDates}
                                      onChange={e => setDateEdit(prev => prev ? { ...prev, remoteEnd: e.target.value } : prev)}
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                  <div style={{ gridColumn: '1 / -1' }}>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Duration (e.g. 2 days)</div>
                                    <input type="text" value={dateEdit.remoteDuration} disabled={!canEditDates}
                                      onChange={e => setDateEdit(prev => prev ? { ...prev, remoteDuration: e.target.value } : prev)}
                                      placeholder="e.g. 2 days"
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                </div>
                              )}
                            </div>

                            {/* Onsite section */}
                            <div>
                              <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>
                                On-Site Support
                              </div>
                              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, marginBottom: 8, cursor: canEditDates ? 'pointer' : 'default' }}>
                                <input type="checkbox" checked={dateEdit.onsiteTbd}
                                  disabled={!canEditDates}
                                  onChange={e => setDateEdit(prev => prev ? { ...prev, onsiteTbd: e.target.checked, onsiteStart: '', onsiteEnd: '' } : prev)} />
                                TBD — dates not yet set
                              </label>
                              {!dateEdit.onsiteTbd && (
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Start</div>
                                    <input type="date" value={dateEdit.onsiteStart} disabled={!canEditDates}
                                      onChange={e => {
                                        const s = e.target.value;
                                        setDateEdit(prev => prev ? { ...prev, onsiteStart: s, onsiteDuration: calcOnsiteDays(s, prev.onsiteEnd) } : prev);
                                      }}
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>End</div>
                                    <input type="date" value={dateEdit.onsiteEnd} disabled={!canEditDates}
                                      onChange={e => {
                                        const en = e.target.value;
                                        setDateEdit(prev => prev ? { ...prev, onsiteEnd: en, onsiteDuration: calcOnsiteDays(prev.onsiteStart, en) } : prev);
                                      }}
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Duration (auto-calculated)</div>
                                    <input type="text" value={dateEdit.onsiteDuration} disabled={!canEditDates}
                                      onChange={e => setDateEdit(prev => prev ? { ...prev, onsiteDuration: e.target.value } : prev)}
                                      placeholder="Set dates above"
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box', background: '#f5f5f5' }} />
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Destination</div>
                                    <input type="text" value={dateEdit.onsiteDestination} disabled={!canEditDates}
                                      onChange={e => setDateEdit(prev => prev ? { ...prev, onsiteDestination: e.target.value } : prev)}
                                      placeholder="City, ST"
                                      style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                                  </div>
                                </div>
                              )}
                            </div>
                          </div>

                          {/* Action bar */}
                          <div style={{ marginTop: 16, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', borderTop: '1px solid #e0d8f0', paddingTop: 12 }}>

                            {/* SE / admin: save dates */}
                            {canEditDates && (
                              <button
                                disabled={savingDates}
                                onClick={() => handleSaveDates(req).catch(() => undefined)}
                                style={{ padding: '6px 18px', background: '#6b2faf', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', opacity: savingDates ? 0.6 : 1 }}>
                                {savingDates ? 'Saving…' : req.scheduleStatus === 'Dates Confirmed' ? 'Save (will flag as Rescheduling)' : 'Save Dates'}
                              </button>
                            )}

                            {/* Admin / Charlie: confirm or decline proposed dates */}
                            {showDateActions && !canEditDates === false && (
                              declineDatesId === req.id ? (
                                <>
                                  <input type="text" value={declineDatesNote}
                                    onChange={e => setDeclineDatesNote(e.target.value)}
                                    placeholder="Reason dates don't work (optional)"
                                    style={{ flex: 1, minWidth: 200, fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                                  <button disabled={savingDates}
                                    onClick={() => handleDeclineDatesConfirm(req.id!).catch(() => undefined)}
                                    style={{ padding: '5px 14px', background: '#a4262c', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                    {savingDates ? '…' : 'Confirm Decline'}
                                  </button>
                                  <button onClick={() => { setDeclineDatesId(null); setDeclineDatesNote(''); }}
                                    style={{ padding: '5px 14px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                                    Cancel
                                  </button>
                                </>
                              ) : (
                                <>
                                  <button disabled={savingDates}
                                    onClick={() => handleConfirmDates(req.id!).catch(() => undefined)}
                                    style={{ padding: '6px 18px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                    ✓ Confirm Dates
                                  </button>
                                  <button onClick={() => setDeclineDatesId(req.id!)}
                                    style={{ padding: '6px 14px', background: '#fde7e9', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                    ✕ Dates Don&apos;t Work
                                  </button>
                                </>
                              )
                            )}

                            <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
                              {canCancel(req) && cancelId !== req.id && (
                                <button
                                  onClick={() => { setCancelId(req.id!); setCancelReason(''); setCancelNote(''); }}
                                  style={{ padding: '5px 14px', background: '#fff', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                  Cancel Request
                                </button>
                              )}
                              <button onClick={() => { setExpandedId(null); setDateEdit(null); setCancelId(null); }}
                                style={{ padding: '5px 14px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                                Close
                              </button>
                            </div>
                          </div>

                          {/* Cancel confirmation section */}
                          {cancelId === req.id && (
                            <div style={{ marginTop: 12, padding: '14px 16px', background: '#fde7e9', border: '1px solid #a4262c', borderRadius: 6 }}>
                              <div style={{ fontSize: 12, fontWeight: 700, color: '#a4262c', marginBottom: 10 }}>Cancel this SSE Request</div>
                              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 10 }}>
                                <div>
                                  <div style={{ fontSize: 11, color: '#555', marginBottom: 4 }}>Reason <span style={{ color: '#a4262c' }}>*</span></div>
                                  <select value={cancelReason} onChange={e => setCancelReason(e.target.value)}
                                    style={{ width: '100%', fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
                                    <option value="">— Select reason —</option>
                                    {CANCEL_REASONS.map(r => <option key={r} value={r}>{r}</option>)}
                                  </select>
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, color: '#555', marginBottom: 4 }}>Additional notes (optional)</div>
                                  <input type="text" value={cancelNote} onChange={e => setCancelNote(e.target.value)}
                                    placeholder="Any additional context..."
                                    style={{ width: '100%', fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
                                </div>
                              </div>
                              <div style={{ display: 'flex', gap: 8 }}>
                                <button disabled={!cancelReason || savingDates}
                                  onClick={() => handleCancelConfirm(req.id!).catch(() => undefined)}
                                  style={{ padding: '6px 18px', background: cancelReason ? '#a4262c' : '#e0e0e0', color: cancelReason ? '#fff' : '#aaa', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: cancelReason ? 'pointer' : 'not-allowed' }}>
                                  {savingDates ? 'Cancelling…' : 'Confirm Cancellation'}
                                </button>
                                <button onClick={() => { setCancelId(null); setCancelReason(''); setCancelNote(''); }}
                                  style={{ padding: '6px 14px', background: '#fff', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                                  Never Mind
                                </button>
                              </div>
                            </div>
                          )}

                          {req.scheduleStatus !== 'TBD' && (
                            <div style={{ marginTop: 8, fontSize: 11, color: '#888' }}>
                              Current schedule status: <strong>{req.scheduleStatus}</strong>
                              {req.scheduleStatus === 'Dates Confirmed' && ' — saving new dates will flag this as Rescheduling and notify the SSE.'}
                              {req.scheduleStatus === 'Dates Proposed' && ' — awaiting SSE confirmation.'}
                              {req.scheduleStatus === 'Rescheduling' && ' — SSE declined the previous dates. Update and save to re-propose.'}
                            </div>
                          )}
                        </td>
                      </tr>
                    )}
                    </React.Fragment>
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
