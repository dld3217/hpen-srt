import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { APP_VERSION } from '../../../appVersion';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { CseRequestService } from '../../../services/CseRequestService';
import { ConfigService } from '../../../services/ConfigService';
import { ICseRequest, CseRequestStatus, CSE_STATUS_STYLE, CUST_TEMP_STYLE, SCHEDULE_STATUS_STYLE, ScheduleStatus } from '../../../models/ICseRequest';
import { SOLUTIONS, SOLUTION_CATEGORIES } from '../../../models/ISolution';
import { DISPOSITION_STYLE, IEnvironmentRow } from '../../../models/StrategicEngagement';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';
import { SrtAdminPanel } from './SrtAdminPanel';

const parseEnvRows = (json: string | undefined): IEnvironmentRow[] => {
  try { const a = JSON.parse(json || '[]'); return Array.isArray(a) ? a : []; } catch { return []; }
};
const LTH: React.CSSProperties = { textAlign: 'left', fontSize: 10, fontWeight: 700, color: '#8a7fb0', textTransform: 'uppercase', letterSpacing: '0.3px', padding: '4px 8px', borderBottom: '1px solid #e0d8f0' };
const LTD: React.CSSProperties = { fontSize: 12, padding: '5px 8px', borderBottom: '1px solid #efeaf7', verticalAlign: 'top' };

const emailToName = (email: string): string => {
  if (!email) return '';
  const local = email.split('@')[0];
  return local.split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
};

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
  // Treat as a calendar date (YYYY-MM-DD…) so UTC→local never rolls it back a day.
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(iso);
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

const VERSION = APP_VERSION;

const TH: React.CSSProperties = {
  padding: '7px 7px', textAlign: 'left', fontSize: 11, fontWeight: 700,
  textTransform: 'uppercase', letterSpacing: '0.4px', whiteSpace: 'nowrap',
};
const TD: React.CSSProperties = {
  padding: '7px 7px', verticalAlign: 'top',
};

// ── Header nav targets (same site collection) ─────────────────────────────────
const POC_HOME_URL   = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/Home.aspx';
const INSIGHTS_URL   = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SSE-Demand-Insights.aspx';
// Front-door page ("Request SSE Support"); ?form=strategic opens straight to the Strategic Engagement form.
const FRONT_DOOR_URL = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/Strategic-Engagement-Request.aspx';
const HDR_GREEN: React.CSSProperties = { padding: '5px 14px', background: HPE_GREEN, border: 'none', borderRadius: 4, color: '#fff', fontSize: 12, fontWeight: 700, cursor: 'pointer', textDecoration: 'none', whiteSpace: 'nowrap' };
const HDR_OUTLINE: React.CSSProperties = { padding: '5px 14px', background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.35)', borderRadius: 4, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer', textDecoration: 'none', whiteSpace: 'nowrap' };

// TEST MODE shortcuts — type a nickname instead of the full email
const ACT_AS_ALIASES: Record<string, string> = {
  david: 'david.ball@hpe.com',
  charlie: 'charlie.clemmer@hpe.com',
  mike: 'mike.bruno@hpe.com',
  rick: 'rick.watkins@hpe.com',
};
// Resolve "David" / "Charlie Clemmer" / bare "john.smith" → a usable email.
function resolveActAs(raw: string): string {
  const t = (raw || '').trim().toLowerCase();
  if (!t) return '';
  if (ACT_AS_ALIASES[t]) return ACT_AS_ALIASES[t];
  const first = t.split(/\s+/)[0];
  if (ACT_AS_ALIASES[first]) return ACT_AS_ALIASES[first];
  return t.indexOf('@') === -1 ? `${t}@hpe.com` : t;
}

export const SrtDashboard: React.FC<ISrtDashboardProps> = ({ sp, context }) => {
  const [requests, setRequests]     = useState<ICseRequest[]>([]);
  const [loading, setLoading]       = useState(true);
  const [error, setError]           = useState('');
  const [isAdmin, setIsAdmin]           = useState(false);
  const [isSED, setIsSED]               = useState(false);
  const [showAdmin, setShowAdmin]       = useState(false);
  const [savingId, setSavingId]         = useState<number | null>(null);
  const [decliningId, setDecliningId]   = useState<number | null>(null);
  const [declineNote, setDeclineNote]   = useState('');
  const [needsInfoId, setNeedsInfoId]           = useState<number | null>(null);
  const [needsInfoNote, setNeedsInfoNote]       = useState('');
  const [reassignId, setReassignId]             = useState<number | null>(null);
  const [reassignSse, setReassignSse]           = useState('');
  const [showOnsitePanel, setShowOnsitePanel]   = useState(true);
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
  const [filterType, setFilterType]             = useState('All');
  const [filterSearch, setFilterSearch]         = useState('');
  const [viewMode, setViewMode]                 = useState<'mine' | 'all'>('mine');
  const [actAs, setActAs]                       = useState<string>(() => (localStorage.getItem('srt_actAs') || '').toLowerCase());
  const [actAsInput, setActAsInput]             = useState('');
  const [realIsAdmin, setRealIsAdmin]           = useState(false);
  const [showPendingSection, setShowPendingSection]     = useState(true);
  const [showAcceptedSection, setShowAcceptedSection]   = useState(true);
  const [showParkedSection, setShowParkedSection]       = useState(true);
  const [showDeclinedSection, setShowDeclinedSection]   = useState(false);
  const [showCompletedSection, setShowCompletedSection] = useState(true);
  const [editSolId, setEditSolId]               = useState<number | null>(null);
  const [editSolCodes, setEditSolCodes]         = useState<Set<string>>(new Set());
  const [activeTile, setActiveTile]             = useState<string | null>(null);
  const [sortField, setSortField]               = useState<'customer' | 'priority' | 'status' | 'type'>('customer');
  const [sortDir, setSortDir]                   = useState<'asc' | 'desc'>('asc');
  const [drawerOpportunity, setDrawerOpportunity] = useState('');
  const [drawerNotes, setDrawerNotes]             = useState('');
  const [savingNotes, setSavingNotes]             = useState(false);
  const [urlActionBanner, setUrlActionBanner]     = useState<{ msg: string; type: 'ok' | 'err' } | null>(null);
  const [selectedIds, setSelectedIds]             = useState<Set<number>>(new Set());
  const tableRef                                = useRef<HTMLDivElement>(null);

  const realEmail    = context.pageContext.user.email.toLowerCase();
  const userEmail    = actAs || realEmail;                 // EFFECTIVE identity (spoofable for testing)
  const displayName  = context.pageContext.user.displayName || realEmail;
  const realFirst    = displayName.includes(',')
    ? displayName.split(',')[1].trim().split(' ')[0]
    : displayName.split(' ')[0];
  const userName     = actAs ? (emailToName(actAs).split(' ')[0] || actAs) : realFirst;

  const applyActAs = (email: string): void => {
    const e = resolveActAs(email);
    if (!e) return;
    setActAs(e); localStorage.setItem('srt_actAs', e); setActAsInput(''); setExpandedId(null);
  };
  const resetActAs = (): void => {
    setActAs(''); localStorage.removeItem('srt_actAs'); setActAsInput(''); setExpandedId(null);
  };

  // Requests load once.
  useEffect(() => {
    new CseRequestService(sp).getAll()
      .then(all => { setRequests(all); setLoading(false); })
      .catch(err => { setError(String(err)); setLoading(false); });
  }, []);

  // Real admin flag — from the REAL user, computed once — gates the test bar so you can't lock yourself out.
  useEffect(() => {
    new ConfigService(sp).isSuperUser(realEmail).then(setRealIsAdmin).catch(() => undefined);
  }, []);

  // Role + default view for the EFFECTIVE user — recomputes whenever you "view as" someone.
  useEffect(() => {
    const configSvc = new ConfigService(sp);
    Promise.all([configSvc.isSuperUser(userEmail), configSvc.isSED(userEmail)])
      .then(([admin, sed]) => {
        setIsAdmin(admin); setIsSED(sed);
        setViewMode(admin || sed ? 'all' : 'mine');
      })
      .catch(() => undefined);
  }, [userEmail]);

  useEffect(() => {
    if (loading) return;
    const params = new URLSearchParams(window.location.search);
    const action = params.get('action');
    const requestId = parseInt(params.get('requestId') || '0', 10);
    if (!action || !requestId) return;
    const req = requests.find(r => r.id === requestId);
    if (!req) {
      setUrlActionBanner({ msg: `Request #${requestId} not found or not accessible.`, type: 'err' });
      return;
    }
    const svc = new CseRequestService(sp);
    if (action === 'accept') {
      svc.updateStatus(requestId, 'Accepted')
        .then(() => {
          setRequests(prev => prev.map(r => r.id === requestId ? { ...r, requestStatus: 'Accepted' } : r));
          setUrlActionBanner({ msg: `✓ Request #${requestId} — ${req.customerName} accepted successfully.`, type: 'ok' });
        })
        .catch(e => setUrlActionBanner({ msg: `Accept failed: ${(e as Error).message}`, type: 'err' }));
    } else if (action === 'decline') {
      svc.updateStatus(requestId, 'Declined')
        .then(() => {
          setRequests(prev => prev.map(r => r.id === requestId ? { ...r, requestStatus: 'Declined' } : r));
          setUrlActionBanner({ msg: `Request #${requestId} — ${req.customerName} declined.`, type: 'ok' });
        })
        .catch(e => setUrlActionBanner({ msg: `Decline failed: ${(e as Error).message}`, type: 'err' }));
    } else if (action === 'info') {
      svc.updateStatus(requestId, 'Needs Info')
        .then(() => {
          setRequests(prev => prev.map(r => r.id === requestId ? { ...r, requestStatus: 'Needs Info' } : r));
          setUrlActionBanner({ msg: `Request #${requestId} — ${req.customerName} marked Needs Info.`, type: 'ok' });
        })
        .catch(e => setUrlActionBanner({ msg: `Update failed: ${(e as Error).message}`, type: 'err' }));
    } else if (action === 'reassign') {
      setReassignId(requestId);
      setUrlActionBanner({ msg: `Use the reassign input to assign a new SSE for request #${requestId} — ${req.customerName}.`, type: 'ok' });
    }
  }, [loading]);

  const handleAccept = async (id: number): Promise<void> => {
    setSavingId(id);
    try {
      await new CseRequestService(sp).updateStatus(id, 'Accepted');
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: 'Accepted' } : r));
    } finally {
      setSavingId(null);
    }
  };

  const handleReassignConfirm = async (id: number): Promise<void> => {
    if (!reassignSse.includes('/')) return;
    setSavingId(id);
    try {
      await new CseRequestService(sp).reassign(id, reassignSse.trim(), '');
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestedCse: reassignSse.trim(), requestStatus: 'Accepted' } : r));
      setReassignId(null);
      setReassignSse('');
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
      setDrawerOpportunity('');
      setDrawerNotes('');
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
      setDrawerOpportunity(req.opportunity || '');
      setDrawerNotes(req.notes || '');
      setDeclineDatesId(null);
      setDeclineDatesNote('');
    }
  };

  const handleSaveNotes = async (id: number): Promise<void> => {
    setSavingNotes(true);
    try {
      await new CseRequestService(sp).updateOpportunityAndNotes(id, drawerOpportunity, drawerNotes);
      setRequests(prev => prev.map(r => r.id === id
        ? { ...r, opportunity: drawerOpportunity, notes: drawerNotes }
        : r
      ));
    } finally { setSavingNotes(false); }
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
      // Confirming dates always lands the engagement on Scheduled (dates are set / on the calendar).
      await new CseRequestService(sp).confirmDates(id, 'Scheduled');
      setRequests(prev => prev.map(r => r.id === id
        ? { ...r, scheduleStatus: 'Dates Confirmed', requestStatus: 'Scheduled' }
        : r
      ));
      setExpandedId(null);
    } finally { setSavingDates(false); }
  };

  // SSE accepts a request that has NO proposed dates (TBD) → goes straight to In Progress.
  const handleSseAccept = async (id: number): Promise<void> => {
    setSavingId(id);
    try {
      await new CseRequestService(sp).updateStatus(id, 'In Progress');
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: 'In Progress' } : r));
      setExpandedId(null);
    } finally { setSavingId(null); }
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

  const handleDeleteSelected = async (): Promise<void> => {
    if (selectedIds.size === 0) return;
    const confirmed = window.confirm(`Permanently delete ${selectedIds.size} request${selectedIds.size > 1 ? 's' : ''}? This cannot be undone.`);
    if (!confirmed) return;
    const svc = new CseRequestService(sp);
    await Promise.all(Array.from(selectedIds).map(id => svc.delete(id).catch(() => undefined)));
    setRequests(prev => prev.filter(r => !selectedIds.has(r.id!)));
    setSelectedIds(new Set());
  };

  const handleStatusChange = async (id: number, status: CseRequestStatus): Promise<void> => {
    try {
      await new CseRequestService(sp).updateStatus(id, status);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, requestStatus: status } : r));
    } catch { /* ignore */ }
  };

  const handleCustTemp = async (id: number, temp: string): Promise<void> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const t = temp as any;
    try {
      await new CseRequestService(sp).updateCustTemp(id, t);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, custTemp: t } : r));
    } catch { /* ignore */ }
  };

  const handlePriority = async (id: number, priority: string): Promise<void> => {
    try {
      await new CseRequestService(sp).updatePriority(id, priority);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, csePriority: priority } : r));
    } catch { /* ignore */ }
  };

  const handleSolutionSave = async (id: number): Promise<void> => {
    const codes = Array.from(editSolCodes).join(',');
    try {
      await new CseRequestService(sp).updateSolutions(id, codes);
      setRequests(prev => prev.map(r => r.id === id ? { ...r, solutionsFocus: codes } : r));
      setEditSolId(null);
      setEditSolCodes(new Set());
    } catch { /* ignore */ }
  };

  const handleTileClick = (label: string): void => {
    const next = activeTile === label ? null : label;
    setActiveTile(next);
    if (next !== null) {
      setTimeout(() => tableRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 60);
    }
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
    // A user can cancel (soft-delete) their OWN work — as the submitter (SEPrimary) or the assigned SSE (RequestedCSE).
    const isSubmitter   = req.sePrimary.toLowerCase().includes(userEmail);
    const isAssignedSse = (req.requestedCse || '').toLowerCase().includes(userEmail);
    return isSubmitter || isAssignedSse;
  };

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading SRT Dashboard…" style={{ marginTop: 40 }} />;

  if (error) return (
    <div style={{ padding: 24, color: '#a4262c', background: '#fde7e9', borderRadius: 6, margin: 16 }}>
      <strong>Error loading requests:</strong> {error}
    </div>
  );

  // SEE vs DO: everyone can see the board (read-only); actions stay role-gated elsewhere.
  // 'mine' = requests I submitted (SEPrimary) OR am assigned to (RequestedCSE); 'all' = full board.
  const relevantToMe = (r: ICseRequest): boolean =>
    r.sePrimary.toLowerCase().includes(userEmail) || (r.requestedCse || '').toLowerCase().includes(userEmail);
  const notHidden = (r: ICseRequest): boolean => isAdmin || r.requestStatus !== 'Cancelled';
  const visibleRequests = requests.filter(r => notHidden(r) && (viewMode === 'all' || relevantToMe(r)));
  const mineCount = requests.filter(r => notHidden(r) && relevantToMe(r)).length;
  const allCount  = requests.filter(notHidden).length;

  const buOptions     = Array.from(new Set(visibleRequests.map(r => r.hpenBusinessUnit).filter(Boolean))).sort();
  const regionOptions = Array.from(new Set(visibleRequests.map(r => r.buRegion).filter(Boolean))).sort();

  const tileFiltered = (() => {
    if (!activeTile) return visibleRequests;
    if (activeTile === 'Awaiting SED')   return visibleRequests.filter(r => r.requestStatus === 'Pending');
    if (activeTile === 'Needs Info')     return visibleRequests.filter(r => r.requestStatus === 'Needs Info');
    if (activeTile === 'Awaiting SSE')   return visibleRequests.filter(r => r.requestStatus === 'Accepted');
    if (activeTile === 'Parked')         return visibleRequests.filter(r => r.requestStatus === 'Parked');
    if (activeTile === 'Active')         return visibleRequests.filter(r => ['Scheduled', 'In Progress'].includes(r.requestStatus));
    if (activeTile === 'Complete')       return visibleRequests.filter(r => r.requestStatus === 'Complete');
    if (activeTile === 'Needs Sign-off') return visibleRequests.filter(r => r.requestStatus === 'Complete' && !r.signedOffBy);
    if (activeTile === 'Declined')       return visibleRequests.filter(r => r.requestStatus === 'Declined' || r.requestStatus === 'Cancelled');
    return visibleRequests;
  })();

  const isStrategicReq = (r: ICseRequest): boolean => r.engagementType === 'Strategic Engagement';
  const reqType = (r: ICseRequest): string => isStrategicReq(r) ? 'Strategic' : 'POC';
  // Shared filter-bar predicate — applied to EVERY section, not just Active.
  const matchesFilters = (r: ICseRequest): boolean => {
    if (filterStatus !== 'All' && r.requestStatus !== filterStatus) return false;
    if (filterBU !== 'All' && r.hpenBusinessUnit !== filterBU) return false;
    if (filterRegion !== 'All' && r.buRegion !== filterRegion) return false;
    if (filterPriority !== 'All' && r.csePriority !== filterPriority) return false;
    if (filterType !== 'All' && reqType(r) !== filterType) return false;
    if (filterSearch && !r.customerName.toLowerCase().includes(filterSearch.toLowerCase())) return false;
    return true;
  };

  const pendingRows   = visibleRequests.filter(r => r.requestStatus === 'Pending' && matchesFilters(r));
  const acceptedRows  = visibleRequests.filter(r => r.requestStatus === 'Accepted' && matchesFilters(r));
  const parkedRows    = visibleRequests.filter(r => r.requestStatus === 'Parked' && matchesFilters(r));
  const completedRows = visibleRequests.filter(r => r.requestStatus === 'Complete' && matchesFilters(r));
  const declinedRows  = visibleRequests.filter(r => (r.requestStatus === 'Declined' || r.requestStatus === 'Cancelled') && matchesFilters(r));

  const filteredRequests = tileFiltered.filter(r => {
    if (!activeTile && r.requestStatus === 'Pending') return false;
    if (!activeTile && r.requestStatus === 'Accepted') return false;
    if (!activeTile && r.requestStatus === 'Parked') return false;
    if (!activeTile && r.requestStatus === 'Complete') return false;
    if (!activeTile && (r.requestStatus === 'Declined' || r.requestStatus === 'Cancelled')) return false;
    return matchesFilters(r);
  }).sort((a, b) => {
    const PRIORITY_ORDER: Record<string, number> = { 'High': 0, 'Medium': 1, 'Low': 2, 'Critical': -1 };
    const STATUS_ORDER: Record<string, number> = { 'Pending': 0, 'Needs Info': 1, 'Accepted': 2, 'Scheduled': 3, 'In Progress': 4, 'Complete': 5, 'Declined': 6, 'Cancelled': 7 };
    let cmp = 0;
    if (sortField === 'customer') cmp = a.customerName.localeCompare(b.customerName);
    else if (sortField === 'priority') cmp = (PRIORITY_ORDER[a.csePriority] ?? 99) - (PRIORITY_ORDER[b.csePriority] ?? 99);
    else if (sortField === 'status') cmp = (STATUS_ORDER[a.requestStatus] ?? 99) - (STATUS_ORDER[b.requestStatus] ?? 99);
    else if (sortField === 'type') cmp = (isStrategicReq(a) ? 0 : 1) - (isStrategicReq(b) ? 0 : 1);
    return sortDir === 'asc' ? cmp : -cmp;
  });

  const pending      = pendingRows;
  const needsInfo    = visibleRequests.filter(r => r.requestStatus === 'Needs Info');
  const awaitingSse  = acceptedRows;
  const active       = visibleRequests.filter(r => r.requestStatus === 'Scheduled' || r.requestStatus === 'In Progress');
  const parked       = parkedRows;
  const complete     = visibleRequests.filter(r => r.requestStatus === 'Complete');
  const needsSignOff = complete.filter(r => !r.signedOffBy);
  const declined     = declinedRows;

  // Actions column shows for SEDs, admins, or an assigned SSE who has any pending action
  // (accept a dateless engagement, or confirm proposed/rescheduled dates on an active one).
  const sseHasActionOn = (r: ICseRequest): boolean => {
    if (!(r.requestedCse || '').toLowerCase().includes(userEmail)) return false;
    const accTbd  = r.requestStatus === 'Accepted' && r.scheduleStatus === 'TBD';
    const needConf = ['Accepted', 'Scheduled', 'In Progress'].indexOf(r.requestStatus) !== -1
                     && (r.scheduleStatus === 'Dates Proposed' || r.scheduleStatus === 'Rescheduling');
    return accTbd || needConf;
  };
  const showActionsCol = isSED || isAdmin || visibleRequests.some(sseHasActionOn);

  const renderRow = (req: ICseRequest, i: number, hideReassign = false): JSX.Element => {
    const statusStyle   = CSE_STATUS_STYLE[req.requestStatus]  || CSE_STATUS_STYLE.Pending;
    // A "Dates Proposed/Confirmed" flag is dishonest if no actual date was entered — show TBD until there is one.
    const hasRealDates  = (!req.remoteTbd && !!req.remoteStart) || (!req.onsiteTbd && !!req.onsiteStart);
    const effSchedStatus: ScheduleStatus = (!hasRealDates && (req.scheduleStatus === 'Dates Proposed' || req.scheduleStatus === 'Dates Confirmed')) ? 'TBD' : req.scheduleStatus;
    const schedStyle    = SCHEDULE_STATUS_STYLE[effSchedStatus] || SCHEDULE_STATUS_STYLE.TBD;
    const tempStyle     = CUST_TEMP_STYLE[req.custTemp]    || CUST_TEMP_STYLE.Normal;
    const sseName       = parseSseName(req.requestedCse.includes('/') ? req.requestedCse.split('/')[0].trim() : req.requestedCse);
    const isExpanded    = expandedId === req.id;
    const isCancelled   = req.requestStatus === 'Cancelled';
    const isAssignedSse   = req.requestedCse?.toLowerCase().includes(userEmail);
    const canEditDates  = (isAdmin || req.sePrimary.toLowerCase().includes(userEmail) || isAssignedSse)
                          && !['Declined', 'Complete', 'Cancelled'].includes(req.requestStatus);
    const noDates        = req.scheduleStatus === 'TBD';
    // SSE gets an action to (a) accept a dateless engagement, or (b) confirm dates proposed/rescheduled
    // AT ANY point after SED acceptance — including when the SE adds/changes dates on an already-active engagement.
    const acceptTbd      = req.requestStatus === 'Accepted' && noDates;
    const needsDateConfirm = ['Accepted', 'Scheduled', 'In Progress'].indexOf(req.requestStatus) !== -1
                          && (req.scheduleStatus === 'Dates Proposed' || req.scheduleStatus === 'Rescheduling');
    const showSseActions = (isAssignedSse || isAdmin) && (acceptTbd || needsDateConfirm);
    const colSpan       = 12 + (showActionsCol ? 1 : 0) + (isAdmin ? 1 : 0);
    const isStrat       = isStrategicReq(req);
    const purposeText   = req.engagementPurpose === 'Other' ? (req.engagementPurposeOther || 'Other') : (req.engagementPurpose || '');
    const durationText  = req.remoteDuration || req.onsiteDuration || '';
    const rowBg         = selectedIds.has(req.id!) ? '#ffe9e9' : isExpanded ? '#f0ebff' : isCancelled ? '#f8f8f8' : i % 2 === 0 ? '#fff' : '#faf9f8';
    return (
      <React.Fragment key={req.id ?? i}>
      <tr style={{ background: rowBg, borderBottom: isExpanded ? 'none' : '1px solid #edebe9', cursor: 'pointer' }}
          onClick={() => handleExpand(req)}>
        {isAdmin && (
          <td style={{ ...TD, width: 28, padding: '4px 6px' }} onClick={e => e.stopPropagation()}>
            <input type="checkbox" checked={selectedIds.has(req.id!)}
              onChange={e => setSelectedIds(prev => {
                const next = new Set(prev);
                if (e.target.checked) { next.add(req.id!); } else { next.delete(req.id!); }
                return next;
              })} />
          </td>
        )}
        <td style={{ ...TD, opacity: isCancelled ? 0.6 : 1 }}>
          <span style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 700, whiteSpace: 'nowrap',
            background: isStrat ? '#e8f5e9' : '#ebf3fc', color: isStrat ? '#1a6b2e' : HPE_NAVY }}>
            {isStrat ? '🎯 Strategic' : '🔬 POC'}
          </span>
          {isStrat && (purposeText || durationText) && (
            <div style={{ fontSize: 10, color: '#888', marginTop: 3, maxWidth: 150 }}>
              {purposeText}{purposeText && durationText ? ' · ' : ''}{durationText}
            </div>
          )}
        </td>
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
        <td style={TD}>
          <div>{sseName || '—'}</div>
          {req.sedEmail && <div style={{ fontSize: 10, color: '#888', marginTop: 2 }}>{emailToName(req.sedEmail)}</div>}
        </td>
        <td style={TD}>
          <div>{req.hpenBusinessUnit}</div>
          <div style={{ fontSize: 11, color: '#888' }}>{req.buRegion}</div>
        </td>
        <td style={{ ...TD, fontSize: 11, maxWidth: 160, cursor: isAdmin && editSolId !== req.id ? 'pointer' : 'default' }}
            title={isAdmin && editSolId !== req.id ? 'Click to edit solutions' : undefined}
            onClick={e => { e.stopPropagation(); if (isAdmin && editSolId !== req.id) { setEditSolId(req.id!); setEditSolCodes(new Set(req.solutionsFocus ? req.solutionsFocus.split(',').map(c => c.trim()).filter(Boolean) : [])); } }}>
          {isAdmin && editSolId === req.id ? (
            <div onClick={e => e.stopPropagation()}>
              <div style={{ maxHeight: 180, overflowY: 'auto', border: '1px solid #ddd', borderRadius: 4, padding: '4px 6px', background: '#fff', marginBottom: 6 }}>
                {SOLUTION_CATEGORIES.map((cat, ci) => (
                  <div key={cat}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: '#888', textTransform: 'uppercase', padding: `${ci === 0 ? 2 : 6}px 0 2px`, borderTop: ci === 0 ? 'none' : '1px solid #f0f0f0' }}>{cat}</div>
                    {SOLUTIONS.filter(s => s.category === cat).map(s => (
                      <label key={s.code} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, cursor: 'pointer', padding: '2px 0' }}>
                        <input type="checkbox" checked={editSolCodes.has(s.code)}
                          onChange={ev => { const next = new Set(editSolCodes); if (ev.target.checked) next.add(s.code); else next.delete(s.code); setEditSolCodes(next); }} />
                        {s.name}
                      </label>
                    ))}
                  </div>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 4 }}>
                <button onClick={() => handleSolutionSave(req.id!).catch(() => undefined)}
                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>Save</button>
                <button onClick={() => { setEditSolId(null); setEditSolCodes(new Set()); }}
                  style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
              </div>
            </div>
          ) : (
            <span style={{ textDecoration: isAdmin ? 'underline dotted' : 'none' }}>{codeToName(req.solutionsFocus)}</span>
          )}
        </td>
        <td style={TD} onClick={e => e.stopPropagation()}>
          {(() => {
            const isHi = req.csePriority === 'High' || req.csePriority === 'Critical';
            const bg = isHi ? '#fde7e9' : req.csePriority === 'Medium' ? '#fff4ce' : '#e8faf3';
            const color = isHi ? '#a4262c' : req.csePriority === 'Medium' ? '#8a6000' : '#107c10';
            return (isAdmin || isAssignedSse) ? (
              <select value={req.csePriority || 'Medium'} onChange={e => handlePriority(req.id!, e.target.value).catch(() => undefined)}
                style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 700, border: 'none', cursor: 'pointer', background: bg, color }}>
                {['Low', 'Medium', 'High', 'Critical'].map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            ) : (
              <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 700, background: bg, color }}>{req.csePriority || '—'}</span>
            );
          })()}
        </td>
        <td style={TD} onClick={e => e.stopPropagation()}>
          {(isAdmin || isAssignedSse) ? (
            <select value={req.requestStatus}
              onChange={e => handleStatusChange(req.id!, e.target.value as CseRequestStatus).catch(() => undefined)}
              style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600, background: statusStyle.bg, color: statusStyle.color, border: 'none', cursor: 'pointer' }}>
              {(['Pending','Accepted','Scheduled','In Progress','Parked','Complete','Declined','Needs Info','Cancelled'] as CseRequestStatus[]).map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          ) : (
            <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600, background: statusStyle.bg, color: statusStyle.color }}>{req.requestStatus}</span>
          )}
        </td>
        <td style={TD}>
          <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600, background: schedStyle.bg, color: schedStyle.color }}>{effSchedStatus}</span>
          {!req.remoteTbd && req.remoteStart && <div style={{ fontSize: 10, color: '#555', marginTop: 3 }}>Remote: {fmtDate(req.remoteStart)}{req.remoteEnd ? ` – ${fmtDate(req.remoteEnd)}` : ''}</div>}
          {!req.onsiteTbd && req.onsiteStart && <div style={{ fontSize: 10, color: '#555', marginTop: 1 }}>Onsite: {fmtDate(req.onsiteStart)}{req.onsiteEnd ? ` – ${fmtDate(req.onsiteEnd)}` : ''}</div>}
        </td>
        <td style={TD} onClick={e => e.stopPropagation()}>
          {(isAdmin || isAssignedSse) ? (
            <select value={req.custTemp}
              onChange={e => handleCustTemp(req.id!, e.target.value).catch(() => undefined)}
              style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600, background: tempStyle.bg, color: tempStyle.color, border: 'none', cursor: 'pointer' }}>
              {(['Low','Normal','High','Critical']).map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          ) : (
            <span style={{ padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600, background: tempStyle.bg, color: tempStyle.color }}>{req.custTemp}</span>
          )}
        </td>
        <td style={TD} onClick={e => e.stopPropagation()}>
          {req.signedOffBy ? (
            <div style={{ fontSize: 11, color: '#107c10' }}>✓ {req.signedOffBy}</div>
          ) : req.requestStatus === 'Complete' ? (
            isAdmin && signOffId === req.id ? (
              <div>
                <input type="text" value={signOffName} onChange={e => setSignOffName(e.target.value)} placeholder="Your name" autoFocus
                  style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }} />
                <div style={{ display: 'flex', gap: 4 }}>
                  <button disabled={!signOffName.trim() || savingId === req.id} onClick={() => handleSignOff(req.id!).catch(() => undefined)}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                    {savingId === req.id ? '…' : '✓ Confirm'}
                  </button>
                  <button onClick={() => { setSignOffId(null); setSignOffName(''); }}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : isAdmin ? (
              <button onClick={() => { setSignOffId(req.id!); setSignOffName(''); }}
                style={{ fontSize: 11, padding: '3px 10px', background: '#e8faf3', color: '#107c10', border: '1px solid #107c10', borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>Sign Off</button>
            ) : (
              <span style={{ fontSize: 11, color: '#8a6000' }}>Pending</span>
            )
          ) : (
            <span style={{ fontSize: 11, color: '#aaa' }}>—</span>
          )}
        </td>
        <td style={{ ...TD, fontSize: 11, color: '#888', whiteSpace: 'nowrap' }}>{fmtDate(req.modified || '') || '—'}</td>
        {showActionsCol && <td style={{ ...TD, minWidth: 160 }} onClick={e => e.stopPropagation()}>
          {req.requestStatus === 'Pending' && (isSED || isAdmin) && (
            decliningId === req.id ? (
              <div>
                <input type="text" value={declineNote} onChange={e => setDeclineNote(e.target.value)} placeholder="Reason (optional)"
                  style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }} />
                <div style={{ display: 'flex', gap: 4 }}>
                  <button disabled={savingId === req.id} onClick={() => handleDeclineConfirm(req.id!)}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#a4262c', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                    {savingId === req.id ? '…' : 'Confirm'}
                  </button>
                  <button onClick={() => { setDecliningId(null); setDeclineNote(''); }}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : needsInfoId === req.id ? (
              <div>
                <input type="text" value={needsInfoNote} onChange={e => setNeedsInfoNote(e.target.value)} placeholder="What info is needed?"
                  style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }} />
                <div style={{ display: 'flex', gap: 4 }}>
                  <button disabled={savingId === req.id} onClick={() => handleNeedsInfoConfirm(req.id!)}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#6b2faf', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                    {savingId === req.id ? '…' : 'Send'}
                  </button>
                  <button onClick={() => { setNeedsInfoId(null); setNeedsInfoNote(''); }}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : !hideReassign && reassignId === req.id ? (
              <div>
                <div style={{ fontSize: 10, color: '#605e5c', marginBottom: 3 }}>Enter new SSE as &ldquo;Name / email&rdquo;</div>
                <input type="text" value={reassignSse} onChange={e => setReassignSse(e.target.value)}
                  placeholder="First Last / email@hpe.com" autoFocus
                  style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }} />
                <div style={{ display: 'flex', gap: 4 }}>
                  <button disabled={savingId === req.id || !reassignSse.includes('/')} onClick={() => handleReassignConfirm(req.id!)}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#0078d4', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                    {savingId === req.id ? '…' : 'Reassign'}
                  </button>
                  <button onClick={() => { setReassignId(null); setReassignSse(''); }}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                <button disabled={savingId === req.id} onClick={() => handleAccept(req.id!)}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                  {savingId === req.id ? '…' : '✓ Accept'}
                </button>
                {!hideReassign && (
                  <button onClick={() => { setReassignId(req.id!); setReassignSse(''); }}
                    style={{ fontSize: 11, padding: '4px 10px', background: '#eff6fc', color: '#0078d4', border: '1px solid #0078d4', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                    ↔ Reassign
                  </button>
                )}
                <button onClick={() => { setNeedsInfoId(req.id!); setNeedsInfoNote(''); }}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#f0e6ff', color: '#6b2faf', border: '1px solid #6b2faf', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                  ? Info
                </button>
                <button onClick={() => { setDecliningId(req.id!); setDeclineNote(''); }}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#fde7e9', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                  ✕ Decline
                </button>
              </div>
            )
          )}
          {/* ── SSE action on an Accepted request (Charlie's turn) ── */}
          {showSseActions && (
            acceptTbd ? (
              <button disabled={savingId === req.id} onClick={() => handleSseAccept(req.id!).catch(() => undefined)}
                title="No dates were proposed — accept and start the engagement"
                style={{ fontSize: 11, padding: '4px 12px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>
                {savingId === req.id ? '…' : '✓ Accept'}
              </button>
            ) : declineDatesId === req.id ? (
              <div>
                <input type="text" value={declineDatesNote} onChange={e => setDeclineDatesNote(e.target.value)} placeholder="Reason (optional)"
                  style={{ width: '100%', fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, marginBottom: 4, boxSizing: 'border-box' }} />
                <div style={{ display: 'flex', gap: 4 }}>
                  <button disabled={savingDates} onClick={() => handleDeclineDatesConfirm(req.id!).catch(() => undefined)}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#a4262c', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>{savingDates ? '…' : 'Confirm'}</button>
                  <button onClick={() => { setDeclineDatesId(null); setDeclineDatesNote(''); }}
                    style={{ flex: 1, fontSize: 11, padding: '3px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 3, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                <button disabled={savingDates} onClick={() => handleConfirmDates(req.id!).catch(() => undefined)}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>✓ Confirm</button>
                <button onClick={() => setDeclineDatesId(req.id!)}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#fde7e9', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>✕ Decline</button>
                <button onClick={() => handleExpand(req)}
                  style={{ fontSize: 11, padding: '4px 10px', background: '#f0e6ff', color: '#6b2faf', border: '1px solid #6b2faf', borderRadius: 3, cursor: 'pointer', fontWeight: 600 }}>✎ New dates</button>
              </div>
            )
          )}
          {/* ── nothing to act on for this viewer/row ── */}
          {!(req.requestStatus === 'Pending' && (isSED || isAdmin)) && !showSseActions && (
            <span style={{ fontSize: 11, color: '#aaa' }}>—</span>
          )}
        </td>}
      </tr>

      {/* ── Expanded date drawer ── */}
      {isExpanded && dateEdit && (
        <tr style={{ background: '#f8f5ff', borderBottom: '2px solid #6b2faf' }}>
          <td colSpan={colSpan} style={{ padding: '16px 20px' }} onClick={e => e.stopPropagation()}>
            <div style={{ maxWidth: 1180 }}>
            {/* ── Strategic Engagement review (competitive intel for the SSE) ── */}
            {isStrat && (
              <div style={{ marginBottom: 20, paddingBottom: 16, borderBottom: '1px solid #e0d8f0' }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>🎯 Strategic Engagement — Review</div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 14 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: '#555', marginBottom: 4 }}>Engagement Purpose</div>
                    <div style={{ fontSize: 13, color: '#323130' }}>
                      {req.engagementPurpose || '—'}{req.engagementPurpose === 'Other' && req.engagementPurposeOther ? `: ${req.engagementPurposeOther}` : ''}
                    </div>
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: '#555', marginBottom: 4 }}>Desired Outcome(s)</div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                      {(req.desiredOutcome && req.desiredOutcome.length ? req.desiredOutcome : ['—']).map((o, oi) => (
                        <span key={oi} style={{ fontSize: 11, fontWeight: 600, background: '#f0e6ff', color: '#6b2faf', borderRadius: 10, padding: '2px 9px' }}>{o}</span>
                      ))}
                    </div>
                    {req.desiredOutcomeDetail && <div style={{ fontSize: 11, color: '#605e5c', marginTop: 5, fontStyle: 'italic' }}>{req.desiredOutcomeDetail}</div>}
                  </div>
                </div>
                {(() => {
                  const envRows = parseEnvRows(req.currentEnvironment);
                  if (envRows.length === 0) return <div style={{ fontSize: 12, color: '#888' }}>No solution landscape captured.</div>;
                  return (
                    <div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: '#555', marginBottom: 6 }}>Solution Landscape</div>
                      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                        <thead><tr>
                          <th style={LTH}>We Position</th><th style={LTH}>Current Vendor</th><th style={LTH}>Product / Version</th><th style={LTH}>Disposition</th><th style={LTH}>Detail</th>
                        </tr></thead>
                        <tbody>
                          {envRows.map((row, ri) => {
                            const badge = DISPOSITION_STYLE[row.disposition];
                            return (
                              <tr key={ri}>
                                <td style={{ ...LTD, fontWeight: 600 }}>{row.solution || '—'}</td>
                                <td style={LTD}>{row.vendor || '—'}</td>
                                <td style={LTD}>{[row.product, row.version].filter(Boolean).join(' / ') || '—'}</td>
                                <td style={LTD}>
                                  <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 8, whiteSpace: 'nowrap', background: badge ? badge.bg : '#f0f0f0', color: badge ? badge.color : '#888' }}>
                                    {badge ? badge.label : (row.disposition || '—')}
                                  </span>
                                </td>
                                <td style={{ ...LTD, color: '#605e5c' }}>{row.detail || '—'}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  );
                })()}
              </div>
            )}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>Remote Support</div>
                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, marginBottom: 8, cursor: canEditDates ? 'pointer' : 'default' }}>
                  <input type="checkbox" checked={dateEdit.remoteTbd} disabled={!canEditDates}
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
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10 }}>On-Site Support</div>
                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, marginBottom: 8, cursor: canEditDates ? 'pointer' : 'default' }}>
                  <input type="checkbox" checked={dateEdit.onsiteTbd} disabled={!canEditDates}
                    onChange={e => setDateEdit(prev => prev ? { ...prev, onsiteTbd: e.target.checked, onsiteStart: '', onsiteEnd: '' } : prev)} />
                  TBD — dates not yet set
                </label>
                {!dateEdit.onsiteTbd && (
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                    <div>
                      <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>Start</div>
                      <input type="date" value={dateEdit.onsiteStart} disabled={!canEditDates}
                        onChange={e => { const s = e.target.value; setDateEdit(prev => prev ? { ...prev, onsiteStart: s, onsiteDuration: calcOnsiteDays(s, prev.onsiteEnd) } : prev); }}
                        style={{ width: '100%', fontSize: 12, padding: '4px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' }} />
                    </div>
                    <div>
                      <div style={{ fontSize: 11, color: '#555', marginBottom: 2 }}>End</div>
                      <input type="date" value={dateEdit.onsiteEnd} disabled={!canEditDates}
                        onChange={e => { const en = e.target.value; setDateEdit(prev => prev ? { ...prev, onsiteEnd: en, onsiteDuration: calcOnsiteDays(prev.onsiteStart, en) } : prev); }}
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

            {/* Opportunity + Notes */}
            <div style={{ marginTop: 20, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 6 }}>Opportunity</div>
                <textarea value={drawerOpportunity} onChange={e => setDrawerOpportunity(e.target.value)} rows={4}
                  placeholder="Background, context, or opportunity details…"
                  style={{ width: '100%', fontSize: 12, padding: '6px 8px', border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box', resize: 'vertical', fontFamily: 'inherit' }} />
              </div>
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6b2faf', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 6 }}>Notes</div>
                <textarea value={drawerNotes} onChange={e => setDrawerNotes(e.target.value)} rows={4}
                  placeholder="Running notes, date-tagged updates…"
                  style={{ width: '100%', fontSize: 12, padding: '6px 8px', border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box', resize: 'vertical', fontFamily: 'inherit' }} />
              </div>
            </div>
            <div style={{ marginTop: 8, display: 'flex', justifyContent: 'flex-end' }}>
              <button disabled={savingNotes} onClick={() => handleSaveNotes(req.id!).catch(() => undefined)}
                style={{ padding: '5px 18px', background: '#6b2faf', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', opacity: savingNotes ? 0.6 : 1 }}>
                {savingNotes ? 'Saving…' : 'Save Notes'}
              </button>
            </div>

            {/* Action bar */}
            <div style={{ marginTop: 16, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', borderTop: '1px solid #e0d8f0', paddingTop: 12 }}>
              {canEditDates && (
                <button disabled={savingDates} onClick={() => handleSaveDates(req).catch(() => undefined)}
                  style={{ padding: '6px 18px', background: '#6b2faf', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', opacity: savingDates ? 0.6 : 1 }}>
                  {savingDates ? 'Saving…' : req.scheduleStatus === 'Dates Confirmed' ? 'Save (will flag as Rescheduling)' : 'Save Dates'}
                </button>
              )}
              <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
                {canCancel(req) && cancelId !== req.id && (
                  <button onClick={() => { setCancelId(req.id!); setCancelReason(''); setCancelNote(''); }}
                    style={{ padding: '5px 14px', background: '#fff', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Cancel Request</button>
                )}
                <button onClick={() => { setExpandedId(null); setDateEdit(null); setCancelId(null); }}
                  style={{ padding: '5px 14px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>Close</button>
              </div>
            </div>

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
                  <button disabled={!cancelReason || savingDates} onClick={() => handleCancelConfirm(req.id!).catch(() => undefined)}
                    style={{ padding: '6px 18px', background: cancelReason ? '#a4262c' : '#e0e0e0', color: cancelReason ? '#fff' : '#aaa', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: cancelReason ? 'pointer' : 'not-allowed' }}>
                    {savingDates ? 'Cancelling…' : 'Confirm Cancellation'}
                  </button>
                  <button onClick={() => { setCancelId(null); setCancelReason(''); setCancelNote(''); }}
                    style={{ padding: '6px 14px', background: '#fff', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>Never Mind</button>
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
            </div>
          </td>
        </tr>
      )}
      </React.Fragment>
    );
  };

  return (
    <div style={{ fontFamily: 'inherit', minHeight: 400, width: '100%' }}>

      {/* ── Admin-only TEST bar: view the dashboard as another user ── */}
      {realIsAdmin && (
        <div style={{ background: actAs ? '#8a6000' : '#3b3a39', color: '#fff', padding: '6px 16px', display: 'flex', alignItems: 'center', gap: 10, fontSize: 12, flexWrap: 'wrap' }}>
          <span style={{ fontWeight: 700, letterSpacing: '0.3px' }}>🧪 TEST MODE</span>
          {actAs ? (
            <>
              <span>👁 Viewing as <strong>{actAs}</strong> — role &amp; scope reflect this user (writes still happen as you).</span>
              <button type="button" onClick={resetActAs}
                style={{ marginLeft: 'auto', fontSize: 11, fontWeight: 700, background: '#fff', color: '#8a6000', border: 'none', borderRadius: 4, padding: '4px 12px', cursor: 'pointer' }}>
                ↺ Reset to me
              </button>
            </>
          ) : (
            <>
              <span style={{ color: 'rgba(255,255,255,0.75)' }}>View as:</span>
              <input value={actAsInput} onChange={e => setActAsInput(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') applyActAs(actAsInput); }}
                placeholder="David, Charlie, or user@hpe.com" style={{ fontSize: 12, padding: '3px 8px', border: '1px solid #888', borderRadius: 4, minWidth: 240 }} />
              <button type="button" onClick={() => applyActAs(actAsInput)}
                style={{ fontSize: 11, fontWeight: 700, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, padding: '5px 12px', cursor: 'pointer' }}>
                View as
              </button>
            </>
          )}
        </div>
      )}

      {showAdmin && <SrtAdminPanel sp={sp} context={context} onClose={() => setShowAdmin(false)} />}

      {/* URL action result banner */}
      {urlActionBanner && (
        <div style={{
          margin: '10px 20px 0', padding: '10px 14px', borderRadius: 4, fontSize: 13,
          background: urlActionBanner.type === 'ok' ? '#dff6dd' : '#fde7e9',
          color: urlActionBanner.type === 'ok' ? '#107c10' : '#a4262c',
          border: `1px solid ${urlActionBanner.type === 'ok' ? '#107c10' : '#a4262c'}`,
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10,
        }}>
          <span>{urlActionBanner.msg}</span>
          <button onClick={() => setUrlActionBanner(null)}
            style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 16, lineHeight: 1,
              color: urlActionBanner.type === 'ok' ? '#107c10' : '#a4262c', padding: '0 2px' }}>✕</button>
        </div>
      )}

      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>SRT Dashboard</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — SSE Support Request Tracker</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
          <span style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)', marginRight: 4 }}>{(() => { const h = new Date().getHours(); return h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening'; })()}, {userName}</span>
          {/* Create actions (green) */}
          <a href={`${FRONT_DOOR_URL}?form=strategic`} target="_blank" rel="noreferrer" style={HDR_GREEN}>+ New SSE Request</a>
          <a href={`${POC_HOME_URL}?new=1`} target="_blank" rel="noreferrer" style={HDR_GREEN}>+ New POC</a>
          <span style={{ width: 1, height: 20, background: 'rgba(255,255,255,0.25)' }} />
          {/* Navigation (outline) */}
          <a href={POC_HOME_URL} target="_blank" rel="noreferrer" style={HDR_OUTLINE}>POC Manager</a>
          <a href={INSIGHTS_URL} target="_blank" rel="noreferrer" style={HDR_OUTLINE}>📊 SSE Demand Insights</a>
          {isAdmin && <button onClick={() => setShowAdmin(true)} style={HDR_OUTLINE}>Admin</button>}
          <span style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', letterSpacing: '0.5px' }}>v{VERSION}</span>
        </div>
      </div>

      {/* KPI tiles */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 1fr)', gap: 12, padding: '16px 20px 0' }}>
        {[
          { label: 'Awaiting SED',   count: pending.length,       bg: '#edebe9', color: '#605e5c' },
          { label: 'Needs Info',     count: needsInfo.length,     bg: '#f0e6ff', color: '#6b2faf' },
          { label: 'Awaiting SSE',   count: awaitingSse.length,   bg: '#ffe8cc', color: '#b45309' },
          { label: 'Parked',         count: parked.length,        bg: '#e6e2f0', color: '#5b4b8a' },
          { label: 'Active',         count: active.length,        bg: '#eff6fc', color: '#0078d4' },
          { label: 'Complete',       count: complete.length,      bg: '#e8faf3', color: '#107c10' },
          { label: 'Needs Sign-off', count: needsSignOff.length,  bg: '#fff4ce', color: '#8a6000' },
          { label: 'Declined',       count: declined.length,      bg: '#fde7e9', color: '#a4262c' },
        ].map(tile => {
          const isActive = activeTile === tile.label;
          return (
            <div key={tile.label}
              onClick={() => handleTileClick(tile.label)}
              title={isActive ? `Click to clear ${tile.label} filter` : `Click to filter by ${tile.label}`}
              style={{
                background: isActive ? tile.color : tile.bg,
                borderRadius: 6,
                padding: '12px 16px',
                textAlign: 'center',
                cursor: 'pointer',
                border: `2px solid ${isActive ? tile.color : 'transparent'}`,
                boxShadow: isActive ? `0 2px 10px ${tile.color}55` : 'none',
                transition: 'background 0.15s, box-shadow 0.15s',
                userSelect: 'none' as const,
              }}>
              <div style={{ fontSize: 28, fontWeight: 700, color: isActive ? '#fff' : tile.color }}>{tile.count}</div>
              <div style={{ fontSize: 11, fontWeight: 600, color: isActive ? '#fff' : tile.color, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                {tile.label}{isActive ? ' ✕' : ''}
              </div>
            </div>
          );
        })}
      </div>

      {/* Upcoming Onsite panel */}
      {(() => {
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const upcoming = requests
          .filter(r => !r.onsiteTbd && r.onsiteStart && new Date(r.onsiteStart) >= today && !['Cancelled', 'Declined'].includes(r.requestStatus))
          .sort((a, b) => new Date(a.onsiteStart).getTime() - new Date(b.onsiteStart).getTime());
        if (upcoming.length === 0) return null;
        return (
          <div style={{ margin: '0 20px 12px', border: '1px solid #edebe9', borderRadius: 6, overflow: 'hidden' }}>
            <div
              onClick={() => setShowOnsitePanel(v => !v)}
              style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 14px', background: '#f0f9f4', borderBottom: showOnsitePanel ? '1px solid #edebe9' : 'none', cursor: 'pointer', userSelect: 'none' as const }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 3, height: 14, background: HPE_GREEN, borderRadius: 2 }} />
                <span style={{ fontSize: 11, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Upcoming Onsite Engagements
                </span>
                <span style={{ fontSize: 11, background: HPE_GREEN, color: '#fff', borderRadius: 10, padding: '1px 7px', fontWeight: 700 }}>
                  {upcoming.length}
                </span>
              </div>
              <span style={{ fontSize: 12, color: '#888' }}>{showOnsitePanel ? '▲' : '▼'}</span>
            </div>
            {showOnsitePanel && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
                {upcoming.map((r, i) => (
                  <div key={r.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '8px 14px', background: i % 2 === 0 ? '#fff' : '#faf9f8', borderTop: i === 0 ? 'none' : '1px solid #f0f0f0', fontSize: 12 }}>
                    <div style={{ fontWeight: 600, color: HPE_NAVY, minWidth: 130 }}>{parseSseName(r.requestedCse.split('/')[0]?.trim() || r.requestedCse)}</div>
                    <div style={{ color: '#107c10', fontWeight: 600, whiteSpace: 'nowrap' }}>
                      {fmtDate(r.onsiteStart)}{r.onsiteEnd && r.onsiteEnd !== r.onsiteStart ? ` – ${fmtDate(r.onsiteEnd)}` : ''}
                    </div>
                    {r.onsiteDestination && <div style={{ color: '#605e5c' }}>{r.onsiteDestination}</div>}
                    <div style={{ color: '#888', marginLeft: 'auto', fontStyle: 'italic' }}>{r.customerName}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {/* Request table */}
      <div ref={tableRef} style={{ padding: '16px 20px' }}>

        <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 10, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 6 }}>
          {activeTile ? activeTile : 'SSE Requests'} ({(activeTile ? filteredRequests : [...pendingRows, ...filteredRequests, ...completedRows]).length}{filteredRequests.length !== visibleRequests.length ? ` of ${visibleRequests.length}` : ''})
        </div>

        {/* Filter bar */}
        <div style={{ display: 'flex', gap: 8, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center' }}>
          <div style={{ display: 'inline-flex', border: `1px solid ${HPE_NAVY}`, borderRadius: 4, overflow: 'hidden' }}>
            {([['mine', `👤 Relevant to me (${mineCount})`], ['all', `🌐 All (${allCount})`]] as const).map(([mode, label]) => (
              <button key={mode} type="button" onClick={() => setViewMode(mode)}
                style={{ fontSize: 12, fontWeight: 600, padding: '5px 12px', border: 'none', cursor: 'pointer', whiteSpace: 'nowrap',
                  background: viewMode === mode ? HPE_NAVY : '#fff', color: viewMode === mode ? '#fff' : '#605e5c' }}>
                {label}
              </button>
            ))}
          </div>
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
          <select value={filterType} onChange={e => setFilterType(e.target.value)}
            style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }}>
            <option value="All">All Types</option>
            <option value="Strategic">🎯 Strategic</option>
            <option value="POC">🔬 POC</option>
          </select>
          {(activeTile || filterStatus !== 'All' || filterBU !== 'All' || filterRegion !== 'All' || filterPriority !== 'All' || filterType !== 'All' || filterSearch) && (
            <button onClick={() => { setActiveTile(null); setFilterStatus('All'); setFilterBU('All'); setFilterRegion('All'); setFilterPriority('All'); setFilterType('All'); setFilterSearch(''); }}
              style={{ fontSize: 11, padding: '5px 10px', background: '#f3f2f1', border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#605e5c' }}>
              Clear Filters
            </button>
          )}
          {isAdmin && selectedIds.size > 0 && (
            <button onClick={handleDeleteSelected}
              style={{ fontSize: 11, padding: '5px 10px', background: '#a4262c', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer', fontWeight: 600, marginLeft: 'auto' }}>
              🗑 Delete {selectedIds.size} selected
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
                  {(() => {
                    const sortIcon = (field: 'customer' | 'priority' | 'status' | 'type'): string =>
                      sortField !== field ? ' ⇅' : sortDir === 'asc' ? ' ▲' : ' ▼';
                    const handleSort = (field: 'customer' | 'priority' | 'status' | 'type'): void => {
                      if (sortField === field) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
                      else { setSortField(field); setSortDir('asc'); }
                    };
                    return (<>
                  {isAdmin && <th style={{ ...TH, width: 28, padding: '4px 6px' }} />}
                  <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('type')}>Type{sortIcon('type')}</th>
                  <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('customer')}>Customer{sortIcon('customer')}</th>
                  <th style={TH}>SE</th>
                  <th style={TH}>SSE / SED</th>
                  <th style={TH}>BU / Region</th>
                  <th style={TH}>Solutions</th>
                  <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('priority')}>Priority{sortIcon('priority')}</th>
                  <th style={{ ...TH, cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('status')}>Status{sortIcon('status')}</th>
                    </>);
                  })()}
                  <th style={TH}>Schedule</th>
                  <th style={TH}>Temp</th>
                  <th style={TH}>Sign-off</th>
                  <th style={TH}>Updated</th>
                  {showActionsCol && <th style={TH}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {/* ── Awaiting Acceptance — SED ── */}
                {!activeTile && pendingRows.length > 0 && (
                  <tr style={{ cursor: 'pointer' }} onClick={() => setShowPendingSection(s => !s)}>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#fff4ce', fontWeight: 700, color: '#8a6000', fontSize: 12, userSelect: 'none' as const }}>
                      ⏳ Awaiting Acceptance — SED ({pendingRows.length}) {showPendingSection ? '▾' : '▸'}
                    </td>
                  </tr>
                )}
                {!activeTile && showPendingSection && pendingRows.map((req, i) => renderRow(req, i))}

                {/* ── Pending Confirmation — SSE (David accepted; awaiting Charlie's dates) ── */}
                {!activeTile && acceptedRows.length > 0 && (
                  <tr style={{ cursor: 'pointer' }} onClick={() => setShowAcceptedSection(s => !s)}>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#ffe8cc', fontWeight: 700, color: '#b45309', fontSize: 12, userSelect: 'none' as const }}>
                      📋 Pending Confirmation — SSE ({acceptedRows.length}) {showAcceptedSection ? '▾' : '▸'}
                    </td>
                  </tr>
                )}
                {!activeTile && showAcceptedSection && acceptedRows.map((req, i) => renderRow(req, i))}

                {/* ── Parked — accepted but not ready to move forward ── */}
                {!activeTile && parkedRows.length > 0 && (
                  <tr style={{ cursor: 'pointer' }} onClick={() => setShowParkedSection(s => !s)}>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#e6e2f0', fontWeight: 700, color: '#5b4b8a', fontSize: 12, userSelect: 'none' as const }}>
                      ⏸️ Parked — Not Ready ({parkedRows.length}) {showParkedSection ? '▾' : '▸'}
                    </td>
                  </tr>
                )}
                {!activeTile && showParkedSection && parkedRows.map((req, i) => renderRow(req, i))}

                {/* ── Active Engagements ── */}
                {!activeTile && (
                  <tr>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#ebf3fc', fontWeight: 700, color: HPE_NAVY, fontSize: 12 }}>
                      Active Engagements ({filteredRequests.length})
                    </td>
                  </tr>
                )}
                {filteredRequests.map((req, i) => renderRow(req, i))}

                {/* ── Completed & Awaiting Sign-off ── */}
                {!activeTile && completedRows.length > 0 && (
                  <tr style={{ cursor: 'pointer' }} onClick={() => setShowCompletedSection(s => !s)}>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#e8f5e9', fontWeight: 700, color: '#1a6b2e', fontSize: 12, userSelect: 'none' as const }}>
                      ✓ Completed &amp; Awaiting Sign-off ({completedRows.length}){completedRows.filter(r => !r.signedOffBy).length > 0 ? ` · ${completedRows.filter(r => !r.signedOffBy).length} need sign-off` : ''} {showCompletedSection ? '▾' : '▸'}
                    </td>
                  </tr>
                )}
                {!activeTile && showCompletedSection && completedRows.map((req, i) => renderRow(req, i))}

                {/* ── Declined & Cancelled (archive, collapsed by default) ── */}
                {!activeTile && declinedRows.length > 0 && (
                  <tr style={{ cursor: 'pointer' }} onClick={() => setShowDeclinedSection(s => !s)}>
                    <td colSpan={isAdmin ? 13 : 12} style={{ padding: '6px 12px', background: '#f3f2f1', fontWeight: 700, color: '#a4262c', fontSize: 12, userSelect: 'none' as const }}>
                      ❌ Declined &amp; Cancelled ({declinedRows.length}) {showDeclinedSection ? '▾' : '▸'}
                    </td>
                  </tr>
                )}
                {!activeTile && showDeclinedSection && declinedRows.map((req, i) => renderRow(req, i))}

              </tbody>
            </table>
          </div>
        )}

      </div>

    </div>
  );
};
