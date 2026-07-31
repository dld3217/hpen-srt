import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { ConfigService, BURegionMap, IBUConfig, IRegionConfig, ISSETeam } from '../../../services/ConfigService';
import { ContactDirectoryService, IContact } from '../../../services/ContactDirectoryService';
import { CseRequestService } from '../../../services/CseRequestService';
import { ISolutionDef, SOLUTIONS } from '../../../models/ISolution';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISseRequestFormProps {
  sp: SPFI;
  context: WebPartContext;
}

interface ISseFormData {
  customerName: string;
  hpenBusinessUnit: string;
  buRegion: string;
  requestedSse: string;
  solutionsFocus: string;
  supportType: 'Remote' | 'On-Site' | 'Both' | '';
  remoteTbd: boolean;
  remoteStart: string;
  remoteEnd: string;
  remoteDuration: string;
  onsiteTbd: boolean;
  onsiteStart: string;
  onsiteEnd: string;
  onsiteDuration: string;
  onsiteDestination: string;
  opportunity: string;
  notes: string;
  csePriority: string;
  csePriorityReason: string;
  opportunityAmount: number;
  specialtyType: string;
  additionalResourceNeeded: boolean;
  additionalSse: string;
  additionalSpecialty: string;
  additionalSse2: string;
  additionalSpecialty2: string;
  extraSses: string[];
}

const EMPTY_FORM: ISseFormData = {
  customerName: '',
  hpenBusinessUnit: '',
  buRegion: '',
  requestedSse: '',
  solutionsFocus: '',
  supportType: '',
  remoteTbd: false,
  remoteStart: '',
  remoteEnd: '',
  remoteDuration: '',
  onsiteTbd: false,
  onsiteStart: '',
  onsiteEnd: '',
  onsiteDuration: '',
  onsiteDestination: '',
  opportunity: '',
  notes: '',
  csePriority: '',
  csePriorityReason: '',
  opportunityAmount: 0,
  specialtyType: '',
  additionalResourceNeeded: false,
  additionalSse: '',
  additionalSpecialty: '',
  additionalSse2: '',
  additionalSpecialty2: '',
  extraSses: [],
};

const VERSION = '1.0.63';

const greeting = (): string => {
  const h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
};

const firstName = (displayName: string): string => {
  if (!displayName) return '';
  if (displayName.includes(',')) return displayName.split(',')[1].trim().split(' ')[0];
  return displayName.split(' ')[0];
};

const emailToName = (email: string): string => {
  if (!email) return '';
  const local = email.split('@')[0];
  return local.split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
};

// ── Styles ────────────────────────────────────────────────────────────────────

const SECTION: React.CSSProperties = {
  background: '#fff', border: '1px solid #e0e0e0', borderRadius: 6,
  padding: '14px 18px', marginBottom: 12,
};
const FIELD_ROW: React.CSSProperties = { marginBottom: 10 };
const LABEL_STYLE: React.CSSProperties = {
  display: 'block', fontSize: 12, fontWeight: 600, color: '#323130', marginBottom: 4,
};
const INPUT: React.CSSProperties = {
  width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '6px 8px',
  border: '1px solid #ccc', borderRadius: 3, fontFamily: 'inherit',
};
const TOGGLE_BTN = (active: boolean): React.CSSProperties => ({
  padding: '5px 14px', fontSize: 12, fontWeight: 600, borderRadius: 4,
  border: `1px solid ${active ? HPE_NAVY : '#ccc'}`,
  background: active ? HPE_NAVY : '#fff',
  color: active ? '#fff' : '#605e5c',
  cursor: 'pointer',
});

// ── SectionHeader ─────────────────────────────────────────────────────────────

const SectionHeader: React.FC<{ title: string }> = ({ title }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 12, paddingBottom: 6, borderBottom: `2px solid ${HPE_GREEN}` }}>
    <div style={{ width: 3, height: 16, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
    <span style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px' }}>{title}</span>
  </div>
);

// ── PeoplePickerField ─────────────────────────────────────────────────────────

interface IGraphUser { displayName: string; mail: string; }

const PeoplePickerField: React.FC<{
  label: string;
  value: string;
  onChange: (val: string) => void;
  required?: boolean;
  searchUsers?: (query: string) => Promise<IGraphUser[]>;
}> = ({ label, value, onChange, required, searchUsers }) => {
  const [editing, setEditing]       = useState(false);
  const [localValue, setLocalValue] = useState('');
  const [results, setResults]       = useState<IGraphUser[]>([]);
  const [searching, setSearching]   = useState(false);
  const debounceRef  = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const containerRef = useRef<HTMLDivElement>(null);

  const slashIdx    = value.indexOf(' / ');
  const showDisplay = !editing && slashIdx !== -1;
  const namePart    = showDisplay ? value.substring(0, slashIdx) : '';
  const emailPart   = showDisplay ? value.substring(slashIdx + 3) : '';

  const startEdit = (): void => {
    setLocalValue(slashIdx !== -1 ? value.substring(0, slashIdx) : value);
    setResults([]);
    setEditing(true);
  };

  const selectPerson = (person: IGraphUser): void => {
    onChange(`${person.displayName} / ${person.mail}`);
    setEditing(false);
    setResults([]);
    setLocalValue('');
  };

  const handleLocalChange = (q: string): void => {
    setLocalValue(q);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (q.length < 3 || !searchUsers) { setResults([]); return; }
    debounceRef.current = setTimeout(() => {
      setSearching(true);
      searchUsers(q)
        .then(r => setResults(r))
        .catch(() => setResults([]))
        .finally(() => setSearching(false));
    }, 300);
  };

  const handleBlur = (e: React.FocusEvent<HTMLDivElement>): void => {
    if (containerRef.current?.contains(e.relatedTarget as Node)) return;
    if (!editing) return;
    onChange(localValue.trim());
    setEditing(false);
    setResults([]);
    setLocalValue('');
  };

  return (
    <div ref={containerRef} onBlur={handleBlur} style={FIELD_ROW}>
      <label style={LABEL_STYLE}>{label}{required && <span style={{ color: '#d13438' }}> *</span>}</label>
      {showDisplay ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', background: '#f3f2f1', borderRadius: 3, border: '1px solid #ccc' }}>
          <a href={`mailto:${emailPart}`} style={{ fontWeight: 600, fontSize: 13, color: '#0078d4', textDecoration: 'none' }}>{namePart}</a>
          <span style={{ fontSize: 11, color: '#605e5c' }}>{emailPart}</span>
          <span style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
            <button type="button" onClick={startEdit}
              style={{ fontSize: 11, color: '#0078d4', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✎ Change</button>
            <button type="button" onClick={() => onChange('')}
              style={{ fontSize: 11, color: '#a4262c', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✕</button>
          </span>
        </div>
      ) : (
        <div style={{ position: 'relative' }}>
          <input
            autoFocus={editing}
            value={localValue}
            onChange={e => handleLocalChange(e.target.value)}
            onFocus={() => { if (!editing) startEdit(); }}
            placeholder={`Search for ${label}...`}
            style={INPUT}
          />
          {searching && <span style={{ position: 'absolute', right: 8, top: 8, fontSize: 11, color: '#888' }}>Searching…</span>}
          {results.length > 0 && (
            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #ccc', borderRadius: 3, zIndex: 100, boxShadow: '0 4px 12px rgba(0,0,0,0.12)' }}>
              {results.map(r => (
                <div key={r.mail} tabIndex={0}
                  onMouseDown={() => selectPerson(r)}
                  onKeyDown={e => e.key === 'Enter' && selectPerson(r)}
                  style={{ padding: '7px 12px', cursor: 'pointer', fontSize: 13, borderBottom: '1px solid #f0f0f0' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f3f2f1')}
                  onMouseLeave={e => (e.currentTarget.style.background = '#fff')}>
                  <span style={{ fontWeight: 600 }}>{r.displayName}</span>
                  <span style={{ color: '#888', marginLeft: 8, fontSize: 11 }}>{r.mail}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// ── Schedule Block ────────────────────────────────────────────────────────────

const calcDuration = (start: string, end: string): string => {
  if (!start || !end) return '';
  const ms = new Date(end).getTime() - new Date(start).getTime();
  if (ms <= 0) return '';
  const totalMinutes = Math.round(ms / 60000);
  const days  = Math.floor(totalMinutes / (60 * 24));
  const hours = Math.floor((totalMinutes % (60 * 24)) / 60);
  const mins  = totalMinutes % 60;
  const parts: string[] = [];
  if (days  > 0) parts.push(`${days} day${days > 1 ? 's' : ''}`);
  if (hours > 0) parts.push(`${hours} hour${hours > 1 ? 's' : ''}`);
  if (mins  > 0 && days === 0) parts.push(`${mins} min`);
  return parts.join(' ');
};

const ScheduleBlock: React.FC<{
  label: string;
  tbd: boolean;
  start: string;
  end: string;
  duration: string;
  destination?: string;
  dateOnly?: boolean;
  onTbd: (v: boolean) => void;
  onStart: (v: string) => void;
  onEnd: (v: string) => void;
  onDuration: (v: string) => void;
  onDestination?: (v: string) => void;
}> = ({ label, tbd, start, end, duration, destination, dateOnly, onTbd, onStart, onEnd, onDuration, onDestination }) => {
  const toInput = (iso: string): string => {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return '';
      if (dateOnly) return d.toLocaleDateString('sv');
      return d.toLocaleDateString('sv') + 'T' + d.toLocaleTimeString('sv').substring(0, 5);
    } catch { return ''; }
  };
  const fromInput = (val: string): string => val ? new Date(val).toISOString() : '';
  return (
    <div style={{ background: '#f9f9f9', border: '1px solid #e8e8e8', borderRadius: 4, padding: '10px 14px', marginBottom: 10 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: '#555', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{label}</div>
        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, cursor: 'pointer', margin: 0 }}>
          <input type="checkbox" checked={tbd} onChange={e => onTbd(e.target.checked)} style={{ accentColor: HPE_NAVY }} />
          Dates TBD
        </label>
      </div>
      {!tbd && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto', gap: 8, alignItems: 'end', marginBottom: 8 }}>
          <div>
            <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Start</label>
            <input type={dateOnly ? 'date' : 'datetime-local'} value={toInput(start)}
              onChange={e => { const s = fromInput(e.target.value); onStart(s); if (s && end) onDuration(calcDuration(s, end)); }} style={INPUT} />
          </div>
          <div>
            <label style={{ ...LABEL_STYLE, fontSize: 11 }}>End</label>
            <input type={dateOnly ? 'date' : 'datetime-local'} value={toInput(end)}
              onChange={e => { const en = fromInput(e.target.value); onEnd(en); if (start && en) onDuration(calcDuration(start, en)); }} style={INPUT} />
          </div>
          <button type="button" onClick={() => { onTbd(true); onStart(''); onEnd(''); onDuration(''); }}
            title="Clear dates"
            style={{ height: 32, padding: '0 10px', fontSize: 16, lineHeight: 1, background: 'transparent',
              border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#605e5c', marginBottom: 1 }}>✕</button>
        </div>
      )}
      <div>
        <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Expected Duration / Effort</label>
        <input type="text" value={duration} onChange={e => onDuration(e.target.value)}
          placeholder="e.g. 2 hours, half day, 3 days" style={INPUT} />
      </div>
      {onDestination !== undefined && (
        <div style={{ marginTop: 8 }}>
          <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Destination</label>
          <input type="text" value={destination || ''} onChange={e => onDestination(e.target.value)}
            placeholder="e.g. Denver, CO" style={INPUT} />
        </div>
      )}
    </div>
  );
};

// ── Main Component ────────────────────────────────────────────────────────────

export const SseRequestForm: React.FC<ISseRequestFormProps> = ({ sp, context }) => {
  const [loading, setLoading]               = useState(true);
  const [buRegions, setBuRegions]           = useState<BURegionMap>({});
  const [solutions, setSolutions]           = useState<ISolutionDef[]>(SOLUTIONS);
  const [sseTeams, setSseTeams]             = useState<ISSETeam[]>([]);
  const [sedApprovalRequired, setSedApprovalRequired] = useState(true);
  const [formData, setFormData]             = useState<ISseFormData>(EMPTY_FORM);
  const [submitting, setSubmitting]         = useState(false);
  const [submitted, setSubmitted]           = useState(false);
  const [submitError, setSubmitError]       = useState('');
  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  // Contact Directory state
  const [contacts, setContacts]                   = useState<IContact[]>([]);
  const [contactsLoaded, setContactsLoaded]       = useState(false);
  const [contactsLoading, setContactsLoading]     = useState(false);
  const [dismissedIds, setDismissedIds]           = useState<Set<number>>(new Set());
  const [expandedIds, setExpandedIds]             = useState<Set<number>>(new Set());

  const userEmail       = context.pageContext.user.email;
  const userDisplayName = context.pageContext.user.displayName || userEmail;

  useEffect(() => {
    const configSvc = new ConfigService(sp);
    Promise.all([configSvc.getBURegions(), configSvc.getSolutions(), configSvc.getSSETeams(), configSvc.getSEDApprovalRequired()])
      .then(([bur, sols, teams, sedRequired]) => {
        setBuRegions(bur);
        setSolutions(sols);
        setSseTeams(teams);
        setSedApprovalRequired(sedRequired);
        const savedBU = localStorage.getItem('srt_bu') || '';
        const savedRegion = localStorage.getItem('srt_region') || '';
        if (savedBU && bur[savedBU]) {
          const validRegion = savedRegion && (bur[savedBU] as IBUConfig)?.regions?.[savedRegion] ? savedRegion : '';
          setFormData(prev => ({ ...prev, hpenBusinessUnit: savedBU, buRegion: validRegion }));
        }
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  const set = (field: keyof ISseFormData, value: ISseFormData[keyof ISseFormData]): void =>
    setFormData(prev => ({ ...prev, [field]: value }));

  const handleSearchUsers = async (query: string): Promise<IGraphUser[]> => {
    try {
      const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const response: any = await client.api('/users')
        .header('ConsistencyLevel', 'eventual')
        .search(`"displayName:${query.replace(/"/g, '')}"`)
        .select('displayName,mail')
        .top(10)
        .get();
      interface IGraphResult { displayName?: string; mail?: string; }
      return ((response.value || []) as IGraphResult[])
        .filter(u => !!u.mail)
        .map(u => ({ displayName: (u.displayName || '') as string, mail: (u.mail || '') as string }));
    } catch {
      return [];
    }
  };

  const handleFindResources = async (): Promise<void> => {
    setContactsLoading(true);
    setDismissedIds(new Set());
    try {
      const svc = new ContactDirectoryService(context);
      const all = await svc.getAll();
      setContacts(all);
      setContactsLoaded(true);
    } catch {
      setContactsLoaded(true);
    } finally {
      setContactsLoading(false);
    }
  };

  // Match category keyword to an SSE team name
  const matchTeam = (cat: string): string => {
    const lower = cat.toLowerCase();
    const found = sseTeams.find(t => {
      const tn = t.name.toLowerCase();
      if (lower.includes('mist') && tn.includes('mist')) return true;
      if ((lower.includes('datacenter') || lower.includes('data center') || lower.includes('dcn')) && (tn.includes('dcn') || tn.includes('data'))) return true;
      if ((lower.includes('routing') || lower.includes('ris')) && (tn.includes('routing') || tn.includes('ris'))) return true;
      if (lower.includes('sase') && tn.includes('sase')) return true;
      if (lower.includes('aiops') && tn.includes('aiops')) return true;
      return false;
    });
    return found?.name || '';
  };

  const validate = (): string[] => {
    const errs: string[] = [];
    if (!formData.customerName.trim())          errs.push('Customer Name is required.');
    if (!formData.hpenBusinessUnit)             errs.push('Business Unit is required.');
    if (!formData.buRegion)                     errs.push('Region is required.');
    if (!formData.requestedSse.includes('/'))   errs.push('Requested SSE is required — search and select a person.');
    if (!formData.supportType)                  errs.push('Support Type is required.');
    if (!formData.csePriority)                  errs.push('Priority is required.');
    if (!formData.notes.trim())                 errs.push('Notes is required.');
    return errs;
  };

  const handleSubmit = async (): Promise<void> => {
    const errs = validate();
    if (errs.length > 0) { setValidationErrors(errs); return; }
    setValidationErrors([]);
    setSubmitting(true);
    setSubmitError('');

    try {
      const buConfig        = buRegions[formData.hpenBusinessUnit] as IBUConfig | undefined;
      const buSedEmail      = buConfig?.sedEmail || '';
      const regionCfg: IRegionConfig = buConfig?.regions[formData.buRegion] || {};
      const semEmail        = regionCfg.semEmail || '';
      const activeTeam      = sseTeams.find(t => t.name === formData.specialtyType);
      const sedEmail        = activeTeam?.managerEmail || buSedEmail;
      const sseManagerEmail = activeTeam?.managerEmail || '';
      const schedStatus     = (formData.remoteTbd && formData.onsiteTbd) ? 'TBD' as const : 'Dates Proposed' as const;

      const basePayload = {
        source: 'SE Landing Page' as const,
        linkedPocId: 0,
        requestStatus: sedApprovalRequired ? 'Pending' as const : 'Accepted' as const,
        scheduleStatus: schedStatus,
        sseManagerEmail,
        cseDescription: formData.notes,
        csePriority: formData.csePriority,
        csePriorityReason: formData.csePriorityReason,
        solutionsFocus: formData.solutionsFocus,
        supportType: formData.supportType,
        remoteTbd: formData.remoteTbd,
        remoteStart: formData.remoteStart,
        remoteEnd: formData.remoteEnd,
        remoteDuration: formData.remoteDuration,
        onsiteTbd: formData.onsiteTbd,
        onsiteStart: formData.onsiteStart,
        onsiteEnd: formData.onsiteEnd,
        onsiteDuration: formData.onsiteDuration,
        onsiteDestination: formData.onsiteDestination,
        sePrimary: `${userDisplayName} / ${userEmail}`,
        semPrimary: semEmail,
        sedEmail,
        buRegion: formData.buRegion,
        hpenBusinessUnit: formData.hpenBusinessUnit,
        customerName: formData.customerName,
        pocName: '',
        opportunityAmount: formData.opportunityAmount,
        custTemp: 'Normal' as const,
        signedOffBy: '',
        signOffDate: '',
        opportunity: formData.opportunity,
        notes: formData.notes,
        specialtyType: formData.specialtyType,
      };

      const svc = new CseRequestService(sp);
      await svc.create({ ...basePayload, title: `${formData.customerName} — SSE Request`, requestedCse: formData.requestedSse });

      setSubmitted(true);
    } catch (err) {
      setSubmitError(`Submit failed: ${(err as Error)?.message || String(err)}`);
    } finally {
      setSubmitting(false);
    }
  };

  // Derived values
  const buKeys            = Object.keys(buRegions).sort();
  const regionKeys        = formData.hpenBusinessUnit
    ? Object.keys((buRegions[formData.hpenBusinessUnit] as IBUConfig)?.regions || {}).sort()
    : [];
  const focusCodes        = formData.solutionsFocus ? formData.solutionsFocus.split(',').map(c => c.trim()).filter(Boolean) : [];
  const activeBuConfig    = formData.hpenBusinessUnit ? (buRegions[formData.hpenBusinessUnit] as IBUConfig | undefined) : undefined;
  const reviewingSedEmail = activeBuConfig?.sedEmail || '';
  const activeTeam        = sseTeams.find(t => t.name === formData.specialtyType);
  const effectiveSedEmail = activeTeam?.managerEmail || reviewingSedEmail;

  const suggestedContacts: IContact[] = React.useMemo(() => {
    if (!contactsLoaded || !contacts.length || !focusCodes.length) return [];
    return contacts.filter(c => {
      if (dismissedIds.has(c.id)) return false;
      const cSols = c.pocSolutions.split(';').map(s => s.trim()).filter(Boolean);
      return focusCodes.some(s => cSols.includes(s));
    });
  }, [contacts, contactsLoaded, focusCodes.join(','), dismissedIds]);

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading…" style={{ marginTop: 32 }} />;

  if (submitted) {
    return (
      <div style={{ maxWidth: 680, margin: '40px auto', textAlign: 'center', padding: 32 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>✅</div>
        <div style={{ fontSize: 22, fontWeight: 700, color: HPE_NAVY, marginBottom: 8 }}>Request Sent</div>
        <div style={{ fontSize: 14, color: '#605e5c', marginBottom: 24 }}>
          Your SSE support request for <strong>{formData.customerName}</strong> has been submitted.
          {sedApprovalRequired
            ? ' Your SED/GM will review it and assign the SSE — typically within 24 hours.'
            : ' The SSE has been notified directly.'}
        </div>
        <button onClick={() => { setFormData(EMPTY_FORM); setSubmitted(false); setContacts([]); setContactsLoaded(false); setDismissedIds(new Set()); }}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 600, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
          Submit Another Request
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 700, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>

      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '8px 16px', borderRadius: '6px 6px 0 0', marginBottom: 12, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 22, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, lineHeight: 1.2 }}>SSE Support Request</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — Strategic Systems Engineer</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
          <a href="https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SRT-Resource-Dashboard.aspx"
            target="_blank" rel="noreferrer"
            style={{ padding: '4px 14px', background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.35)',
              borderRadius: 4, color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer', textDecoration: 'none', whiteSpace: 'nowrap' }}>
            SRT Dashboard ↗
          </a>
          <div style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>
            <div style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)' }}>{greeting()}, {firstName(userDisplayName)}</div>
            <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', letterSpacing: '0.5px' }}>v{VERSION}</div>
          </div>
        </div>
      </div>

      {validationErrors.length > 0 && (
        <div style={{ background: '#fde7e9', border: '1px solid #d13438', borderRadius: 4, padding: '10px 14px', marginBottom: 12 }}>
          {validationErrors.map((e, i) => <div key={i} style={{ fontSize: 13, color: '#a4262c' }}>• {e}</div>)}
        </div>
      )}
      {submitError && (
        <div style={{ background: '#fde7e9', border: '1px solid #d13438', borderRadius: 4, padding: '10px 14px', marginBottom: 12 }}>
          <div style={{ fontSize: 13, color: '#a4262c' }}>{submitError}</div>
        </div>
      )}

      {/* ── Section 1: Request Info ── */}
      <div style={SECTION}>
        <SectionHeader title="Request Info" />

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Customer Name <span style={{ color: '#d13438' }}>*</span></label>
          <input type="text" value={formData.customerName}
            onChange={e => set('customerName', e.target.value)}
            placeholder="e.g. Acme Corp" style={INPUT} />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Business Unit <span style={{ color: '#d13438' }}>*</span></label>
            <select value={formData.hpenBusinessUnit}
              onChange={e => {
                const newBU = e.target.value;
                setFormData(prev => ({ ...prev, hpenBusinessUnit: newBU, buRegion: '' }));
                if (newBU) localStorage.setItem('srt_bu', newBU); else localStorage.removeItem('srt_bu');
                localStorage.removeItem('srt_region');
              }}
              style={INPUT}>
              <option value="">— Select BU —</option>
              {buKeys.map(bu => <option key={bu} value={bu}>{bu}</option>)}
            </select>
          </div>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Region <span style={{ color: '#d13438' }}>*</span></label>
            <select value={formData.buRegion}
              onChange={e => {
                set('buRegion', e.target.value);
                if (e.target.value) localStorage.setItem('srt_region', e.target.value); else localStorage.removeItem('srt_region');
              }}
              disabled={!formData.hpenBusinessUnit} style={INPUT}>
              <option value="">— Select Region —</option>
              {regionKeys.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Priority <span style={{ color: '#d13438' }}>*</span></label>
            <div style={{ display: 'flex', gap: 6 }}>
              {['High', 'Medium', 'Low'].map(p => (
                <button key={p} type="button" onClick={() => set('csePriority', p)}
                  style={{ ...TOGGLE_BTN(formData.csePriority === p),
                    borderColor: p === 'High' ? '#d13438' : p === 'Medium' ? '#ca5010' : '#107c10',
                    background: formData.csePriority === p ? (p === 'High' ? '#d13438' : p === 'Medium' ? '#ca5010' : '#107c10') : '#fff',
                    color: formData.csePriority === p ? '#fff' : '#323130' }}>
                  {p}
                </button>
              ))}
            </div>
          </div>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Opportunity $</label>
            <input type="number" value={formData.opportunityAmount || ''}
              onChange={e => set('opportunityAmount', parseFloat(e.target.value) || 0)}
              placeholder="0" style={INPUT} />
          </div>
        </div>

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Priority Reason</label>
          <input type="text" value={formData.csePriorityReason}
            onChange={e => set('csePriorityReason', e.target.value)}
            placeholder="Why is this priority level needed?" style={INPUT} />
        </div>
      </div>

      {/* ── Section 2: Solutions Focus ── */}
      <div style={SECTION}>
        <SectionHeader title="Solutions Focus" />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 20px', marginBottom: 14 }}>
          {solutions.map(sol => {
            const checked = focusCodes.includes(sol.code);
            return (
              <label key={sol.code} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 13, cursor: 'pointer' }}>
                <input type="checkbox" checked={checked}
                  onChange={e => {
                    const next = e.target.checked ? [...focusCodes, sol.code] : focusCodes.filter(c => c !== sol.code);
                    set('solutionsFocus', next.join(','));
                  }}
                  style={{ accentColor: HPE_GREEN }} />
                {sol.name}
              </label>
            );
          })}
        </div>
        <button type="button"
          onClick={() => handleFindResources().catch(() => undefined)}
          disabled={focusCodes.length === 0 || contactsLoading}
          style={{ padding: '7px 20px', fontSize: 13, fontWeight: 700,
            background: focusCodes.length === 0 || contactsLoading ? '#ccc' : HPE_GREEN,
            color: '#fff', border: 'none', borderRadius: 4,
            cursor: focusCodes.length === 0 || contactsLoading ? 'not-allowed' : 'pointer' }}>
          {contactsLoading ? 'Searching…' : '🔍 Find Resources'}
        </button>
      </div>

      {/* ── Suggested Resources ── */}
      {contactsLoaded && (
        <div style={{ padding: '10px 14px', marginBottom: 12, background: '#f0f7f4', border: '1px solid #a5d6a7', borderRadius: 6 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: suggestedContacts.length > 0 ? 10 : 0 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#107c10' }}>
              {suggestedContacts.length > 0
                ? 'Suggested resources from Contact Directory:'
                : 'No matching resources found in Contact Directory for the selected solutions.'}
            </div>
            {suggestedContacts.length > 0 && (
              <button onClick={() => setDismissedIds(new Set())}
                style={{ fontSize: 11, background: 'none', border: '1px solid #107c10', borderRadius: 3, color: '#107c10', padding: '2px 8px', cursor: 'pointer', lineHeight: 1.4 }}>
                ↺ Reset
              </button>
            )}
          </div>
          {suggestedContacts.length > 0 && (
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              {suggestedContacts.map(c => {
                const isExpanded = expandedIds.has(c.id);
                const slot1Free = !formData.requestedSse.includes('/');
                const slot2Free = !formData.additionalSse.includes('/');
                const slot3Free = !formData.additionalSse2.includes('/');
                const available = [slot1Free && 'primary', slot2Free && 'secondary', slot3Free && 'tertiary'].filter(Boolean) as string[];
                const lower = (c.category || '').toLowerCase();
                const naturalSlot = lower.includes('mist') ? 'tertiary' : (lower.includes('datacenter') || lower.includes('data center') || lower.includes('dcn')) ? 'secondary' : 'primary';
                const applyRole = available.includes(naturalSlot) ? naturalSlot : (available[0] || 'primary');
                const roleLabel = applyRole === 'secondary' ? 'DC Specialist' : applyRole === 'tertiary' ? 'Mist Specialist' : (matchTeam(c.category) || 'Legacy HPE/Aruba');
                const displayName = c.name.replace(/\s*\(.*\)$/, '').trim();
                return (
                  <div key={c.id} style={{ background: '#fff', border: '1px solid #a5d6a7', borderRadius: 4, padding: '8px 10px', minWidth: 170, flex: '1 1 170px', maxWidth: 250, display: 'flex', flexDirection: 'column' }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: '#107c10', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 3 }}>{c.category}</div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: '#222', marginBottom: 4 }}>{displayName}</div>
                    <button onClick={() => setExpandedIds(prev => prev.has(c.id) ? new Set(Array.from(prev).filter(id => id !== c.id)) : new Set(Array.from(prev).concat(c.id)))}
                      style={{ alignSelf: 'flex-start', background: 'none', border: 'none', padding: 0, fontSize: 11, color: '#0078d4', cursor: 'pointer', marginBottom: isExpanded ? 6 : 0 }}>
                      ℹ Details {isExpanded ? '▴' : '▾'}
                    </button>
                    {isExpanded && (
                      <div style={{ fontSize: 11, color: '#444', marginBottom: 6, borderTop: '1px dashed #c8e6c9', paddingTop: 6, lineHeight: 1.6 }}>
                        {c.geography    && <div><strong>Territory:</strong> {c.geography}</div>}
                        {c.businessUnit && <div><strong>BU:</strong> {c.businessUnit}</div>}
                        {c.email        && <div style={{ wordBreak: 'break-all' }}><strong>Email:</strong> {c.email}</div>}
                        {c.phone        && <div><strong>Phone:</strong> {c.phone}</div>}
                      </div>
                    )}
                    {available.length > 0 ? (
                      <div style={{ marginTop: 'auto', paddingTop: 8 }}>
                        {slot1Free && (
                          <label style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11, cursor: 'pointer', marginBottom: 5, color: '#323130' }}>
                            <input type="checkbox" checked={false}
                              onChange={e => {
                                if (!e.target.checked) return;
                                const val = `${c.name} / ${c.email}`;
                                const teamName = matchTeam(c.category);
                                setFormData(prev => ({ ...prev, requestedSse: val, specialtyType: teamName || prev.specialtyType }));
                                setDismissedIds(prev => new Set(Array.from(prev).concat(c.id)));
                              }}
                              style={{ accentColor: HPE_GREEN }} />
                            Make Primary SSE
                          </label>
                        )}
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button onClick={() => {
                            const val = `${c.name} / ${c.email}`;
                            const teamName = matchTeam(c.category);
                            if (applyRole === 'primary') {
                              setFormData(prev => ({ ...prev, requestedSse: val, specialtyType: teamName || prev.specialtyType }));
                            } else if (applyRole === 'secondary') {
                              setFormData(prev => ({ ...prev, additionalSse: val, additionalSpecialty: teamName || prev.additionalSpecialty }));
                            } else {
                              setFormData(prev => ({ ...prev, additionalSse2: val, additionalSpecialty2: teamName || prev.additionalSpecialty2 }));
                            }
                            setDismissedIds(prev => new Set(Array.from(prev).concat(c.id)));
                          }}
                            style={{ flex: 1, fontSize: 12, fontWeight: 700, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 3, padding: '5px 0', cursor: 'pointer' }}>
                            ✓ {roleLabel}
                          </button>
                          <button onClick={() => setDismissedIds(prev => new Set(Array.from(prev).concat(c.id)))}
                            style={{ fontSize: 12, background: '#fff', color: '#999', border: '1px solid #ccc', borderRadius: 3, padding: '5px 8px', cursor: 'pointer' }}>✕</button>
                        </div>
                      </div>
                    ) : (
                      <div style={{ marginTop: 'auto', paddingTop: 8, fontSize: 11, color: '#888', fontStyle: 'italic' }}>All slots filled</div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* ── Section 3: SSE Selection ── */}
      <div style={SECTION}>
        <SectionHeader title="SSE Selection" />

        {effectiveSedEmail && (
          <div style={{ background: '#eff6fc', border: '1px solid #0078d4', borderRadius: 4, padding: '8px 12px', marginBottom: 12, fontSize: 12, color: '#0078d4' }}>
            <strong>{emailToName(effectiveSedEmail)}</strong> (SED/GM) will review this request before the SSE is notified.
          </div>
        )}

        <PeoplePickerField label="Primary SSE" required value={formData.requestedSse} onChange={v => set('requestedSse', v)} searchUsers={handleSearchUsers} />

        {/* Card-populated entries appear as they're applied */}
        {!!formData.additionalSse && (
          <PeoplePickerField
            label={formData.additionalSpecialty || 'DC Specialist'}
            value={formData.additionalSse} onChange={v => set('additionalSse', v)} searchUsers={handleSearchUsers} />
        )}
        {!!formData.additionalSse2 && (
          <PeoplePickerField
            label={formData.additionalSpecialty2 || 'Mist Specialist'}
            value={formData.additionalSse2} onChange={v => set('additionalSse2', v)} searchUsers={handleSearchUsers} />
        )}

        {/* Dynamic extra SSEs — each checkbox click adds one more picker */}
        {formData.extraSses.map((val, i) => (
          <PeoplePickerField key={i}
            label="Additional SSE"
            value={val}
            onChange={v => setFormData(prev => {
              const next = [...prev.extraSses];
              next[i] = v;
              return { ...prev, extraSses: next };
            })}
            searchUsers={handleSearchUsers} />
        ))}
        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, cursor: 'pointer', color: '#323130', marginTop: 6 }}>
          <input type="checkbox" checked={false}
            onChange={e => { if (e.target.checked) setFormData(prev => ({ ...prev, extraSses: [...prev.extraSses, ''] })); }}
            style={{ accentColor: HPE_GREEN }} />
          Additional SSEs
        </label>
      </div>

      {/* ── Section 4: Support Details ── */}
      <div style={SECTION}>
        <SectionHeader title="Support Details" />

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Opportunity</label>
          <textarea rows={2} value={formData.opportunity}
            onChange={e => set('opportunity', e.target.value)}
            placeholder="e.g. Upgrade WLAN, new campus deployment, competitive displacement…"
            style={{ ...INPUT, resize: 'vertical' }} />
        </div>

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Support Type <span style={{ color: '#d13438' }}>*</span></label>
          <div style={{ display: 'flex', gap: 6 }}>
            {(['Remote', 'On-Site', 'Both'] as const).map(t => (
              <button key={t} type="button" onClick={() => set('supportType', t)} style={TOGGLE_BTN(formData.supportType === t)}>{t}</button>
            ))}
          </div>
        </div>

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Notes <span style={{ color: '#d13438' }}>*</span></label>
          <textarea rows={4} value={formData.notes}
            onChange={e => set('notes', e.target.value)}
            placeholder="Describe what SSE help is needed, required skill set, and any relevant context…"
            style={{ ...INPUT, resize: 'vertical' }} />
        </div>

        {(formData.supportType === 'Remote' || formData.supportType === 'Both') && (
          <ScheduleBlock label="Remote Schedule" dateOnly
            tbd={formData.remoteTbd} start={formData.remoteStart} end={formData.remoteEnd} duration={formData.remoteDuration}
            onTbd={v => set('remoteTbd', v)} onStart={v => set('remoteStart', v)} onEnd={v => set('remoteEnd', v)} onDuration={v => set('remoteDuration', v)} />
        )}
        {(formData.supportType === 'On-Site' || formData.supportType === 'Both') && (
          <ScheduleBlock label="On-Site Schedule" dateOnly
            tbd={formData.onsiteTbd} start={formData.onsiteStart} end={formData.onsiteEnd} duration={formData.onsiteDuration} destination={formData.onsiteDestination}
            onTbd={v => set('onsiteTbd', v)} onStart={v => set('onsiteStart', v)} onEnd={v => set('onsiteEnd', v)} onDuration={v => set('onsiteDuration', v)} onDestination={v => set('onsiteDestination', v)} />
        )}
      </div>

      {/* Submit */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 4, marginBottom: 32 }}>
        <button disabled={submitting} onClick={handleSubmit}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 700,
            background: submitting ? '#ccc' : HPE_GREEN, color: '#fff',
            border: 'none', borderRadius: 4, cursor: submitting ? 'not-allowed' : 'pointer' }}>
          {submitting ? 'Sending…' : '✉ Send SSE Request'}
        </button>
        <span style={{ fontSize: 12, color: '#605e5c' }}>
          {sedApprovalRequired
            ? 'Notifies your SED/GM for review — SSE assigned after approval'
            : 'Routes directly to the SSE — SED approval bypassed'}
        </span>
      </div>

    </div>
  );
};
