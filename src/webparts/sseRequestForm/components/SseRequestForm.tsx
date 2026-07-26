import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { SPFI, PrincipalType, PrincipalSource } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/sputilities';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { ConfigService, BURegionMap, IBUConfig } from '../../../services/ConfigService';
import { CseRequestService } from '../../../services/CseRequestService';
import { ISolutionDef, SOLUTIONS } from '../../../models/ISolution';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISseRequestFormProps {
  sp: SPFI;
  context: WebPartContext;
}

interface ISseFormData {
  customerName: string;
  pocName: string;
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
  cseDescription: string;
  csePriority: string;
  csePriorityReason: string;
  opportunityAmount: number;
}

const EMPTY_FORM: ISseFormData = {
  customerName: '',
  pocName: '',
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
  cseDescription: '',
  csePriority: '',
  csePriorityReason: '',
  opportunityAmount: 0,
};

const VERSION = '1.0.23';

const greeting = (): string => {
  const h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
};

const firstName = (displayName: string): string => {
  if (!displayName) return '';
  // Handle "Last, First Middle" enterprise AD format
  if (displayName.includes(',')) return displayName.split(',')[1].trim().split(' ')[0];
  // Handle "First Last" standard format
  return displayName.split(' ')[0];
};

// ── Helpers ──────────────────────────────────────────────────────────────────


// ── Inline styles ─────────────────────────────────────────────────────────────

const SECTION: React.CSSProperties = {
  background: '#fff', border: '1px solid #e0e0e0', borderRadius: 6,
  padding: '16px 20px', marginBottom: 14,
};
const SECTION_TITLE: React.CSSProperties = {
  fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase',
  letterSpacing: '0.6px', marginBottom: 12, borderBottom: `2px solid ${HPE_GREEN}`,
  paddingBottom: 6,
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
    onChange(localValue.trim());
    setEditing(false);
    setResults([]);
    setLocalValue('');
  };

  return (
    <div ref={containerRef} onBlur={handleBlur} style={FIELD_ROW}>
      <label style={LABEL_STYLE}>{label}{required && <span style={{ color: '#d13438' }}> *</span>}</label>
      {showDisplay ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', background: '#f3f2f1', borderRadius: 3, border: '1px solid #ccc', cursor: 'pointer' }}
          onClick={startEdit}>
          <span style={{ fontWeight: 600, fontSize: 13 }}>{namePart}</span>
          <span style={{ fontSize: 11, color: '#605e5c' }}>{emailPart}</span>
          <span style={{ marginLeft: 'auto', fontSize: 11, color: '#0078d4' }}>✎ Change</span>
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
      <>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto', gap: 8, alignItems: 'end', marginBottom: 8 }}>
          <div>
            <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Start</label>
            <input type={dateOnly ? 'date' : 'datetime-local'} value={toInput(start)}
              onChange={e => {
                const newStart = fromInput(e.target.value);
                onStart(newStart);
                if (newStart && end) onDuration(calcDuration(newStart, end));
              }} style={INPUT} />
          </div>
          <div>
            <label style={{ ...LABEL_STYLE, fontSize: 11 }}>End</label>
            <input type={dateOnly ? 'date' : 'datetime-local'} value={toInput(end)}
              onChange={e => {
                const newEnd = fromInput(e.target.value);
                onEnd(newEnd);
                if (start && newEnd) onDuration(calcDuration(start, newEnd));
              }} style={INPUT} />
          </div>
          <button type="button" onClick={() => { onTbd(true); onStart(''); onEnd(''); onDuration(''); }}
            title="Clear dates"
            style={{ height: 32, padding: '0 10px', fontSize: 16, lineHeight: 1, background: 'transparent',
              border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#605e5c', marginBottom: 1 }}>
            ✕
          </button>
        </div>
      </>
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
  const [loading, setLoading]       = useState(true);
  const [buRegions, setBuRegions]   = useState<BURegionMap>({});
  const [solutions, setSolutions]   = useState<ISolutionDef[]>(SOLUTIONS);
  const [formData, setFormData]     = useState<ISseFormData>(EMPTY_FORM);
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted]   = useState(false);
  const [submitError, setSubmitError] = useState('');
  const [validationErrors, setValidationErrors] = useState<string[]>([]);

  const userEmail       = context.pageContext.user.email;
  const userDisplayName = context.pageContext.user.displayName || userEmail;

  useEffect(() => {
    const configSvc = new ConfigService(sp);
    Promise.all([configSvc.getBURegions(), configSvc.getSolutions()])
      .then(([bur, sols]) => {
        setBuRegions(bur);
        setSolutions(sols);
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
      const results = await sp.utility.searchPrincipals(
        query, PrincipalType.User, PrincipalSource.All, '', 10
      );
      return results
        .filter(r => r.Email)
        .map(r => ({ displayName: r.DisplayName, mail: r.Email }));
    } catch (err) {
      console.error('[SRT] People picker exception:', err);
      return [];
    }
  };

  const validate = (): string[] => {
    const errs: string[] = [];
    if (!formData.customerName.trim())   errs.push('Customer Name is required.');
    if (!formData.hpenBusinessUnit)       errs.push('Business Unit is required.');
    if (!formData.buRegion)               errs.push('Region is required.');
    if (!formData.requestedSse.includes('/')) errs.push('Requested SSE is required — search and select a person.');
    if (!formData.supportType)            errs.push('Support Type is required.');
    if (!formData.csePriority)            errs.push('Priority is required.');
    if (!formData.cseDescription.trim())  errs.push('Description of Need is required.');
    return errs;
  };

  const handleSubmit = async (): Promise<void> => {
    const errs = validate();
    if (errs.length > 0) { setValidationErrors(errs); return; }
    setValidationErrors([]);
    setSubmitting(true);
    setSubmitError('');

    try {
      const buConfig  = buRegions[formData.hpenBusinessUnit] as IBUConfig | undefined;
      const sedEmail  = buConfig?.sedEmail || '';
      const regionCfg = buConfig?.regions[formData.buRegion] || {};
      const semEmail  = regionCfg.semEmail || '';

      const svc = new CseRequestService(sp);
      await svc.create({
        title: `${formData.customerName} — SSE Request`,
        source: 'SE Landing Page',
        linkedPocId: 0,
        requestStatus: 'Pending',
        scheduleStatus: (formData.remoteTbd && formData.onsiteTbd) ? 'TBD' : 'Dates Proposed',
        requestedCse: formData.requestedSse,
        sseManagerEmail: '',
        cseDescription: formData.cseDescription,
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
        pocName: formData.pocName,
        opportunityAmount: formData.opportunityAmount,
        custTemp: 'Normal',
        signedOffBy: '',
        signOffDate: '',
        notes: '',
      });

      setSubmitted(true);
    } catch (err) {
      setSubmitError(`Submit failed: ${(err as Error)?.message || String(err)}`);
    } finally {
      setSubmitting(false);
    }
  };

  const buKeys     = Object.keys(buRegions).sort();
  const regionKeys = formData.hpenBusinessUnit
    ? Object.keys((buRegions[formData.hpenBusinessUnit] as IBUConfig)?.regions || {}).sort()
    : [];
  const focusCodes = formData.solutionsFocus ? formData.solutionsFocus.split(',').map(c => c.trim()).filter(Boolean) : [];

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading…" style={{ marginTop: 32 }} />;

  if (submitted) {
    return (
      <div style={{ maxWidth: 680, margin: '40px auto', textAlign: 'center', padding: 32 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>✅</div>
        <div style={{ fontSize: 22, fontWeight: 700, color: HPE_NAVY, marginBottom: 8 }}>Request Sent</div>
        <div style={{ fontSize: 14, color: '#605e5c', marginBottom: 24 }}>
          Your SSE support request for <strong>{formData.customerName}</strong> has been submitted. The SSE team will be notified shortly.
        </div>
        <button
          onClick={() => { setFormData(EMPTY_FORM); setSubmitted(false); }}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 600, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
          Submit Another Request
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 680, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>

      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '8px 16px', borderRadius: '6px 6px 0 0', marginBottom: 2, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 22, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, lineHeight: 1.2 }}>SSE Support Request</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — Strategic Systems Engineer</div>
          </div>
        </div>
        <div style={{ textAlign: 'right', flexShrink: 0, whiteSpace: 'nowrap' }}>
          <div style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)' }}>{greeting()}, {firstName(userDisplayName)}</div>
          <div style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', letterSpacing: '0.5px' }}>v{VERSION}</div>
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

      {/* Section 1 — Request Info */}
      <div style={SECTION}>
        <div style={SECTION_TITLE}>Request Info</div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Customer Name <span style={{ color: '#d13438' }}>*</span></label>
            <input type="text" value={formData.customerName}
              onChange={e => set('customerName', e.target.value)}
              placeholder="e.g. Acme Corp" style={INPUT} />
          </div>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>POC / Opportunity Name</label>
            <input type="text" value={formData.pocName}
              onChange={e => set('pocName', e.target.value)}
              placeholder="Optional" style={INPUT} />
          </div>
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
              disabled={!formData.hpenBusinessUnit}
              style={INPUT}>
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
                <button key={p} onClick={() => set('csePriority', p)}
                  style={{ ...TOGGLE_BTN(formData.csePriority === p), borderColor: p === 'High' ? '#d13438' : p === 'Medium' ? '#ca5010' : '#107c10', background: formData.csePriority === p ? (p === 'High' ? '#d13438' : p === 'Medium' ? '#ca5010' : '#107c10') : '#fff', color: formData.csePriority === p ? '#fff' : '#323130' }}>
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

      {/* Section 2 — SSE Selection */}
      <div style={SECTION}>
        <div style={SECTION_TITLE}>SSE Selection</div>
        <PeoplePickerField
          label="Requested SSE"
          required
          value={formData.requestedSse}
          onChange={v => set('requestedSse', v)}
          searchUsers={handleSearchUsers}
        />
      </div>

      {/* Section 3 — Solutions Focus */}
      <div style={SECTION}>
        <div style={SECTION_TITLE}>Solutions Focus</div>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 20px' }}>
          {solutions.map(sol => {
            const checked = focusCodes.includes(sol.code);
            return (
              <label key={sol.code} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 13, cursor: 'pointer' }}>
                <input type="checkbox" checked={checked}
                  onChange={e => {
                    const next = e.target.checked
                      ? [...focusCodes, sol.code]
                      : focusCodes.filter(c => c !== sol.code);
                    set('solutionsFocus', next.join(','));
                  }}
                  style={{ accentColor: HPE_NAVY }} />
                {sol.name}
              </label>
            );
          })}
        </div>
      </div>

      {/* Section 4 — Support Details */}
      <div style={SECTION}>
        <div style={SECTION_TITLE}>Support Details</div>

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Support Type <span style={{ color: '#d13438' }}>*</span></label>
          <div style={{ display: 'flex', gap: 6 }}>
            {(['Remote', 'On-Site', 'Both'] as const).map(t => (
              <button key={t} onClick={() => set('supportType', t)} style={TOGGLE_BTN(formData.supportType === t)}>
                {t}
              </button>
            ))}
          </div>
        </div>

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Description of Need & Skill Set Required <span style={{ color: '#d13438' }}>*</span></label>
          <textarea rows={4} value={formData.cseDescription}
            onChange={e => set('cseDescription', e.target.value)}
            placeholder="Describe what SSE help is needed and the required skill set..."
            style={{ ...INPUT, resize: 'vertical' }} />
        </div>

        {(formData.supportType === 'Remote' || formData.supportType === 'Both') && (
          <ScheduleBlock
            label="Remote Schedule"
            dateOnly
            tbd={formData.remoteTbd}
            start={formData.remoteStart}
            end={formData.remoteEnd}
            duration={formData.remoteDuration}
            onTbd={v => set('remoteTbd', v)}
            onStart={v => set('remoteStart', v)}
            onEnd={v => set('remoteEnd', v)}
            onDuration={v => set('remoteDuration', v)}
          />
        )}

        {(formData.supportType === 'On-Site' || formData.supportType === 'Both') && (
          <ScheduleBlock
            label="On-Site Schedule"
            tbd={formData.onsiteTbd}
            start={formData.onsiteStart}
            end={formData.onsiteEnd}
            duration={formData.onsiteDuration}
            destination={formData.onsiteDestination}
            onTbd={v => set('onsiteTbd', v)}
            onStart={v => set('onsiteStart', v)}
            onEnd={v => set('onsiteEnd', v)}
            onDuration={v => set('onsiteDuration', v)}
            onDestination={v => set('onsiteDestination', v)}
          />
        )}
      </div>

      {/* Submit */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 4 }}>
        <button
          disabled={submitting}
          onClick={handleSubmit}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 700, background: submitting ? '#ccc' : HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: submitting ? 'not-allowed' : 'pointer' }}>
          {submitting ? 'Sending…' : '✉ Send SSE Request'}
        </button>
        <span style={{ fontSize: 12, color: '#605e5c' }}>
          Sends email to SEM, SED &amp; SSE
          {formData.supportType && !formData.remoteTbd && formData.remoteStart ? ' + calendar invite attached' :
           formData.supportType && !formData.onsiteTbd && formData.onsiteStart ? ' + calendar invite attached' : ''}
        </span>
      </div>

    </div>
  );
};
