import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { ConfigService, BURegionMap, IBUConfig, IRegionConfig } from '../../../services/ConfigService';
import { CseRequestService } from '../../../services/CseRequestService';
import { ISolutionDef, SOLUTIONS } from '../../../models/ISolution';
import {
  ENGAGEMENT_PURPOSES, IEnvironmentRow, OUR_VENDORS,
  COMPETITOR_VENDORS, isDisplacement, environmentHasDisplacement,
} from '../../../models/StrategicEngagement';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface IStrategicEngagementFormProps {
  sp: SPFI;
  context: WebPartContext;
}

interface IStrategicFormData {
  customerName: string;
  hpenBusinessUnit: string;
  buRegion: string;
  requestedSse: string;
  engagementPurpose: string;
  environment: IEnvironmentRow[];
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
  notes: string;
  csePriority: string;
  opportunityAmount: number;
}

const EMPTY_FORM: IStrategicFormData = {
  customerName: '', hpenBusinessUnit: '', buRegion: '', requestedSse: '',
  engagementPurpose: '', environment: [], solutionsFocus: '', supportType: '',
  remoteTbd: false, remoteStart: '', remoteEnd: '', remoteDuration: '',
  onsiteTbd: false, onsiteStart: '', onsiteEnd: '', onsiteDuration: '', onsiteDestination: '',
  notes: '', csePriority: '', opportunityAmount: 0,
};

const VERSION = '1.0.0';

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

// ── Styles ──────────────────────────────────────────────────────────────────
const SECTION: React.CSSProperties = { background: '#fff', border: '1px solid #e0e0e0', borderRadius: 6, padding: '14px 18px', marginBottom: 12 };
const FIELD_ROW: React.CSSProperties = { marginBottom: 10 };
const LABEL_STYLE: React.CSSProperties = { display: 'block', fontSize: 12, fontWeight: 600, color: '#323130', marginBottom: 4 };
const INPUT: React.CSSProperties = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '6px 8px', border: '1px solid #ccc', borderRadius: 3, fontFamily: 'inherit' };
const TOGGLE_BTN = (active: boolean): React.CSSProperties => ({
  padding: '5px 14px', fontSize: 12, fontWeight: 600, borderRadius: 4,
  border: `1px solid ${active ? HPE_NAVY : '#ccc'}`, background: active ? HPE_NAVY : '#fff',
  color: active ? '#fff' : '#605e5c', cursor: 'pointer',
});

const SectionHeader: React.FC<{ title: string; hint?: string }> = ({ title, hint }) => (
  <div style={{ marginBottom: 12, paddingBottom: 6, borderBottom: `2px solid ${HPE_GREEN}` }}>
    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
      <div style={{ width: 3, height: 16, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
      <span style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px' }}>{title}</span>
    </div>
    {hint && <div style={{ fontSize: 11, color: '#888', marginTop: 3, marginLeft: 9 }}>{hint}</div>}
  </div>
);

// ── PeoplePickerField ─────────────────────────────────────────────────────────
interface IGraphUser { displayName: string; mail: string; }

const PeoplePickerField: React.FC<{
  label: string; value: string; onChange: (val: string) => void;
  required?: boolean; searchUsers?: (query: string) => Promise<IGraphUser[]>;
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

  const startEdit = (): void => { setLocalValue(slashIdx !== -1 ? value.substring(0, slashIdx) : value); setResults([]); setEditing(true); };
  const selectPerson = (person: IGraphUser): void => { onChange(`${person.displayName} / ${person.mail}`); setEditing(false); setResults([]); setLocalValue(''); };
  const handleLocalChange = (q: string): void => {
    setLocalValue(q);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (q.length < 3 || !searchUsers) { setResults([]); return; }
    debounceRef.current = setTimeout(() => {
      setSearching(true);
      searchUsers(q).then(r => setResults(r)).catch(() => setResults([])).finally(() => setSearching(false));
    }, 300);
  };
  const handleBlur = (e: React.FocusEvent<HTMLDivElement>): void => {
    if (containerRef.current?.contains(e.relatedTarget as Node)) return;
    if (!editing) return;
    onChange(localValue.trim()); setEditing(false); setResults([]); setLocalValue('');
  };

  return (
    <div ref={containerRef} onBlur={handleBlur} style={FIELD_ROW}>
      <label style={LABEL_STYLE}>{label}{required && <span style={{ color: '#d13438' }}> *</span>}</label>
      {showDisplay ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', background: '#f3f2f1', borderRadius: 3, border: '1px solid #ccc' }}>
          <a href={`mailto:${emailPart}`} style={{ fontWeight: 600, fontSize: 13, color: '#0078d4', textDecoration: 'none' }}>{namePart}</a>
          <span style={{ fontSize: 11, color: '#605e5c' }}>{emailPart}</span>
          <span style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
            <button type="button" onClick={startEdit} style={{ fontSize: 11, color: '#0078d4', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✎ Change</button>
            <button type="button" onClick={() => onChange('')} style={{ fontSize: 11, color: '#a4262c', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✕</button>
          </span>
        </div>
      ) : (
        <div style={{ position: 'relative' }}>
          <input autoFocus={editing} value={localValue} onChange={e => handleLocalChange(e.target.value)}
            onFocus={() => { if (!editing) startEdit(); }} placeholder={`Search for ${label}...`} style={INPUT} />
          {searching && <span style={{ position: 'absolute', right: 8, top: 8, fontSize: 11, color: '#888' }}>Searching…</span>}
          {results.length > 0 && (
            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #ccc', borderRadius: 3, zIndex: 100, boxShadow: '0 4px 12px rgba(0,0,0,0.12)' }}>
              {results.map(r => (
                <div key={r.mail} tabIndex={0} onMouseDown={() => selectPerson(r)} onKeyDown={e => e.key === 'Enter' && selectPerson(r)}
                  style={{ padding: '7px 12px', cursor: 'pointer', fontSize: 13, borderBottom: '1px solid #f0f0f0' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f3f2f1')} onMouseLeave={e => (e.currentTarget.style.background = '#fff')}>
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

// ── EnvironmentRows — current environment / competitive landscape ──────────────
const EnvironmentRows: React.FC<{
  rows: IEnvironmentRow[];
  areas: string[];
  onChange: (rows: IEnvironmentRow[]) => void;
}> = ({ rows, areas, onChange }) => {
  const update = (i: number, patch: Partial<IEnvironmentRow>): void => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, ...patch } : r));
    onChange(next);
  };
  const remove = (i: number): void => onChange(rows.filter((_, idx) => idx !== i));
  const add = (): void => onChange([...rows, { area: '', vendor: '', product: '', version: '' }]);

  return (
    <div>
      {rows.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1.2fr 1.2fr 0.8fr auto auto', gap: 6, alignItems: 'center', marginBottom: 4, fontSize: 10, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
          <div>Solution Area</div><div>Vendor</div><div>Product / Model</div><div>Version</div><div /><div />
        </div>
      )}
      {rows.map((row, i) => {
        const displacement = isDisplacement(row.vendor);
        return (
          <div key={i} style={{ display: 'grid', gridTemplateColumns: '1.2fr 1.2fr 1.2fr 0.8fr auto auto', gap: 6, alignItems: 'center', marginBottom: 6 }}>
            <select value={row.area} onChange={e => update(i, { area: e.target.value })} style={{ ...INPUT, padding: '5px 6px' }}>
              <option value="">— Area —</option>
              {areas.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
            <select value={row.vendor} onChange={e => update(i, { vendor: e.target.value })} style={{ ...INPUT, padding: '5px 6px' }}>
              <option value="">— Vendor —</option>
              <optgroup label="HPE (ours)">
                {OUR_VENDORS.map(v => <option key={v} value={v}>{v}</option>)}
              </optgroup>
              <optgroup label="Competitors">
                {COMPETITOR_VENDORS.map(v => <option key={v} value={v}>{v}</option>)}
              </optgroup>
            </select>
            <input type="text" value={row.product} onChange={e => update(i, { product: e.target.value })} placeholder="e.g. Catalyst 9300" style={{ ...INPUT, padding: '5px 6px' }} />
            <input type="text" value={row.version} onChange={e => update(i, { version: e.target.value })} placeholder="e.g. 17.9" style={{ ...INPUT, padding: '5px 6px' }} />
            <span style={{ fontSize: 10, fontWeight: 700, whiteSpace: 'nowrap', padding: '2px 6px', borderRadius: 8,
              background: !row.vendor ? '#f0f0f0' : displacement ? '#fde7e9' : '#e8faf3',
              color: !row.vendor ? '#a19f9d' : displacement ? '#a4262c' : '#107c10' }}>
              {!row.vendor ? '—' : displacement ? '🎯 Displace' : '✅ Expand'}
            </span>
            <button type="button" onClick={() => remove(i)} title="Remove row"
              style={{ background: 'transparent', border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#a4262c', fontSize: 13, lineHeight: 1, padding: '4px 8px' }}>✕</button>
          </div>
        );
      })}
      <button type="button" onClick={add}
        style={{ marginTop: 4, padding: '6px 14px', fontSize: 12, fontWeight: 600, background: '#fff', color: HPE_NAVY, border: `1px dashed ${HPE_NAVY}`, borderRadius: 4, cursor: 'pointer' }}>
        + Add {rows.length === 0 ? 'current solution / competitor' : 'another'}
      </button>
    </div>
  );
};

// ── Main Component ────────────────────────────────────────────────────────────
export const StrategicEngagementForm: React.FC<IStrategicEngagementFormProps> = ({ sp, context }) => {
  const [loading, setLoading]         = useState(true);
  const [buRegions, setBuRegions]     = useState<BURegionMap>({});
  const [solutions, setSolutions]     = useState<ISolutionDef[]>(SOLUTIONS);
  const [formData, setFormData]       = useState<IStrategicFormData>(EMPTY_FORM);
  const [submitting, setSubmitting]   = useState(false);
  const [submitted, setSubmitted]     = useState(false);
  const [submitError, setSubmitError] = useState('');
  const [validationErrors, setValidationErrors] = useState<string[]>([]);

  const userEmail       = context.pageContext.user.email;
  const userDisplayName = context.pageContext.user.displayName || userEmail;

  useEffect(() => {
    const configSvc = new ConfigService(sp);
    Promise.all([configSvc.getBURegions(), configSvc.getSolutions()])
      .then(([bur, sols]) => {
        setBuRegions(bur); setSolutions(sols);
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

  const set = (field: keyof IStrategicFormData, value: IStrategicFormData[keyof IStrategicFormData]): void =>
    setFormData(prev => ({ ...prev, [field]: value }));

  const handleSearchUsers = async (query: string): Promise<IGraphUser[]> => {
    try {
      const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const response: any = await client.api('/users')
        .header('ConsistencyLevel', 'eventual')
        .search(`"displayName:${query.replace(/"/g, '')}"`)
        .select('displayName,mail').top(10).get();
      interface IGraphResult { displayName?: string; mail?: string; }
      return ((response.value || []) as IGraphResult[])
        .filter(u => !!u.mail)
        .map(u => ({ displayName: (u.displayName || '') as string, mail: (u.mail || '') as string }));
    } catch { return []; }
  };

  const validate = (): string[] => {
    const errs: string[] = [];
    if (!formData.customerName.trim())        errs.push('Customer Name is required.');
    if (!formData.hpenBusinessUnit)           errs.push('Business Unit is required.');
    if (!formData.buRegion)                   errs.push('Region is required.');
    if (!formData.engagementPurpose)          errs.push('Engagement Purpose is required.');
    if (!formData.requestedSse.includes('/')) errs.push('Requested SSE is required — search and select a person.');
    if (!formData.notes.trim())               errs.push('Description is required.');
    return errs;
  };

  const handleSubmit = async (): Promise<void> => {
    const errs = validate();
    if (errs.length > 0) { setValidationErrors(errs); return; }
    setValidationErrors([]); setSubmitting(true); setSubmitError('');
    try {
      const buConfig   = buRegions[formData.hpenBusinessUnit] as IBUConfig | undefined;
      const regionCfg: IRegionConfig = buConfig?.regions[formData.buRegion] || {};
      const semEmail   = regionCfg.semEmail || '';
      const sedEmail   = buConfig?.sedEmail || '';
      const cleanRows  = formData.environment.filter(r => r.area || r.vendor || r.product || r.version);
      const schedStatus = (formData.remoteTbd && formData.onsiteTbd) ? 'TBD' as const : 'Dates Proposed' as const;

      const svc = new CseRequestService(sp);
      await svc.create({
        title: `${formData.customerName} — Strategic Engagement`,
        source: 'SE Landing Page',
        linkedPocId: 0,
        requestStatus: 'Pending',
        scheduleStatus: schedStatus,
        requestedCse: formData.requestedSse,
        sseManagerEmail: '',
        cseDescription: formData.notes,
        csePriority: formData.csePriority || 'Medium',
        csePriorityReason: '',
        solutionsFocus: formData.solutionsFocus,
        supportType: formData.supportType,
        remoteTbd: formData.remoteTbd, remoteStart: formData.remoteStart, remoteEnd: formData.remoteEnd, remoteDuration: formData.remoteDuration,
        onsiteTbd: formData.onsiteTbd, onsiteStart: formData.onsiteStart, onsiteEnd: formData.onsiteEnd, onsiteDuration: formData.onsiteDuration, onsiteDestination: formData.onsiteDestination,
        sePrimary: `${userDisplayName} / ${userEmail}`,
        semPrimary: semEmail, sedEmail,
        buRegion: formData.buRegion, hpenBusinessUnit: formData.hpenBusinessUnit,
        customerName: formData.customerName, pocName: '',
        opportunityAmount: formData.opportunityAmount,
        custTemp: 'Normal', signedOffBy: '', signOffDate: '',
        opportunity: '', notes: formData.notes, specialtyType: '',
        // Strategic Engagement fields
        engagementType: 'Strategic Engagement',
        engagementPurpose: formData.engagementPurpose,
        currentEnvironment: JSON.stringify(cleanRows),
        hasDisplacement: environmentHasDisplacement(cleanRows),
        engagementOutcome: 'Advisory Only',
      });
      setSubmitted(true);
    } catch (err) {
      setSubmitError(`Submit failed: ${(err as Error)?.message || String(err)}`);
    } finally { setSubmitting(false); }
  };

  const buKeys     = Object.keys(buRegions).sort();
  const regionKeys = formData.hpenBusinessUnit ? Object.keys((buRegions[formData.hpenBusinessUnit] as IBUConfig)?.regions || {}).sort() : [];
  const focusCodes = formData.solutionsFocus ? formData.solutionsFocus.split(',').map(c => c.trim()).filter(Boolean) : [];
  const areas: string[] = Array.from(new Set(solutions.map(s => s.category)));

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading…" style={{ marginTop: 32 }} />;

  if (submitted) {
    return (
      <div style={{ maxWidth: 680, margin: '40px auto', textAlign: 'center', padding: 32 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>🎯</div>
        <div style={{ fontSize: 22, fontWeight: 700, color: HPE_NAVY, marginBottom: 8 }}>Strategic Engagement Requested</div>
        <div style={{ fontSize: 14, color: '#605e5c', marginBottom: 24 }}>
          Your request for <strong>{formData.customerName}</strong> has been submitted and will route to the SSE for scheduling.
        </div>
        <button onClick={() => { setFormData(EMPTY_FORM); setSubmitted(false); }}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 600, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
          Submit Another
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 760, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '8px 16px', borderRadius: '6px 6px 0 0', marginBottom: 12, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 22, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, lineHeight: 1.2 }}>Strategic Engagement Request</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — SSE time outside an active POC</div>
          </div>
        </div>
        <div style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>
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

      {/* Request Info */}
      <div style={SECTION}>
        <SectionHeader title="Request Info" />
        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Customer Name <span style={{ color: '#d13438' }}>*</span></label>
          <input type="text" value={formData.customerName} onChange={e => set('customerName', e.target.value)} placeholder="e.g. Acme Corp" style={INPUT} />
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
              }} style={INPUT}>
              <option value="">— Select BU —</option>
              {buKeys.map(bu => <option key={bu} value={bu}>{bu}</option>)}
            </select>
          </div>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Region <span style={{ color: '#d13438' }}>*</span></label>
            <select value={formData.buRegion}
              onChange={e => { set('buRegion', e.target.value); if (e.target.value) localStorage.setItem('srt_region', e.target.value); else localStorage.removeItem('srt_region'); }}
              disabled={!formData.hpenBusinessUnit} style={INPUT}>
              <option value="">— Select Region —</option>
              {regionKeys.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>
        <PeoplePickerField label="Requested SSE" required value={formData.requestedSse} onChange={v => set('requestedSse', v)} searchUsers={handleSearchUsers} />
      </div>

      {/* Engagement Purpose */}
      <div style={SECTION}>
        <SectionHeader title="Engagement Purpose" hint="What kind of strategic engagement is this?" />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
          {ENGAGEMENT_PURPOSES.map(p => (
            <button key={p} type="button" onClick={() => set('engagementPurpose', p)} style={TOGGLE_BTN(formData.engagementPurpose === p)}>{p}</button>
          ))}
        </div>
      </div>

      {/* Current Environment & Competitive Landscape */}
      <div style={SECTION}>
        <SectionHeader title="Current Environment & Competitive Landscape" hint="What does the customer run today? Flag competitors we could displace — the more detail, the sharper the SSE's prep." />
        <EnvironmentRows rows={formData.environment} areas={areas} onChange={rows => set('environment', rows)} />
      </div>

      {/* Solutions Focus */}
      <div style={SECTION}>
        <SectionHeader title="Solutions Focus" hint="What we'd position in the conversation" />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 20px' }}>
          {solutions.map(sol => {
            const checked = focusCodes.includes(sol.code);
            return (
              <label key={sol.code} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 13, cursor: 'pointer' }}>
                <input type="checkbox" checked={checked}
                  onChange={e => { const next = e.target.checked ? [...focusCodes, sol.code] : focusCodes.filter(c => c !== sol.code); set('solutionsFocus', next.join(',')); }}
                  style={{ accentColor: HPE_GREEN }} />
                {sol.name}
              </label>
            );
          })}
        </div>
      </div>

      {/* Time Coordination */}
      <div style={SECTION}>
        <SectionHeader title="Time Coordination" />
        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Format</label>
          <div style={{ display: 'flex', gap: 6 }}>
            {(['Remote', 'On-Site', 'Both'] as const).map(t => (
              <button key={t} type="button" onClick={() => set('supportType', t)} style={TOGGLE_BTN(formData.supportType === t)}>{t}</button>
            ))}
          </div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Requested Start</label>
            <input type="date" value={formData.remoteStart ? formData.remoteStart.substring(0, 10) : ''}
              onChange={e => set('remoteStart', e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
          </div>
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Requested End</label>
            <input type="date" value={formData.remoteEnd ? formData.remoteEnd.substring(0, 10) : ''}
              onChange={e => set('remoteEnd', e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
          </div>
        </div>
        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Expected Duration / Effort</label>
          <input type="text" value={formData.remoteDuration} onChange={e => set('remoteDuration', e.target.value)} placeholder="e.g. 90-min briefing, half day on-site" style={INPUT} />
        </div>
        {(formData.supportType === 'On-Site' || formData.supportType === 'Both') && (
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Location</label>
            <input type="text" value={formData.onsiteDestination} onChange={e => set('onsiteDestination', e.target.value)} placeholder="e.g. Denver, CO" style={INPUT} />
          </div>
        )}
      </div>

      {/* Description */}
      <div style={SECTION}>
        <SectionHeader title="Description & Desired Outcome" />
        <textarea rows={4} value={formData.notes} onChange={e => set('notes', e.target.value)}
          placeholder="What should the SSE prepare for? Audience (exec / technical), the customer's situation, the outcome you're driving toward…"
          style={{ ...INPUT, resize: 'vertical' }} />
      </div>

      {/* Submit */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 4, marginBottom: 32 }}>
        <button disabled={submitting} onClick={handleSubmit}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 700, background: submitting ? '#ccc' : HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: submitting ? 'not-allowed' : 'pointer' }}>
          {submitting ? 'Sending…' : '🎯 Request Strategic Engagement'}
        </button>
        <span style={{ fontSize: 12, color: '#605e5c' }}>Routes to the SSE for scheduling — lands in the SRT tracker</span>
      </div>
    </div>
  );
};
