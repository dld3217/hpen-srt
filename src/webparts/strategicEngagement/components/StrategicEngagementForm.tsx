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
  ENGAGEMENT_PURPOSES, IEnvironmentRow, OUR_VENDORS, GREENFIELD_VENDOR, COMPETITOR_VENDORS,
  isCompetitor, autoDisposition, DISPOSITION_STYLE, environmentHasDisplacement,
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
  landscape: IEnvironmentRow[];
  notes: string;
  supportType: 'Remote' | 'On-Site' | 'Both' | '';
  datesTbd: boolean;
  startDate: string;
  endDate: string;
  duration: string;
  location: string;
  csePriority: string;
  opportunityAmount: number;
}

const EMPTY_FORM: IStrategicFormData = {
  customerName: '', hpenBusinessUnit: '', buRegion: '', requestedSse: '',
  engagementPurpose: '', landscape: [], notes: '',
  supportType: '', datesTbd: false, startDate: '', endDate: '', duration: '', location: '',
  csePriority: '', opportunityAmount: 0,
};

const VERSION = '1.0.1';

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
const CELL_INPUT: React.CSSProperties = { ...INPUT, padding: '5px 6px', fontSize: 12 };
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

// ── SolutionLandscape — what we position + what they run + disposition ─────────
const SolutionLandscape: React.FC<{
  rows: IEnvironmentRow[];
  solutions: ISolutionDef[];
  onChange: (rows: IEnvironmentRow[]) => void;
}> = ({ rows, solutions, onChange }) => {
  const cats = Array.from(new Set(solutions.map(s => s.category)));
  const update = (i: number, patch: Partial<IEnvironmentRow>): void =>
    onChange(rows.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  const setSolution = (i: number, code: string): void => {
    const sol = solutions.filter(s => s.code === code)[0];
    update(i, { solutionCode: code, solution: sol ? sol.name : '' });
  };
  const setVendor = (i: number, vendor: string): void =>
    update(i, { vendor, disposition: autoDisposition(vendor) });
  const remove = (i: number): void => onChange(rows.filter((_, idx) => idx !== i));
  const add = (): void => onChange([...rows, { solutionCode: '', solution: '', vendor: '', product: '', version: '', disposition: '' }]);

  const GRID = '1.4fr 1.1fr 1fr 0.7fr 132px 28px';

  return (
    <div>
      {rows.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: GRID, gap: 6, alignItems: 'center', marginBottom: 4, fontSize: 10, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
          <div>Solution (we position)</div><div>Current Vendor</div><div>Product / Model</div><div>Version</div><div>Disposition</div><div />
        </div>
      )}
      {rows.map((row, i) => {
        const comp = isCompetitor(row.vendor);
        const badge = DISPOSITION_STYLE[row.disposition];
        return (
          <div key={i} style={{ display: 'grid', gridTemplateColumns: GRID, gap: 6, alignItems: 'center', marginBottom: 6 }}>
            <select value={row.solutionCode} onChange={e => setSolution(i, e.target.value)} style={CELL_INPUT}>
              <option value="">— Solution —</option>
              {cats.map(cat => (
                <optgroup key={cat} label={cat}>
                  {solutions.filter(s => s.category === cat).map(s => <option key={s.code} value={s.code}>{s.name}</option>)}
                </optgroup>
              ))}
            </select>
            <select value={row.vendor} onChange={e => setVendor(i, e.target.value)} style={CELL_INPUT}>
              <option value="">— Vendor —</option>
              <optgroup label="HPE (ours)">{OUR_VENDORS.map(v => <option key={v} value={v}>{v}</option>)}</optgroup>
              <optgroup label="Greenfield"><option value={GREENFIELD_VENDOR}>{GREENFIELD_VENDOR}</option></optgroup>
              <optgroup label="Competitors">{COMPETITOR_VENDORS.map(v => <option key={v} value={v}>{v}</option>)}</optgroup>
            </select>
            <input type="text" value={row.product} onChange={e => update(i, { product: e.target.value })} placeholder="e.g. Catalyst 9300" style={CELL_INPUT} />
            <input type="text" value={row.version} onChange={e => update(i, { version: e.target.value })} placeholder="17.9" style={CELL_INPUT} />
            {/* Disposition */}
            {comp ? (
              <div style={{ display: 'flex', gap: 3 }}>
                {(['Integrate', 'Displace'] as const).map(d => {
                  const active = row.disposition === d;
                  const c = DISPOSITION_STYLE[d];
                  return (
                    <button key={d} type="button" onClick={() => update(i, { disposition: d })}
                      style={{ flex: 1, fontSize: 10, fontWeight: 700, padding: '3px 2px', borderRadius: 4, cursor: 'pointer',
                        border: `1px solid ${active ? c.color : '#ccc'}`, background: active ? c.bg : '#fff', color: active ? c.color : '#888' }}>
                      {d === 'Integrate' ? '🤝 Integ.' : '🎯 Displ.'}
                    </button>
                  );
                })}
              </div>
            ) : (
              <span style={{ fontSize: 10, fontWeight: 700, textAlign: 'center', padding: '2px 4px', borderRadius: 8, whiteSpace: 'nowrap',
                background: badge ? badge.bg : '#f0f0f0', color: badge ? badge.color : '#a19f9d' }}>
                {badge ? badge.label : '—'}
              </span>
            )}
            <button type="button" onClick={() => remove(i)} title="Remove row"
              style={{ background: 'transparent', border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', color: '#a4262c', fontSize: 12, lineHeight: 1, padding: '4px 6px' }}>✕</button>
          </div>
        );
      })}
      <button type="button" onClick={add}
        style={{ marginTop: 4, padding: '6px 14px', fontSize: 12, fontWeight: 600, background: '#fff', color: HPE_NAVY, border: `1px dashed ${HPE_NAVY}`, borderRadius: 4, cursor: 'pointer' }}>
        + Add {rows.length === 0 ? 'solution' : 'another'}
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
    const compNoDisp = formData.landscape.some(r => isCompetitor(r.vendor) && !r.disposition);
    if (compNoDisp)                           errs.push('For each competitor row, pick Integrate or Displace.');
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
      const rows       = formData.landscape.filter(r => r.solutionCode || r.vendor || r.product || r.version);
      const focusCodes = Array.from(new Set(rows.map(r => r.solutionCode).filter(Boolean)));
      const schedStatus = formData.datesTbd ? 'TBD' as const : 'Dates Proposed' as const;

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
        solutionsFocus: focusCodes.join(','),
        supportType: formData.supportType,
        remoteTbd: formData.datesTbd, remoteStart: formData.startDate, remoteEnd: formData.endDate, remoteDuration: formData.duration,
        onsiteTbd: formData.datesTbd, onsiteStart: '', onsiteEnd: '', onsiteDuration: '', onsiteDestination: formData.location,
        sePrimary: `${userDisplayName} / ${userEmail}`,
        semPrimary: semEmail, sedEmail,
        buRegion: formData.buRegion, hpenBusinessUnit: formData.hpenBusinessUnit,
        customerName: formData.customerName, pocName: '',
        opportunityAmount: formData.opportunityAmount,
        custTemp: 'Normal', signedOffBy: '', signOffDate: '',
        opportunity: '', notes: formData.notes, specialtyType: '',
        engagementType: 'Strategic Engagement',
        engagementPurpose: formData.engagementPurpose,
        currentEnvironment: JSON.stringify(rows),
        hasDisplacement: environmentHasDisplacement(rows),
        engagementOutcome: 'Advisory Only',
      });
      setSubmitted(true);
    } catch (err) {
      setSubmitError(`Submit failed: ${(err as Error)?.message || String(err)}`);
    } finally { setSubmitting(false); }
  };

  const buKeys     = Object.keys(buRegions).sort();
  const regionKeys = formData.hpenBusinessUnit ? Object.keys((buRegions[formData.hpenBusinessUnit] as IBUConfig)?.regions || {}).sort() : [];

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
    <div style={{ maxWidth: 820, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
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

      {/* Solution Landscape */}
      <div style={SECTION}>
        <SectionHeader title="Solution Landscape" hint="What we'd position, what the customer runs today, and whether we're expanding, integrating, or displacing. Pick 'None / Greenfield' for net-new." />
        <SolutionLandscape rows={formData.landscape} solutions={solutions} onChange={rows => set('landscape', rows)} />
      </div>

      {/* Description & Desired Outcome (moved up) */}
      <div style={SECTION}>
        <SectionHeader title="Description & Desired Outcome" />
        <textarea rows={4} value={formData.notes} onChange={e => set('notes', e.target.value)}
          placeholder="What should the SSE prepare for? Audience (exec / technical), the customer's situation, the outcome you're driving toward…"
          style={{ ...INPUT, resize: 'vertical' }} />
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
        <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, cursor: 'pointer', color: '#323130', marginBottom: 10 }}>
          <input type="checkbox" checked={formData.datesTbd} onChange={e => set('datesTbd', e.target.checked)} style={{ accentColor: HPE_NAVY }} />
          Dates TBD / Flexible — SSE proposes times back
        </label>
        {!formData.datesTbd && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
            <div style={FIELD_ROW}>
              <label style={LABEL_STYLE}>Requested Start</label>
              <input type="date" value={formData.startDate ? formData.startDate.substring(0, 10) : ''}
                onChange={e => set('startDate', e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
            </div>
            <div style={FIELD_ROW}>
              <label style={LABEL_STYLE}>Requested End</label>
              <input type="date" value={formData.endDate ? formData.endDate.substring(0, 10) : ''}
                onChange={e => set('endDate', e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
            </div>
          </div>
        )}
        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Expected Duration / Effort</label>
          <input type="text" value={formData.duration} onChange={e => set('duration', e.target.value)} placeholder="e.g. 90-min briefing, half day on-site" style={INPUT} />
        </div>
        {(formData.supportType === 'On-Site' || formData.supportType === 'Both') && (
          <div style={FIELD_ROW}>
            <label style={LABEL_STYLE}>Location</label>
            <input type="text" value={formData.location} onChange={e => set('location', e.target.value)} placeholder="e.g. Denver, CO" style={INPUT} />
          </div>
        )}
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
