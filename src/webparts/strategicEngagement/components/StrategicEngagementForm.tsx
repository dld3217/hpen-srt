import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { APP_VERSION } from '../../../appVersion';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { ConfigService, BURegionMap, IBUConfig, IRegionConfig } from '../../../services/ConfigService';
import { CseRequestService } from '../../../services/CseRequestService';
import { ISseCommitment } from '../../../models/ICseRequest';
import { ISolutionDef, SOLUTIONS } from '../../../models/ISolution';
import {
  ENGAGEMENT_PURPOSES, DESIRED_OUTCOMES, OUTCOME_OBJECTION, OUTCOME_OTHER, IEnvironmentRow,
  OUR_VENDORS, GREENFIELD_VENDOR, COMPETITOR_VENDORS,
  isCompetitor, autoDisposition, DISPOSITION_STYLE, environmentHasDisplacement,
} from '../../../models/StrategicEngagement';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface IStrategicEngagementFormProps {
  sp: SPFI;
  context: WebPartContext;
  onBack?: () => void;
}

interface IStrategicFormData {
  customerName: string;
  hpenBusinessUnit: string;
  buRegion: string;
  primarySe: string;
  requestedSse: string;
  engagementPurpose: string;
  engagementPurposeOther: string;
  landscape: IEnvironmentRow[];
  notes: string;
  desiredOutcomes: string[];
  objectionText: string;
  outcomeOtherText: string;
  supportType: 'Remote' | 'On-Site' | 'Both' | '';
  remoteTbd: boolean;
  onsiteTbd: boolean;
  remoteStart: string;
  remoteEnd: string;
  remoteDuration: string;
  onsiteStart: string;
  onsiteEnd: string;
  onsiteDuration: string;
  location: string;
  csePriority: string;
  csePriorityReason: string;
  opportunityAmount: number;
}

const EMPTY_FORM: IStrategicFormData = {
  customerName: '', hpenBusinessUnit: '', buRegion: '', primarySe: '', requestedSse: '',
  engagementPurpose: '', engagementPurposeOther: '', landscape: [], notes: '',
  desiredOutcomes: [], objectionText: '', outcomeOtherText: '',
  supportType: '', remoteTbd: true, onsiteTbd: true,
  remoteStart: '', remoteEnd: '', remoteDuration: '1-hour meeting',
  onsiteStart: '', onsiteEnd: '', onsiteDuration: '', location: '',
  csePriority: 'Medium', csePriorityReason: '', opportunityAmount: 0,
};

const VERSION = APP_VERSION;

// Quick-pick duration presets per schedule format (SE can still type a custom value)
const REMOTE_PRESETS = ['30-min call', '1-hour meeting', '2-hour session', 'Half day'];
const ONSITE_PRESETS = ['Half day', 'Full day', '2 days', 'Multi-day'];
const SRT_DASHBOARD_URL = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/SRT-Resource-Dashboard.aspx';

// ── Demo data (admin-only quick-fill for live demos) ──────────────────────────
// Each click of the header pill loads the next scenario. Edit any value below to
// change what loads; BU/Region + solution names resolve from live config at fill time.
interface IDemoRow { code: string; vendor: string; disposition: string; product: string; version: string; detail: string; }
interface IDemoScenario {
  name: string; customerName: string; engagementPurpose: string;
  requestedSse: string; notes: string; desiredOutcomes: string[];
  supportType: 'Remote' | 'On-Site' | 'Both'; datesTbd?: boolean; location?: string;
  remoteDuration?: string; onsiteDuration?: string; csePriority: string; priorityReason?: string; opportunityAmount: number;
  rows: IDemoRow[];
}
const DEMO_SSE = 'Charlie Clemmer / charlie.clemmer@hpe.com';
const DEMO_SCENARIOS: IDemoScenario[] = [
  {
    name: 'Cisco displacement',
    customerName: '[SAMPLE] Acme Corporation',
    engagementPurpose: 'EOL / Migration Planning',
    requestedSse: DEMO_SSE,
    notes: 'Cisco Catalyst estate hits end-of-support next fiscal year. Exec + network-architecture audience wants a migration roadmap to HPE Aruba CX plus a story for their Meraki wireless. Competitive displacement play — Cisco incumbent 8+ years.',
    desiredOutcomes: ['Displace the incumbent', 'Prove the proposed solution'],
    supportType: 'On-Site', location: 'Denver, CO', onsiteDuration: 'Full-day workshop',
    csePriority: 'High', priorityReason: 'Catalyst support expires end of next fiscal year — customer needs the migration roadmap before their budget lock in Q3.', opportunityAmount: 850000,
    rows: [
      { code: 'HCXS', vendor: 'Cisco', disposition: 'Displace', product: 'Catalyst 9300', version: '17.9', detail: 'Catalyst 9300 fabric EOL FY27; migrate access + core to CX 8100/8360 over 3 phases.' },
      { code: 'HWRL', vendor: 'Cisco Meraki', disposition: 'Displace', product: 'MR46', version: '', detail: 'Meraki wireless up for renewal; position AOS10 to consolidate licensing.' },
    ],
  },
  {
    name: 'Mist replaces Meraki (Cisco integrate)',
    customerName: '[SAMPLE] Globex Industries',
    engagementPurpose: 'Architecture / Exec Conversation',
    requestedSse: DEMO_SSE,
    notes: 'Globex keeps its Cisco Catalyst 9300 switching core but wants a new Mist AI WLAN to replace its aging Cisco Meraki wireless. They already run HPE Aruba ClearPass for NAC — the engagement must prove ClearPass and the new Mist WLAN interoperate cleanly with the retained Catalyst switches. Technical + architecture audience; the ask is a credible interoperability + migration story before they commit.',
    desiredOutcomes: ['Prove the proposed solution', 'Displace the incumbent'],
    supportType: 'Remote', remoteDuration: '90-min briefing',
    csePriority: 'Medium', opportunityAmount: 300000,
    rows: [
      { code: '', vendor: 'Cisco', disposition: 'Integrate', product: 'Catalyst 9300', version: '17.9', detail: 'Customer retains the Catalyst 9300 access/core. The new Mist WLAN and existing ClearPass NAC must interoperate across the mixed fabric — 802.1X / RADIUS, dynamic segmentation, and consistent policy enforcement on the Cisco switch ports.' },
      { code: 'MWRL', vendor: 'Cisco Meraki', disposition: 'Displace', product: 'MR44 / MR46', version: '', detail: 'Replace the Meraki wireless estate with a new Mist AI deployment; ClearPass enforcement policy must carry over to the Mist WLAN.' },
      { code: 'NSCP', vendor: 'HPE Aruba', disposition: 'Expansion', product: 'ClearPass Policy Manager', version: '', detail: '' },
    ],
  },
  {
    name: 'Greenfield / net-new',
    customerName: '[SAMPLE] Initech LLC',
    engagementPurpose: 'Pre-Sales Advisory',
    requestedSse: DEMO_SSE,
    notes: 'New logo, net-new build for a manufacturing site. No incumbent networking vendor for the OT / private-wireless layer. Early advisory to shape requirements around Private 5G before the RFP.',
    desiredOutcomes: ["Shape the requirements / get spec'd in", 'Prove the proposed solution'],
    supportType: 'Both', datesTbd: true,
    csePriority: 'Medium', opportunityAmount: 500000,
    rows: [
      { code: 'PW5G', vendor: 'None / Greenfield', disposition: 'New', product: '', version: '', detail: '' },
    ],
  },
];

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

// ── ScheduleBlock — per-block TBD toggle, quick-pick duration, dates, location ──
const ScheduleBlock: React.FC<{
  label: string; start: string; end: string; duration: string; location?: string;
  tbd: boolean; presets: string[];
  onStart: (v: string) => void; onEnd: (v: string) => void; onDuration: (v: string) => void;
  onTbd: (v: boolean) => void; onLocation?: (v: string) => void;
}> = ({ label, start, end, duration, location, tbd, presets, onStart, onEnd, onDuration, onTbd, onLocation }) => (
  <div style={{ background: '#f9f9f9', border: '1px solid #e8e8e8', borderRadius: 4, padding: '10px 14px', marginBottom: 10 }}>
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, marginBottom: 8 }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: '#555', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{label}</div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11, cursor: 'pointer', color: '#605e5c', whiteSpace: 'nowrap' }}>
        <input type="checkbox" checked={tbd} onChange={e => onTbd(e.target.checked)} style={{ accentColor: HPE_NAVY }} />
        Dates TBD / flexible
      </label>
    </div>
    {!tbd && (
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 8 }}>
        <div>
          <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Requested Start</label>
          <input type="date" value={start ? start.substring(0, 10) : ''} onChange={e => onStart(e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
        </div>
        <div>
          <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Requested End</label>
          <input type="date" value={end ? end.substring(0, 10) : ''} onChange={e => onEnd(e.target.value ? new Date(e.target.value).toISOString() : '')} style={INPUT} />
        </div>
      </div>
    )}
    <div style={{ marginBottom: onLocation ? 8 : 0 }}>
      <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Expected Duration / Effort</label>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 6 }}>
        {presets.map(p => {
          const active = duration === p;
          return (
            <button key={p} type="button" onClick={() => onDuration(active ? '' : p)}
              style={{ fontSize: 11, fontWeight: 600, padding: '3px 10px', borderRadius: 12, cursor: 'pointer',
                border: `1px solid ${active ? HPE_NAVY : '#ccc'}`, background: active ? HPE_NAVY : '#fff', color: active ? '#fff' : '#605e5c' }}>
              {p}
            </button>
          );
        })}
      </div>
      <input type="text" value={duration} onChange={e => onDuration(e.target.value)} placeholder="…or type a custom duration" style={INPUT} />
    </div>
    {onLocation && (
      <div>
        <label style={{ ...LABEL_STYLE, fontSize: 11 }}>Location</label>
        <input type="text" value={location || ''} onChange={e => onLocation(e.target.value)} placeholder="e.g. Denver, CO" style={INPUT} />
      </div>
    )}
  </div>
);

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
  const add = (): void => onChange([...rows, { solutionCode: '', solution: '', vendor: '', product: '', version: '', disposition: '', detail: '' }]);

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
        const needsDetail = row.disposition === 'Integrate' && !(row.detail || '').trim();
        return (
          <div key={i} style={{ marginBottom: 8 }}>
            <div style={{ display: 'grid', gridTemplateColumns: GRID, gap: 6, alignItems: 'center' }}>
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
            {comp && !!row.disposition && (
              <input type="text" value={row.detail || ''} onChange={e => update(i, { detail: e.target.value })}
                placeholder={row.disposition === 'Integrate'
                  ? 'Integration detail — how do they coexist? where\'s the boundary? (required)'
                  : 'Displacement detail — migration story / timeline / what\'s driving the switch (optional)'}
                style={{ ...CELL_INPUT, marginTop: 4, borderColor: needsDetail ? '#d13438' : '#ccc',
                  background: needsDetail ? '#fef6f6' : '#fff' }} />
            )}
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
export const StrategicEngagementForm: React.FC<IStrategicEngagementFormProps> = ({ sp, context, onBack }) => {
  const [loading, setLoading]         = useState(true);
  const [buRegions, setBuRegions]     = useState<BURegionMap>({});
  const [solutions, setSolutions]     = useState<ISolutionDef[]>(SOLUTIONS);
  const [formData, setFormData]       = useState<IStrategicFormData>(EMPTY_FORM);
  const [submitting, setSubmitting]   = useState(false);
  const [submitted, setSubmitted]     = useState(false);
  const [submitError, setSubmitError] = useState('');
  const [validationErrors, setValidationErrors] = useState<string[]>([]);
  const [isAdmin, setIsAdmin]         = useState(false);
  const [demoIdx, setDemoIdx]         = useState(0);
  const [commitments, setCommitments] = useState<ISseCommitment[]>([]);
  const [commitLoading, setCommitLoading] = useState(false);

  const userEmail       = context.pageContext.user.email;
  const userDisplayName = context.pageContext.user.displayName || userEmail;

  // Load the selected SSE's confirmed upcoming commitments so the SE can avoid double-booking.
  const sseEmail = formData.requestedSse.includes('/') ? formData.requestedSse.split('/')[1].trim() : '';
  useEffect(() => {
    if (!sseEmail) { setCommitments([]); return; }
    let cancelled = false;
    setCommitLoading(true);
    new CseRequestService(sp).getSseCommitments(sseEmail)
      .then(c => { if (!cancelled) setCommitments(c); })
      .catch(() => { if (!cancelled) setCommitments([]); })
      .finally(() => { if (!cancelled) setCommitLoading(false); });
    return () => { cancelled = true; };
  }, [sseEmail]);

  useEffect(() => {
    const configSvc = new ConfigService(sp);
    Promise.all([configSvc.getBURegions(), configSvc.getSolutions(), configSvc.isSuperUser(userEmail)])
      .then(([bur, sols, admin]) => {
        setBuRegions(bur); setSolutions(sols); setIsAdmin(admin);
        const savedBU = localStorage.getItem('srt_bu') || '';
        const savedRegion = localStorage.getItem('srt_region') || '';
        const buOk = !!(savedBU && bur[savedBU]);
        const validRegion = (buOk && savedRegion && (bur[savedBU] as IBUConfig)?.regions?.[savedRegion]) ? savedRegion : '';
        const savedSseEmail = buOk ? (((bur[savedBU] as IBUConfig)?.regions?.[validRegion]?.sseEmail) || (bur[savedBU] as IBUConfig)?.sseEmail || '').trim() : '';
        const savedSse = savedSseEmail ? `${savedSseEmail.split('@')[0].split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ')} / ${savedSseEmail}` : '';
        setFormData(prev => ({
          ...prev,
          primarySe: prev.primarySe || `${userDisplayName} / ${userEmail}`,
          hpenBusinessUnit: buOk ? savedBU : prev.hpenBusinessUnit,
          buRegion: buOk ? validRegion : prev.buRegion,
          requestedSse: prev.requestedSse || savedSse,
        }));
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  const set = (field: keyof IStrategicFormData, value: IStrategicFormData[keyof IStrategicFormData]): void =>
    setFormData(prev => ({ ...prev, [field]: value }));

  // Covering SSE for a BU/Region as a "Name / email" picker value: region-level wins, else BU default. "" if none.
  const emailToName = (email: string): string =>
    email.split('@')[0].split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
  const resolveSse = (bu: string, region: string): string => {
    const buCfg = buRegions[bu] as IBUConfig | undefined;
    const email = ((region && buCfg?.regions?.[region]?.sseEmail) || buCfg?.sseEmail || '').trim();
    return email ? `${emailToName(email)} / ${email}` : '';
  };

  const loadDemo = (): void => {
    const s = DEMO_SCENARIOS[demoIdx % DEMO_SCENARIOS.length];
    // Resolve a valid BU + Region from live config (keep current selection if still valid)
    const buKey = (formData.hpenBusinessUnit && buRegions[formData.hpenBusinessUnit])
      ? formData.hpenBusinessUnit : (Object.keys(buRegions).sort()[0] || '');
    const regionsForBu = buKey ? Object.keys((buRegions[buKey] as IBUConfig)?.regions || {}).sort() : [];
    const regionKey = (formData.buRegion && regionsForBu.indexOf(formData.buRegion) !== -1)
      ? formData.buRegion : (regionsForBu[0] || '');
    const landscape: IEnvironmentRow[] = s.rows.map(r => {
      const sol = solutions.filter(x => x.code === r.code)[0];
      return { solutionCode: r.code, solution: sol ? sol.name : '', vendor: r.vendor, product: r.product, version: r.version, disposition: r.disposition, detail: r.detail };
    });
    setFormData({
      ...EMPTY_FORM,
      primarySe: `${userDisplayName} / ${userEmail}`,
      hpenBusinessUnit: buKey, buRegion: regionKey,
      customerName: s.customerName,
      engagementPurpose: s.engagementPurpose,
      requestedSse: s.requestedSse,
      notes: s.notes,
      desiredOutcomes: s.desiredOutcomes,
      landscape,
      supportType: s.supportType,
      remoteTbd: !!s.datesTbd,
      onsiteTbd: !!s.datesTbd,
      remoteDuration: s.remoteDuration || '',
      onsiteDuration: s.onsiteDuration || '',
      location: s.location || '',
      csePriority: s.csePriority,
      csePriorityReason: s.priorityReason || '',
      opportunityAmount: s.opportunityAmount,
    });
    setValidationErrors([]); setSubmitError('');
    setDemoIdx(demoIdx + 1);
  };

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
    if (!formData.primarySe.includes('/'))    errs.push('Primary SE is required.');
    if (!formData.engagementPurpose)          errs.push('Engagement Purpose is required.');
    if (formData.engagementPurpose === 'Other' && !formData.engagementPurposeOther.trim())
      errs.push('Describe the engagement (Engagement Purpose = Other).');
    if (!formData.requestedSse.includes('/')) errs.push('Requested SSE is required — search and select a person.');
    if (!formData.notes.trim())               errs.push('Description is required.');
    if (formData.landscape.some(r => isCompetitor(r.vendor) && !r.disposition))
      errs.push('For each competitor row, pick Integrate or Displace.');
    if (formData.landscape.some(r => r.disposition === 'Integrate' && !(r.detail || '').trim()))
      errs.push('For each Integrate row, add the integration detail (how they coexist).');
    if (formData.landscape.some(r => ((r.vendor || '').trim() || (r.product || '').trim() || (r.version || '').trim()) && !r.solutionCode))
      errs.push('Each Solution Landscape row with a vendor or product needs a Solution selected — rows without one are not captured.');
    if (formData.desiredOutcomes.indexOf(OUTCOME_OBJECTION) !== -1 && !formData.objectionText.trim())
      errs.push('List the specific objection(s) to overcome.');
    if ((formData.csePriority === 'High' || formData.csePriority === 'Critical') && !formData.csePriorityReason.trim())
      errs.push(`Explain why this is ${formData.csePriority} priority.`);
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
      const wantRemote = formData.supportType === 'Remote' || formData.supportType === 'Both';
      const wantOnsite = formData.supportType === 'On-Site' || formData.supportType === 'Both';
      // "Dates Proposed" requires an actual start date, not just the TBD toggle being off —
      // otherwise the tracker shows a proposed schedule with no dates and prompts the SSE to confirm nothing.
      const hasFirmDates = (wantRemote && !formData.remoteTbd && !!formData.remoteStart)
                        || (wantOnsite && !formData.onsiteTbd && !!formData.onsiteStart);
      const schedStatus = hasFirmDates ? 'Dates Proposed' as const : 'TBD' as const;

      const svc = new CseRequestService(sp);
      await svc.create({
        title: `${formData.customerName} — Strategic Engagement`,
        source: 'SE Landing Page',
        linkedPocId: 0,
        requestStatus: 'Pending',
        scheduleStatus: schedStatus,
        datesProposedBy: hasFirmDates ? 'SE' : '', // SE proposes at submit; SSE confirms
        requestedCse: formData.requestedSse,
        sseManagerEmail: '',
        cseDescription: formData.notes,
        csePriority: formData.csePriority || 'Medium',
        csePriorityReason: formData.csePriorityReason,
        solutionsFocus: focusCodes.join(','),
        supportType: formData.supportType,
        remoteTbd: formData.remoteTbd, remoteStart: formData.remoteTbd ? '' : formData.remoteStart, remoteEnd: formData.remoteTbd ? '' : formData.remoteEnd, remoteDuration: formData.remoteDuration,
        onsiteTbd: formData.onsiteTbd, onsiteStart: formData.onsiteTbd ? '' : formData.onsiteStart, onsiteEnd: formData.onsiteTbd ? '' : formData.onsiteEnd, onsiteDuration: formData.onsiteDuration, onsiteDestination: formData.location,
        sePrimary: formData.primarySe,
        semPrimary: semEmail, sedEmail,
        buRegion: formData.buRegion, hpenBusinessUnit: formData.hpenBusinessUnit,
        customerName: formData.customerName, pocName: '',
        opportunityAmount: formData.opportunityAmount,
        custTemp: 'Normal', signedOffBy: '', signOffDate: '',
        opportunity: '', notes: formData.notes, specialtyType: '',
        engagementType: 'Strategic Engagement',
        engagementPurpose: formData.engagementPurpose,
        engagementPurposeOther: formData.engagementPurpose === 'Other' ? formData.engagementPurposeOther : '',
        currentEnvironment: JSON.stringify(rows),
        hasDisplacement: environmentHasDisplacement(rows),
        engagementOutcome: 'Advisory Only',
        desiredOutcome: formData.desiredOutcomes,
        desiredOutcomeDetail: [
          formData.objectionText.trim() ? `Objection(s): ${formData.objectionText.trim()}` : '',
          formData.outcomeOtherText.trim() ? `Other: ${formData.outcomeOtherText.trim()}` : '',
        ].filter(Boolean).join(' | '),
      });
      setSubmitted(true);
    } catch (err) {
      setSubmitError(`Submit failed: ${(err as Error)?.message || String(err)}`);
    } finally { setSubmitting(false); }
  };

  const buKeys     = Object.keys(buRegions).sort();
  const regionKeys = formData.hpenBusinessUnit ? Object.keys((buRegions[formData.hpenBusinessUnit] as IBUConfig)?.regions || {}).sort() : [];
  const showRemote = formData.supportType === 'Remote' || formData.supportType === 'Both';
  const showOnsite = formData.supportType === 'On-Site' || formData.supportType === 'Both';

  // ── SSE availability helpers ──
  const dstr = (iso: string): string => (iso || '').substring(0, 10);
  const fmtD = (iso: string): string => {
    const p = dstr(iso); if (!p) return '';
    const [y, m, d] = p.split('-').map(Number);
    return new Date(y, m - 1, d).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  };
  const fmtRange = (a: string, b: string): string => {
    const s = fmtD(a), e = fmtD(b);
    return e && e !== s ? `${s} – ${e}` : s;
  };
  const rangesOverlap = (aS: string, aE: string, bS: string, bE: string): boolean => aS <= bE && bS <= aE;
  // The dates the SE is actively proposing on this form (only firm, non-TBD blocks).
  const proposedBlocks: Array<{ s: string; e: string }> = [];
  if (showRemote && !formData.remoteTbd && formData.remoteStart) proposedBlocks.push({ s: dstr(formData.remoteStart), e: dstr(formData.remoteEnd || formData.remoteStart) });
  if (showOnsite && !formData.onsiteTbd && formData.onsiteStart) proposedBlocks.push({ s: dstr(formData.onsiteStart), e: dstr(formData.onsiteEnd || formData.onsiteStart) });
  const commitConflicts = (c: ISseCommitment): boolean =>
    proposedBlocks.some(p => rangesOverlap(p.s, p.e, dstr(c.start), dstr(c.end)));
  const anyConflict = commitments.some(commitConflicts);
  const sseShortName = formData.requestedSse.includes('/') ? formData.requestedSse.split('/')[0].trim() : '';

  if (loading) return <Spinner size={SpinnerSize.large} label="Loading…" style={{ marginTop: 32 }} />;

  if (submitted) {
    return (
      <div style={{ maxWidth: 680, margin: '40px auto', textAlign: 'center', padding: 32 }}>
        <div style={{ fontSize: 48, marginBottom: 12 }}>🎯</div>
        <div style={{ fontSize: 22, fontWeight: 700, color: HPE_NAVY, marginBottom: 8 }}>Strategic Engagement Requested</div>
        <div style={{ fontSize: 14, color: '#605e5c', marginBottom: 24 }}>
          Your request for <strong>{formData.customerName}</strong> has been submitted and will route to the SSE for scheduling.
        </div>
        <div style={{ display: 'flex', gap: 12, justifyContent: 'center', flexWrap: 'wrap' }}>
          <button onClick={() => { setFormData({ ...EMPTY_FORM, primarySe: `${userDisplayName} / ${userEmail}` }); setSubmitted(false); }}
            style={{ padding: '10px 28px', fontSize: 14, fontWeight: 600, background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}>
            Submit Another
          </button>
          <a href={SRT_DASHBOARD_URL}
            style={{ padding: '10px 28px', fontSize: 14, fontWeight: 600, background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 4, cursor: 'pointer', textDecoration: 'none' }}>
            Return to Dashboard
          </a>
        </div>
      </div>
    );
  }

  return (
    <div style={{ maxWidth: 820, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {onBack && (
        <button type="button" onClick={onBack}
          style={{ background: 'none', border: 'none', color: HPE_NAVY, fontSize: 13, fontWeight: 600, cursor: 'pointer', padding: '0 0 8px', display: 'flex', alignItems: 'center', gap: 4 }}>← Back to options</button>
      )}
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '8px 16px', borderRadius: '6px 6px 0 0', marginBottom: 12, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 22, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, lineHeight: 1.2 }}>Strategic Engagement Request</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — SSE time outside an active POC</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, whiteSpace: 'nowrap' }}>
          {isAdmin && (
            <button type="button" onClick={loadDemo}
              title={`Fill with demo data — next: ${DEMO_SCENARIOS[demoIdx % DEMO_SCENARIOS.length].name} (admins only)`}
              style={{ fontSize: 11, fontWeight: 700, color: '#fff', background: 'rgba(255,255,255,0.14)', border: '1px solid rgba(255,255,255,0.4)', borderRadius: 12, padding: '4px 11px', cursor: 'pointer' }}>
              🧪 Demo {(demoIdx % DEMO_SCENARIOS.length) + 1}/{DEMO_SCENARIOS.length}
            </button>
          )}
          <div style={{ textAlign: 'right' }}>
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
                const sse = resolveSse(newBU, '');
                setFormData(prev => ({ ...prev, hpenBusinessUnit: newBU, buRegion: '', requestedSse: sse || prev.requestedSse }));
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
              onChange={e => {
                const newRegion = e.target.value;
                setFormData(prev => { const sse = resolveSse(prev.hpenBusinessUnit, newRegion); return { ...prev, buRegion: newRegion, requestedSse: sse || prev.requestedSse }; });
                if (newRegion) localStorage.setItem('srt_region', newRegion); else localStorage.removeItem('srt_region');
              }}
              disabled={!formData.hpenBusinessUnit} style={INPUT}>
              <option value="">— Select Region —</option>
              {regionKeys.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>
        <PeoplePickerField label="Primary SE" required value={formData.primarySe} onChange={v => set('primarySe', v)} searchUsers={handleSearchUsers} />
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
        {formData.engagementPurpose === 'Other' && (
          <div style={{ marginTop: 10 }}>
            <label style={LABEL_STYLE}>Describe the engagement <span style={{ color: '#d13438' }}>*</span></label>
            <input type="text" value={formData.engagementPurposeOther} onChange={e => set('engagementPurposeOther', e.target.value)}
              placeholder="What kind of engagement is this?"
              style={{ ...INPUT, borderColor: !formData.engagementPurposeOther.trim() ? '#d13438' : '#ccc', background: !formData.engagementPurposeOther.trim() ? '#fef6f6' : '#fff' }} />
          </div>
        )}
      </div>

      {/* Description */}
      <div style={SECTION}>
        <SectionHeader title="Description" hint="What should the SSE prepare for? Audience, the customer's situation, relevant context." />
        <textarea rows={4} value={formData.notes} onChange={e => set('notes', e.target.value)}
          placeholder="Audience (exec / technical), the customer's situation, and any context that sharpens the SSE's prep…"
          style={{ ...INPUT, resize: 'vertical' }} />
      </div>

      {/* Desired Outcome */}
      <div style={SECTION}>
        <SectionHeader title="Desired Outcome" hint="What are you driving toward? Pick all that apply — a blended audience can have several." />
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
          {DESIRED_OUTCOMES.map(o => {
            const active = formData.desiredOutcomes.indexOf(o) !== -1;
            return (
              <button key={o} type="button"
                onClick={() => setFormData(prev => ({
                  ...prev,
                  desiredOutcomes: active ? prev.desiredOutcomes.filter(x => x !== o) : [...prev.desiredOutcomes, o],
                }))}
                style={TOGGLE_BTN(active)}>{active ? '✓ ' : ''}{o}</button>
            );
          })}
        </div>
        {formData.desiredOutcomes.indexOf(OUTCOME_OBJECTION) !== -1 && (
          <div style={{ marginTop: 10 }}>
            <label style={LABEL_STYLE}>Specific objection(s) to overcome <span style={{ color: '#d13438' }}>*</span></label>
            <textarea rows={2} value={formData.objectionText} onChange={e => set('objectionText', e.target.value)}
              placeholder="List the specific objection(s) the SSE needs to address…"
              style={{ ...INPUT, resize: 'vertical', borderColor: !formData.objectionText.trim() ? '#d13438' : '#ccc', background: !formData.objectionText.trim() ? '#fef6f6' : '#fff' }} />
          </div>
        )}
        {formData.desiredOutcomes.indexOf(OUTCOME_OTHER) !== -1 && (
          <div style={{ marginTop: 10 }}>
            <label style={LABEL_STYLE}>Other desired outcome</label>
            <input type="text" value={formData.outcomeOtherText} onChange={e => set('outcomeOtherText', e.target.value)}
              placeholder="Describe the other outcome…" style={INPUT} />
          </div>
        )}
      </div>

      {/* Solution Landscape */}
      <div style={SECTION}>
        <SectionHeader title="Solution Landscape" hint="What we'd position, what the customer runs today, and whether we're expanding, integrating, or displacing. Pick 'None / Greenfield' for net-new." />
        <SolutionLandscape rows={formData.landscape} solutions={solutions} onChange={rows => set('landscape', rows)} />
      </div>

      {/* Priority */}
      <div style={SECTION}>
        <SectionHeader title="Priority" hint="How soon does this need to happen? High or Critical requires a reason." />
        <div style={{ display: 'flex', gap: 6 }}>
          {(['Low', 'Medium', 'High', 'Critical'] as const).map(p => {
            const active = formData.csePriority === p;
            const c = ({ Low: '#107c10', Medium: '#8a6000', High: '#a4262c', Critical: '#7a0b12' } as Record<string, string>)[p];
            return (
              <button key={p} type="button" onClick={() => set('csePriority', p)}
                style={{ padding: '5px 16px', fontSize: 12, fontWeight: 600, borderRadius: 4, cursor: 'pointer',
                  border: `1px solid ${active ? c : '#ccc'}`, background: active ? c : '#fff', color: active ? '#fff' : '#605e5c' }}>
                {p}
              </button>
            );
          })}
        </div>
        {(formData.csePriority === 'High' || formData.csePriority === 'Critical') && (
          <div style={{ marginTop: 10 }}>
            <label style={LABEL_STYLE}>Why is this {formData.csePriority.toLowerCase()} priority? <span style={{ color: '#d13438' }}>*</span></label>
            <textarea rows={2} value={formData.csePriorityReason} onChange={e => set('csePriorityReason', e.target.value)}
              placeholder="What's driving the urgency? (deal timeline, exec ask, competitive pressure, event/EOL date…)"
              style={{ ...INPUT, resize: 'vertical', borderColor: !formData.csePriorityReason.trim() ? '#d13438' : '#ccc', background: !formData.csePriorityReason.trim() ? '#fef6f6' : '#fff' }} />
          </div>
        )}
      </div>

      {/* Time Coordination */}
      <div style={SECTION}>
        <SectionHeader title="Time Coordination" />

        {/* ── SSE availability: the selected SSE's confirmed upcoming commitments (avoid double-booking) ── */}
        {sseEmail && (
          <div style={{ marginBottom: 14, padding: '10px 14px', borderRadius: 6,
            background: anyConflict ? '#fdf2f3' : '#f3f9f4',
            border: `1px solid ${anyConflict ? '#d68a90' : '#bcdcc4'}` }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#3b3a39', marginBottom: 6 }}>
              🗓️ {sseShortName || 'SSE'}’s upcoming commitments
              {commitLoading && <span style={{ fontWeight: 400, color: '#888' }}> — checking…</span>}
            </div>
            {!commitLoading && commitments.length === 0 && (
              <div style={{ fontSize: 12, color: '#605e5c' }}>No confirmed on-site or remote commitments coming up. 👍</div>
            )}
            {!commitLoading && commitments.length > 0 && (
              <>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                  {commitments.map((c, i) => {
                    const clash = commitConflicts(c);
                    return (
                      <div key={i} style={{ fontSize: 12, color: clash ? '#a4262c' : '#323130', fontWeight: clash ? 700 : 400 }}>
                        {c.type === 'On-site' ? '🏢' : '💻'} <strong>{fmtRange(c.start, c.end)}</strong> · {c.type}
                        {c.location ? ` · ${c.location}` : ''}
                        {clash && ' · ⚠️ overlaps your proposed dates'}
                      </div>
                    );
                  })}
                </div>
                {anyConflict && (
                  <div style={{ marginTop: 8, fontSize: 12, fontWeight: 700, color: '#a4262c' }}>
                    ⚠️ Your proposed dates overlap {sseShortName || 'the SSE'}’s existing commitment(s). Pick different dates or confirm this is intended.
                  </div>
                )}
              </>
            )}
          </div>
        )}

        <div style={FIELD_ROW}>
          <label style={LABEL_STYLE}>Format</label>
          <div style={{ display: 'flex', gap: 6 }}>
            {(['Remote', 'On-Site', 'Both'] as const).map(t => (
              <button key={t} type="button" onClick={() => set('supportType', t)} style={TOGGLE_BTN(formData.supportType === t)}>{t}</button>
            ))}
          </div>
        </div>
        {showRemote && (
          <ScheduleBlock label="Remote Schedule" start={formData.remoteStart} end={formData.remoteEnd} duration={formData.remoteDuration}
            tbd={formData.remoteTbd} presets={REMOTE_PRESETS}
            onStart={v => set('remoteStart', v)} onEnd={v => set('remoteEnd', v)} onDuration={v => set('remoteDuration', v)} onTbd={v => set('remoteTbd', v)} />
        )}
        {showOnsite && (
          <ScheduleBlock label="On-Site Schedule" start={formData.onsiteStart} end={formData.onsiteEnd} duration={formData.onsiteDuration} location={formData.location}
            tbd={formData.onsiteTbd} presets={ONSITE_PRESETS}
            onStart={v => set('onsiteStart', v)} onEnd={v => set('onsiteEnd', v)} onDuration={v => set('onsiteDuration', v)} onTbd={v => set('onsiteTbd', v)} onLocation={v => set('location', v)} />
        )}
        {!formData.supportType && <div style={{ fontSize: 12, color: '#888' }}>Pick a format above to set the schedule.</div>}
      </div>

      {/* Submit */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 4, marginBottom: 32 }}>
        <button disabled={submitting} onClick={handleSubmit}
          style={{ padding: '10px 28px', fontSize: 14, fontWeight: 700, background: submitting ? '#ccc' : HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, cursor: submitting ? 'not-allowed' : 'pointer' }}>
          {submitting ? 'Sending…' : '🎯 Request Strategic Engagement'}
        </button>
        <button type="button" disabled={submitting} onClick={() => { window.location.href = SRT_DASHBOARD_URL; }}
          style={{ padding: '10px 22px', fontSize: 14, fontWeight: 600, background: '#fff', color: '#605e5c', border: '1px solid #ccc', borderRadius: 4, cursor: submitting ? 'not-allowed' : 'pointer' }}>
          Cancel
        </button>
        <span style={{ fontSize: 12, color: '#605e5c' }}>Routes to the SSE for scheduling — lands in the SRT tracker</span>
      </div>
    </div>
  );
};
