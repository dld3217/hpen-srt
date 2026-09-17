import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { CseRequestService } from '../../../services/CseRequestService';
import { ConfigService, SpecialProjectsMap } from '../../../services/ConfigService';
import { ScheduleBlockEditor } from '../../../components/ScheduleBlockEditor';
import { IScheduleBlock, demoScheduleBlocks } from '../../../models/ScheduleBlock';
import { ISseCommitment } from '../../../models/ICseRequest';
import { PeoplePickerField, searchGraphUsers } from '../../../components/PeoplePickerField';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface INewSpecialProjectModalProps {
  sp: SPFI;
  context: WebPartContext;
  showDemo?: boolean;      // gates the 🧪 demo-fill (admins + SSEs)
  onClose: () => void;
  onCreated: () => void;   // ask the dashboard to reload
}

const PRIORITIES = ['Low', 'Medium', 'High', 'Critical'];
const LABEL: React.CSSProperties = { fontSize: 12, fontWeight: 700, color: HPE_NAVY, marginBottom: 4 };
const INPUT: React.CSSProperties = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '7px 10px', border: '1px solid #ccc', borderRadius: 4 };

// Admin/Charlie-side creation of a non-geo Special Project engagement (CIC / Marketing).
// Category + Initiative come from the shared SRTSpecialProjects config; a brand-new initiative
// typed here is persisted back so the next requester picks the SAME one (→ it aggregates).
// One CSERequests row is created per staffed SSE — cross-geo by design; no BU/Region set.
export const NewSpecialProjectModal: React.FC<INewSpecialProjectModalProps> = ({ sp, context, showDemo, onClose, onCreated }) => {
  const me = `${context.pageContext.user.displayName} / ${context.pageContext.user.email}`;
  const [spMap, setSpMap]           = useState<SpecialProjectsMap>({});
  const [loading, setLoading]       = useState(true);
  const [category, setCategory]     = useState('');
  const [initiative, setInitiative] = useState('');
  const [title, setTitle]           = useState('');
  const [priority, setPriority]     = useState('Medium');
  const [description, setDescription] = useState('');
  const [sses, setSses]             = useState<string[]>([me]);   // default the first SSE to the creator
  const searchUsers = React.useCallback((q: string) => searchGraphUsers(context, q), [context]);
  const [saving, setSaving]         = useState(false);
  const [error, setError]           = useState('');
  const [blocks, setBlocks]         = useState<IScheduleBlock[]>([]);
  const [commitments, setCommitments] = useState<ISseCommitment[]>([]);

  const configSvc = React.useMemo(() => new ConfigService(sp), [sp]);

  useEffect(() => {
    configSvc.getSpecialProjects()
      .then(m => { setSpMap(m); setCategory(Object.keys(m)[0] || ''); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  // Load the first SSE's existing commitments → free/busy + conflict flags in the schedule editor.
  useEffect(() => {
    const sseEmail = (sses[0] ? (sses[0].split('/')[1] || '') : '').trim().toLowerCase();
    if (!sseEmail) { setCommitments([]); return; }
    new CseRequestService(sp).getSseCommitments(sseEmail).then(setCommitments).catch(() => setCommitments([]));
  }, [sses]);

  const initiatives = category ? (spMap[category] || []) : [];
  const setSse = (i: number, v: string): void => setSses(prev => prev.map((s, k) => k === i ? v : s));
  const addSseField = (): void => setSses(prev => [...prev, '']);
  const removeSseField = (i: number): void => setSses(prev => prev.length > 1 ? prev.filter((_, k) => k !== i) : prev);

  const canSave = !!category && !!initiative.trim() && sses.some(s => s.trim());

  // Admin demo-fill — populates the whole form + schedule, tagged [SAMPLE] so it's obviously demo.
  const fillDemo = (): void => {
    const cats = Object.keys(spMap);
    const cat = cats.filter(c => /houston/i.test(c))[0] || cats[0] || 'Houston CIC';
    const inits = spMap[cat] || [];
    const init = inits[0] || 'Bee Counting';
    setCategory(cat);
    setInitiative(init);
    setTitle(`[SAMPLE] ${cat} — ${init}`);
    setPriority('High');
    setDescription(`[SAMPLE] Demo Special Project — SSE supporting the ${init} effort (design review, lab build/prep, and on-site delivery).`);
    setSses(['Charlie Clemmer / charlie.clemmer@hpe.com']);
    setBlocks(demoScheduleBlocks());
  };

  const handleSave = async (): Promise<void> => {
    if (!canSave || saving) return;
    setSaving(true); setError('');
    try {
      const initName = initiative.trim();
      // Persist a brand-new initiative to the shared config so it becomes a picker option for everyone.
      const existing = spMap[category] || [];
      if (!existing.some(i => i.toLowerCase() === initName.toLowerCase())) {
        await configSvc.saveSpecialProjects({ ...spMap, [category]: [...existing, initName] });
      }
      const finalTitle = title.trim() || `${category} — ${initName}`;
      const svc = new CseRequestService(sp);
      const cleanSses = sses.map(s => s.trim()).filter(Boolean);
      // One row per SSE — cross-geo staffing; they aggregate under the same Category/Initiative.
      // Schedule blocks attach to the FIRST SSE (the creator's own time); others schedule on the dashboard.
      for (let i = 0; i < cleanSses.length; i++) {
        const sse = cleanSses[i];
        await svc.create({
          title: finalTitle,
          source: 'Special Project',
          linkedPocId: 0,
          requestStatus: 'Accepted',
          scheduleStatus: 'TBD',
          requestedCse: sse,
          sseManagerEmail: '',
          cseDescription: description,
          csePriority: priority,
          csePriorityReason: '',
          solutionsFocus: '',
          supportType: '',
          remoteTbd: true, remoteStart: '', remoteEnd: '', remoteDuration: '',
          onsiteTbd: true, onsiteStart: '', onsiteEnd: '', onsiteDuration: '', onsiteDestination: '',
          sePrimary: '',
          semPrimary: '',
          sedEmail: '',
          buRegion: '',            // non-geo — keeps Special Projects out of geo rollups
          hpenBusinessUnit: '',
          customerName: finalTitle, // shows the project subject in the Customer column
          pocName: '',
          opportunityAmount: 0,
          custTemp: 'Normal',
          signedOffBy: '',
          signOffDate: '',
          notes: '',
          engagementType: 'Special Project',
          specialProjectCategory: category,
          specialProjectInitiative: initName,
          scheduleBlocks: i === 0 ? blocks : [],
        });
      }
      onCreated();
      onClose();
    } catch (e) {
      setError((e as Error).message);
      setSaving(false);
    }
  };

  return (
    <>
      <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.4)', zIndex: 1000 }} />
      <div style={{
        position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)',
        width: 'min(560px, 94vw)', maxHeight: '90vh', overflowY: 'auto',
        background: '#fff', borderRadius: 8, boxShadow: '0 12px 40px rgba(0,0,0,0.28)',
        zIndex: 1001, fontFamily: 'inherit',
      }}>
        {/* Header */}
        <div style={{ background: HPE_NAVY, color: '#fff', padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderRadius: '8px 8px 0 0' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 3, height: 20, background: HPE_GREEN, borderRadius: 2 }} />
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>⭐ New Special Project</div>
              <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>CIC / Marketing — non-geo engagement</div>
            </div>
          </div>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer', lineHeight: 1, padding: 4 }}>✕</button>
        </div>

        <div style={{ padding: '18px 20px' }}>
          {loading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>

              {showDemo && (
                <button onClick={fillDemo} title="Fill a tagged [SAMPLE] demo special project + schedule"
                  style={{ alignSelf: 'flex-start', fontSize: 12, padding: '5px 12px', background: '#f3e8ff', color: '#6b2faf', border: '1px solid #6b2faf', borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>
                  🧪 Demo fill
                </button>
              )}

              {/* Category */}
              <div>
                <div style={LABEL}>Category</div>
                {Object.keys(spMap).length === 0 ? (
                  <div style={{ fontSize: 12, color: '#a4262c' }}>No Special Project categories configured yet — add one in Admin → ⭐ Special.</div>
                ) : (
                  <select value={category} onChange={e => { setCategory(e.target.value); setInitiative(''); }} style={INPUT}>
                    {Object.keys(spMap).map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                )}
              </div>

              {/* Initiative — type-or-pick (typeahead of existing + open add) */}
              <div>
                <div style={LABEL}>Initiative</div>
                <input list="sp-initiatives" value={initiative} onChange={e => setInitiative(e.target.value)}
                  placeholder="Pick an existing initiative or type a new one" style={INPUT} />
                <datalist id="sp-initiatives">
                  {initiatives.map(i => <option key={i} value={i} />)}
                </datalist>
                <div style={{ fontSize: 11, color: '#888', marginTop: 4 }}>
                  Typing a new name adds it to the picker for everyone under <strong>{category || 'this category'}</strong> so other SSEs join the same initiative.
                </div>
              </div>

              {/* Title (optional) */}
              <div>
                <div style={LABEL}>Title <span style={{ fontWeight: 400, color: '#888' }}>(optional)</span></div>
                <input value={title} onChange={e => setTitle(e.target.value)}
                  placeholder={category && initiative ? `${category} — ${initiative}` : 'Short description of the work'} style={INPUT} />
              </div>

              {/* Priority */}
              <div>
                <div style={LABEL}>Priority</div>
                <div style={{ display: 'flex', gap: 6 }}>
                  {PRIORITIES.map(p => {
                    const active = priority === p;
                    const color = p === 'Critical' ? '#a4262c' : p === 'High' ? '#d83b01' : p === 'Medium' ? '#8a6000' : '#107c10';
                    return (
                      <button key={p} onClick={() => setPriority(p)}
                        style={{ flex: 1, padding: '6px 0', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                          border: `1px solid ${active ? color : '#ccc'}`, borderRadius: 4,
                          background: active ? color : '#fff', color: active ? '#fff' : '#605e5c' }}>
                        {p}
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Staffed SSEs */}
              <div>
                <div style={LABEL}>Staffed SSE(s)</div>
                <div style={{ fontSize: 11, color: '#888', marginBottom: 6 }}>One tracker row is created per SSE — they can be from any region.</div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {sses.map((s, i) => (
                    <div key={i} style={{ display: 'flex', gap: 6, alignItems: 'flex-start' }}>
                      <PeoplePickerField value={s} onChange={v => setSse(i, v)} searchUsers={searchUsers} placeholder="Search SSE name…" />
                      {sses.length > 1 && (
                        <button onClick={() => removeSseField(i)} title="Remove"
                          style={{ background: 'none', border: '1px solid #ccc', borderRadius: 4, color: '#d13438', fontSize: 15, cursor: 'pointer', padding: '0 10px' }}>✕</button>
                      )}
                    </div>
                  ))}
                </div>
                <button onClick={addSseField}
                  style={{ marginTop: 6, fontSize: 12, padding: '4px 12px', background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>
                  + Add another SSE
                </button>
              </div>

              {/* Description */}
              <div>
                <div style={LABEL}>Description <span style={{ fontWeight: 400, color: '#888' }}>(optional)</span></div>
                <textarea value={description} onChange={e => setDescription(e.target.value)} rows={3}
                  placeholder="What is the SSE being asked to help with?" style={{ ...INPUT, resize: 'vertical' }} />
              </div>

              {/* Schedule — attaches to the first SSE (the creator's own time) */}
              <div>
                <div style={LABEL}>Schedule <span style={{ fontWeight: 400, color: '#888' }}>(optional — your Prep &amp; On-Site time; applies to the first SSE)</span></div>
                <ScheduleBlockEditor blocks={blocks} onChange={setBlocks} allowTypes={['Prep', 'On-Site']} commitments={commitments} />
              </div>

              {error && (
                <div style={{ fontSize: 12, color: '#a4262c', background: '#fde7e9', border: '1px solid #a4262c', borderRadius: 4, padding: '8px 10px' }}>
                  {error}
                </div>
              )}

              <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                <button onClick={() => { handleSave().catch(() => undefined); }} disabled={!canSave || saving}
                  style={{ flex: 1, padding: '9px 0', background: canSave && !saving ? HPE_GREEN : '#c8c6c4', color: '#fff', border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 700, cursor: canSave && !saving ? 'pointer' : 'not-allowed' }}>
                  {saving ? 'Creating…' : 'Create Special Project'}
                </button>
                <button onClick={onClose} disabled={saving}
                  style={{ padding: '9px 20px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 13, cursor: 'pointer' }}>
                  Cancel
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </>
  );
};
