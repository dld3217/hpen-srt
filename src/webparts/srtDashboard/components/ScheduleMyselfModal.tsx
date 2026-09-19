import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { CseRequestService } from '../../../services/CseRequestService';
import { ICseRequest } from '../../../models/ICseRequest';
import { IScheduleBlock, newBlock } from '../../../models/ScheduleBlock';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

const fmtShort = (iso: string): string => {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  if (!m) return '';
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
};

export interface IScheduleMyselfModalProps {
  sp: SPFI;
  context: WebPartContext;
  forUser?: string;        // "Name / email" of the EFFECTIVE user — set when an admin is viewing-as someone,
                           // so the block books that person (and demos work). Empty = book the real signed-in user.
  onClose: () => void;
  onCreated: () => void;   // ask the dashboard to reload (refreshes the availability calendar)
}

const LABEL: React.CSSProperties = { fontSize: 12, fontWeight: 700, color: HPE_NAVY, marginBottom: 4 };
const INPUT: React.CSSProperties = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '7px 10px', border: '1px solid #ccc', borderRadius: 4 };

// Personal (non-opportunity) time an SSE blocks on their OWN calendar — PTO, holidays, appointments,
// or free-form. Stored as a hidden EngagementType='Personal' row carrying a single self-booked (Prep)
// schedule block, so it holds the SSE's availability WITHOUT ever appearing on the request board.
const KINDS: { key: string; label: string; icon: string }[] = [
  { key: 'PTO',         label: 'PTO / Vacation',   icon: '🏖️' },
  { key: 'Holiday',     label: 'Holiday',          icon: '🎉' },
  { key: 'Appointment', label: 'Appointment',      icon: '🩺' },
  { key: 'Other',       label: 'Other / Free-form', icon: '📌' },
];

export const ScheduleMyselfModal: React.FC<IScheduleMyselfModalProps> = ({ sp, context, forUser, onClose, onCreated }) => {
  // Book for the effective (viewed-as) user when an admin is spoofing; otherwise the real signed-in user.
  const me = (forUser && forUser.trim()) ? forUser.trim() : `${context.pageContext.user.displayName} / ${context.pageContext.user.email}`;
  const [kind, setKind]       = useState('PTO');
  const [note, setNote]       = useState('');
  const [allDay, setAllDay]   = useState(true);
  const [start, setStart]     = useState('');
  const [end, setEnd]         = useState('');
  const [hours, setHours]     = useState('4');
  const [saving, setSaving]   = useState(false);
  const [error, setError]     = useState('');
  const [okMsg, setOkMsg]     = useState('');
  const [existing, setExisting] = useState<ICseRequest[]>([]);
  const [loadingList, setLoadingList] = useState(true);
  const [deletingId, setDeletingId]   = useState<number | null>(null);

  const myEmail = (me.split('/')[1] || '').trim().toLowerCase();

  // Load this person's own upcoming personal-time entries so they can be removed (e.g. an
  // appointment changed). Personal rows are hidden from the board, so this is the only manage path.
  const loadExisting = React.useCallback((): void => {
    setLoadingList(true);
    const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
    const now = new Date();
    const todayStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    new CseRequestService(sp).getAll()
      .then(all => {
        const mine = all.filter(r =>
          r.engagementType === 'Personal'
          && (r.requestedCse || '').toLowerCase().includes(myEmail)
          && (r.scheduleBlocks || []).some(b => !b.tbd && b.start && (b.end || b.start).substring(0, 10) >= todayStr));
        mine.sort((a, b) => {
          const as = (a.scheduleBlocks || [])[0]?.start || '';
          const bs = (b.scheduleBlocks || [])[0]?.start || '';
          return as < bs ? -1 : as > bs ? 1 : 0;
        });
        setExisting(mine);
        setLoadingList(false);
      })
      .catch(() => setLoadingList(false));
  }, [sp, myEmail]);

  useEffect(() => { loadExisting(); }, [loadExisting]);

  const handleDelete = async (id: number): Promise<void> => {
    setDeletingId(id);
    try {
      await new CseRequestService(sp).delete(id);
      setExisting(prev => prev.filter(r => r.id !== id));
      onCreated();   // refresh the dashboard availability views
    } finally { setDeletingId(null); }
  };

  const kindLabel = KINDS.filter(k => k.key === kind)[0]?.label || kind;
  const isOther   = kind === 'Other';
  // End defaults to start for a single-day entry; require a note only for free-form "Other".
  const effEnd    = allDay ? (end || start) : start;
  const hoursNum  = parseFloat(hours) || 0;
  const canSave   = !!start
    && (!allDay || (effEnd >= start))
    && (allDay || hoursNum > 0)
    && (!isOther || !!note.trim());

  const handleSave = async (): Promise<void> => {
    if (!canSave || saving) return;
    setSaving(true); setError(''); setOkMsg('');
    try {
      const displayLabel = `${kindLabel}${note.trim() ? ' — ' + note.trim() : ''}`;
      // A self-booked Prep block: full days over the range, or a set number of hours on one day.
      const block: IScheduleBlock = {
        ...newBlock('Prep'),
        tbd: false,
        start,
        end: allDay ? effEnd : start,
        label: displayLabel,
        unit: allDay ? 'days' : 'hours',
        hours: allDay ? undefined : hoursNum,
      };
      await new CseRequestService(sp).create({
        title: displayLabel,
        source: 'Personal Time',
        linkedPocId: 0,
        requestStatus: 'In Progress',   // counts toward availability immediately; hidden from the board
        scheduleStatus: 'TBD',
        requestedCse: me,               // this is MY time
        sseManagerEmail: '',
        cseDescription: '',
        csePriority: 'Low',
        csePriorityReason: '',
        solutionsFocus: '',
        supportType: '',
        remoteTbd: true, remoteStart: '', remoteEnd: '', remoteDuration: '',
        onsiteTbd: true, onsiteStart: '', onsiteEnd: '', onsiteDuration: '', onsiteDestination: '',
        sePrimary: '',
        semPrimary: '',
        sedEmail: '',
        buRegion: '',                   // non-geo — never in any BU/Region rollup
        hpenBusinessUnit: '',
        customerName: displayLabel,
        pocName: '',
        opportunityAmount: 0,
        custTemp: 'Normal',
        signedOffBy: '',
        signOffDate: '',
        notes: '',
        engagementType: 'Personal',
        scheduleBlocks: [block],
      });
      onCreated();                 // refresh the dashboard availability views
      // Stay open so you can add more or remove one — reset the form + refresh the list below.
      setOkMsg(`✓ Added: ${displayLabel} (${fmtShort(start)}${allDay && effEnd !== start ? ` – ${fmtShort(effEnd)}` : ''}).`);
      setNote(''); setStart(''); setEnd(''); setHours('4');
      setSaving(false);
      loadExisting();
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
        width: 'min(480px, 94vw)', maxHeight: '90vh', overflowY: 'auto',
        background: '#fff', borderRadius: 8, boxShadow: '0 12px 40px rgba(0,0,0,0.28)',
        zIndex: 1001, fontFamily: 'inherit',
      }}>
        {/* Header */}
        <div style={{ background: HPE_NAVY, color: '#fff', padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderRadius: '8px 8px 0 0' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 3, height: 20, background: HPE_GREEN, borderRadius: 2 }} />
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>🗓️ Schedule Myself</div>
              <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>Block personal time on your availability — not tied to any opportunity</div>
            </div>
          </div>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer', lineHeight: 1, padding: 4 }}>✕</button>
        </div>

        <div style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: 14 }}>

          {/* Kind */}
          <div>
            <div style={LABEL}>Type</div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
              {KINDS.map(k => {
                const active = kind === k.key;
                return (
                  <button key={k.key} onClick={() => setKind(k.key)}
                    style={{ padding: '8px 10px', fontSize: 12, fontWeight: 600, cursor: 'pointer', textAlign: 'left',
                      border: `1px solid ${active ? HPE_GREEN : '#ccc'}`, borderRadius: 4,
                      background: active ? HPE_GREEN : '#fff', color: active ? '#fff' : '#605e5c' }}>
                    {k.icon} {k.label}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Note / free-form */}
          <div>
            <div style={LABEL}>{isOther ? 'Description' : 'Note'} {isOther ? <span style={{ color: '#a4262c' }}>*</span> : <span style={{ fontWeight: 400, color: '#888' }}>(optional)</span>}</div>
            <input value={note} onChange={e => setNote(e.target.value)}
              placeholder={isOther ? 'What are you blocking time for?' : 'e.g. Family vacation, dentist…'} style={INPUT} />
          </div>

          {/* All-day vs partial */}
          <div>
            <div style={{ display: 'inline-flex', border: `1px solid ${HPE_NAVY}`, borderRadius: 4, overflow: 'hidden' }}>
              {([[true, 'All day(s)'], [false, 'Partial (hours)']] as const).map(([v, lbl]) => (
                <button key={String(v)} type="button" onClick={() => setAllDay(v)}
                  style={{ fontSize: 12, fontWeight: 600, padding: '5px 14px', border: 'none', cursor: 'pointer',
                    background: allDay === v ? HPE_NAVY : '#fff', color: allDay === v ? '#fff' : '#605e5c' }}>
                  {lbl}
                </button>
              ))}
            </div>
          </div>

          {/* Dates */}
          {allDay ? (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <div>
                <div style={LABEL}>Start</div>
                <input type="date" value={start} onChange={e => setStart(e.target.value)} style={INPUT} />
              </div>
              <div>
                <div style={LABEL}>End <span style={{ fontWeight: 400, color: '#888' }}>(same day if blank)</span></div>
                <input type="date" value={end} min={start} onChange={e => setEnd(e.target.value)} style={INPUT} />
              </div>
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <div>
                <div style={LABEL}>Date</div>
                <input type="date" value={start} onChange={e => setStart(e.target.value)} style={INPUT} />
              </div>
              <div>
                <div style={LABEL}>Hours</div>
                <input type="number" min={0.5} step={0.5} value={hours} onChange={e => setHours(e.target.value)} style={INPUT} />
              </div>
            </div>
          )}

          <div style={{ fontSize: 11, color: '#888', background: '#f6f5f4', borderRadius: 4, padding: '8px 10px' }}>
            This blocks <strong>your own</strong> availability only. It will not appear on the request board — just on your
            calendar and the team availability views, so SEs will not schedule you during this time.
          </div>

          {error && (
            <div style={{ fontSize: 12, color: '#a4262c', background: '#fde7e9', border: '1px solid #a4262c', borderRadius: 4, padding: '8px 10px' }}>
              {error}
            </div>
          )}
          {okMsg && (
            <div style={{ fontSize: 12, color: '#107c10', background: '#dff6dd', border: '1px solid #107c10', borderRadius: 4, padding: '8px 10px' }}>
              {okMsg}
            </div>
          )}

          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => { handleSave().catch(() => undefined); }} disabled={!canSave || saving}
              style={{ flex: 1, padding: '9px 0', background: canSave && !saving ? HPE_GREEN : '#c8c6c4', color: '#fff', border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 700, cursor: canSave && !saving ? 'pointer' : 'not-allowed' }}>
              {saving ? 'Saving…' : 'Block This Time'}
            </button>
            <button onClick={onClose} disabled={saving}
              style={{ padding: '9px 20px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 13, cursor: 'pointer' }}>
              Done
            </button>
          </div>

          {/* ── Your scheduled personal time — remove an entry if plans change ── */}
          <div style={{ marginTop: 6, paddingTop: 12, borderTop: '1px solid #edebe9' }}>
            <div style={{ ...LABEL, marginBottom: 6 }}>Your scheduled personal time</div>
            {loadingList ? (
              <div style={{ fontSize: 12, color: '#888' }}>Loading…</div>
            ) : existing.length === 0 ? (
              <div style={{ fontSize: 12, color: '#888' }}>Nothing blocked yet.</div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                {existing.map(r => {
                  const b = (r.scheduleBlocks || [])[0];
                  const when = b ? `${fmtShort(b.start)}${b.end && b.end !== b.start ? ` – ${fmtShort(b.end)}` : ''}${b.unit === 'hours' && b.hours ? ` · ${b.hours}h` : ''}` : '';
                  return (
                    <div key={r.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, background: '#faf9f8', border: '1px solid #edebe9', borderRadius: 4, padding: '6px 8px' }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, color: HPE_NAVY, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{b?.label || r.customerName || 'Personal time'}</div>
                        {when && <div style={{ fontSize: 11, color: '#605e5c' }}>{when}</div>}
                      </div>
                      <button onClick={() => { handleDelete(r.id!).catch(() => undefined); }} disabled={deletingId === r.id}
                        title="Remove this personal time" style={{ flexShrink: 0, fontSize: 11, padding: '3px 10px', background: '#fde7e9', color: '#a4262c', border: '1px solid #a4262c', borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>
                        {deletingId === r.id ? '…' : '✕ Remove'}
                      </button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
};
