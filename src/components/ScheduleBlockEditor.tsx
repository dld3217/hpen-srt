import * as React from 'react';
import {
  IScheduleBlock, ScheduleBlockType, SCHEDULE_BLOCK_TYPES, BLOCK_STYLE,
  newBlock, blockDays, plannedHours,
} from '../models/ScheduleBlock';
import { HPE_NAVY } from '../styles/hpe';

export interface IScheduleBlockEditorProps {
  blocks: IScheduleBlock[];
  onChange: (blocks: IScheduleBlock[]) => void;
  allowTypes?: ScheduleBlockType[];   // default: all three
  showAccounting?: boolean;           // show the Planned→Actual toggle + actual-hours (dashboard)
}

const FLD: React.CSSProperties = { fontSize: 11, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, boxSizing: 'border-box' };
const XBTN: React.CSSProperties = { background: 'none', border: 'none', color: '#d13438', fontSize: 15, cursor: 'pointer', lineHeight: 1, padding: '0 2px' };

// Repeatable, typed time-block editor. Each block is an independent date range (contiguous or not).
// Planning always; accounting (Planned→Actual switch + optional actual hours) when showAccounting.
export const ScheduleBlockEditor: React.FC<IScheduleBlockEditorProps> = ({ blocks, onChange, allowTypes, showAccounting }) => {
  const types = (allowTypes && allowTypes.length) ? allowTypes : SCHEDULE_BLOCK_TYPES;

  const update = (id: string, patch: Partial<IScheduleBlock>): void =>
    onChange(blocks.map(b => b.id === id ? { ...b, ...patch } : b));
  const add = (t: ScheduleBlockType): void => onChange([...blocks, newBlock(t)]);
  const remove = (id: string): void => onChange(blocks.filter(b => b.id !== id));

  const totalDays = blocks.reduce((s, b) => s + blockDays(b), 0);

  return (
    <div>
      {blocks.length === 0 && (
        <div style={{ fontSize: 12, color: '#888', fontStyle: 'italic', marginBottom: 8 }}>
          No time blocks yet — add the SSE&rsquo;s time below. Each block is its own date range, so they can be back-to-back or spread out.
        </div>
      )}

      {blocks.map(b => {
        const st = BLOCK_STYLE[b.type];
        return (
          <div key={b.id} style={{ border: `1px solid ${st.color}33`, borderLeft: `3px solid ${st.color}`, borderRadius: 6, padding: '8px 10px', marginBottom: 8, background: '#fff' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
              {types.length > 1 ? (
                <select value={b.type}
                  onChange={e => { const t = e.target.value as ScheduleBlockType; update(b.id, { type: t, location: t === 'On-Site' ? b.location : '' }); }}
                  style={{ ...FLD, fontWeight: 700, color: st.color }}>
                  {types.map(t => <option key={t} value={t}>{BLOCK_STYLE[t].icon} {t}</option>)}
                </select>
              ) : (
                <span style={{ fontSize: 12, fontWeight: 700, color: st.color }}>{st.icon} {b.type}</span>
              )}
              <input value={b.label} onChange={e => update(b.id, { label: e.target.value })}
                placeholder="label (optional)" style={{ ...FLD, flex: 1 }} />
              <button onClick={() => remove(b.id)} title="Remove block" style={XBTN}>✕</button>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <label style={{ fontSize: 11, display: 'flex', gap: 4, alignItems: 'center', cursor: 'pointer', color: '#605e5c' }}>
                <input type="checkbox" checked={b.tbd} onChange={e => update(b.id, { tbd: e.target.checked })} /> Dates TBD
              </label>
              {!b.tbd && (
                <>
                  <input type="date" value={b.start}
                    onChange={e => update(b.id, { start: e.target.value, end: (b.end && b.end < e.target.value) ? e.target.value : b.end })}
                    style={FLD} />
                  <span style={{ fontSize: 11, color: '#888' }}>→</span>
                  <input type="date" value={b.end || b.start} min={b.start}
                    onChange={e => update(b.id, { end: e.target.value })} style={FLD} />
                  <span style={{ fontSize: 11, color: st.color, fontWeight: 600 }}>{blockDays(b)}d</span>
                </>
              )}
              {b.type === 'On-Site' && (
                <input value={b.location} onChange={e => update(b.id, { location: e.target.value })}
                  placeholder="location" style={{ ...FLD, minWidth: 120 }} />
              )}
            </div>

            {showAccounting && (
              <div style={{ marginTop: 6, display: 'flex', alignItems: 'center', gap: 8, borderTop: '1px dashed #e0e0e0', paddingTop: 6, flexWrap: 'wrap' }}>
                <label style={{ fontSize: 11, display: 'flex', gap: 4, alignItems: 'center', cursor: 'pointer', fontWeight: 600, color: b.logged ? '#107c10' : '#605e5c' }}>
                  <input type="checkbox" checked={b.logged} onChange={e => update(b.id, { logged: e.target.checked })} />
                  {b.logged ? '✓ Actual (logged)' : 'Mark actual'}
                </label>
                {b.logged && (
                  <>
                    <input type="number" min={0} step={0.5} value={b.actualHours === undefined ? '' : b.actualHours}
                      placeholder={String(plannedHours(b))}
                      onChange={e => update(b.id, { actualHours: e.target.value === '' ? undefined : Number(e.target.value) })}
                      style={{ ...FLD, width: 64 }} />
                    <span style={{ fontSize: 11, color: '#888' }}>actual hrs (blank = planned {plannedHours(b)}h)</span>
                  </>
                )}
              </div>
            )}
          </div>
        );
      })}

      <div style={{ display: 'flex', gap: 6, alignItems: 'center', marginTop: 4, flexWrap: 'wrap' }}>
        {types.map(t => (
          <button key={t} onClick={() => add(t)}
            style={{ fontSize: 11, padding: '4px 10px', background: '#fff', color: BLOCK_STYLE[t].color, border: `1px solid ${BLOCK_STYLE[t].color}`, borderRadius: 4, cursor: 'pointer', fontWeight: 600 }}>
            + {BLOCK_STYLE[t].icon} {t}
          </button>
        ))}
        {totalDays > 0 && (
          <span style={{ marginLeft: 'auto', fontSize: 11, color: HPE_NAVY }}>Planned total: <strong>{totalDays}d</strong></span>
        )}
      </div>
    </div>
  );
};
