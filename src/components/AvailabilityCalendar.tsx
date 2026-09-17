import * as React from 'react';
import { ISseCommitment } from '../models/ICseRequest';
import { HOURS_PER_DAY } from '../models/ScheduleBlock';
import { HPE_NAVY } from '../styles/hpe';

const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
const keyOf = (d: Date): string => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

// A real single-month wall calendar with ‹ › navigation. Days colored by hours booked:
// green = free, yellow = partly booked (some hours under a full day), red = full/over.
export const AvailabilityCalendar: React.FC<{ commitments: ISseCommitment[] }> = ({ commitments }) => {
  const [monthOffset, setMonthOffset] = React.useState(0);   // 0 = current month; nav range [-1, +3] (~90 days)

  // Sum booked hours (+ labels) per day across all commitments.
  const dayHours: Record<string, number> = {};
  const dayLabels: Record<string, string[]> = {};
  for (const c of commitments) {
    if (!c.start) continue;
    const hpd = (typeof c.hoursPerDay === 'number') ? c.hoursPerDay : HOURS_PER_DAY;
    const s = new Date(c.start.substring(0, 10) + 'T00:00:00');
    const e = new Date((c.end || c.start).substring(0, 10) + 'T00:00:00');
    const span = Math.max(0, Math.round((e.getTime() - s.getTime()) / 86400000));
    for (let i = 0; i <= span; i++) {
      const d = new Date(s); d.setDate(s.getDate() + i);
      const k = keyOf(d);
      dayHours[k] = (dayHours[k] || 0) + hpd;
      (dayLabels[k] = dayLabels[k] || []).push(`${d.getMonth() + 1}/${d.getDate()} — ${c.type === 'On-site' ? '📍 On-Site' : '💻 Remote/Prep'}${c.location ? ' · ' + c.location : ''} (${hpd}h)`);
    }
  }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayKey = keyOf(today);
  const base = new Date(today.getFullYear(), today.getMonth() + monthOffset, 1);
  const gridStart = new Date(base); gridStart.setDate(1 - base.getDay());   // Sunday on/before the 1st
  const lastDate = new Date(base.getFullYear(), base.getMonth() + 1, 0).getDate();
  const numRows = Math.ceil((base.getDay() + lastDate) / 7);
  const cells: Date[] = [];
  for (let i = 0; i < numRows * 7; i++) { const d = new Date(gridStart); d.setDate(gridStart.getDate() + i); cells.push(d); }

  const navBtn = (disabled: boolean): React.CSSProperties => ({
    fontSize: 16, lineHeight: 1, padding: '2px 10px', borderRadius: 4, border: '1px solid #ccc',
    background: '#fff', color: disabled ? '#ccc' : HPE_NAVY, cursor: disabled ? 'default' : 'pointer',
  });
  const swatch = (bg: string, br: string): React.CSSProperties => ({ display: 'inline-block', width: 10, height: 10, background: bg, border: `1px solid ${br}`, borderRadius: 2, verticalAlign: 'middle', marginRight: 4 });

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <button disabled={monthOffset <= -1} onClick={() => setMonthOffset(m => m - 1)} style={navBtn(monthOffset <= -1)}>‹</button>
        <div style={{ fontSize: 13, fontWeight: 700, color: HPE_NAVY, minWidth: 150, textAlign: 'center' }}>{MON[base.getMonth()]} {base.getFullYear()}</div>
        <button disabled={monthOffset >= 3} onClick={() => setMonthOffset(m => m + 1)} style={navBtn(monthOffset >= 3)}>›</button>
        {monthOffset !== 0 && (
          <button onClick={() => setMonthOffset(0)} style={{ fontSize: 11, padding: '3px 10px', borderRadius: 4, border: `1px solid ${HPE_NAVY}`, background: '#fff', color: HPE_NAVY, cursor: 'pointer' }}>Today</button>
        )}
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
        {DOW.map((h, i) => <div key={'h' + i} style={{ textAlign: 'center', fontSize: 11, fontWeight: 700, color: '#888', padding: '2px 0' }}>{h}</div>)}
        {cells.map(d => {
          const k = keyOf(d);
          const hrs = dayHours[k] || 0;
          const inMonth = d.getMonth() === base.getMonth();
          const isPast = d < today;
          const isToday = k === todayKey;
          const bg = !inMonth ? '#fafafa' : isPast ? '#f7f6f5' : hrs === 0 ? '#eef9f1' : hrs < HOURS_PER_DAY ? '#fff7d6' : '#fde7e9';
          const fg = !inMonth ? '#ccc' : isPast ? '#c0bdba' : hrs === 0 ? '#107c10' : hrs < HOURS_PER_DAY ? '#8a6000' : '#a4262c';
          return (
            <div key={k} title={inMonth && hrs ? `${hrs}h booked\n${dayLabels[k].join('\n')}` : (inMonth && !isPast ? 'Free' : '')}
              style={{ minHeight: 52, display: 'flex', flexDirection: 'column', padding: '4px 7px',
                background: bg, borderRadius: 5, border: isToday ? `2px solid ${HPE_NAVY}` : '1px solid #ececec',
                opacity: !inMonth ? 0.75 : (isPast ? 0.6 : 1) }}>
              <div style={{ fontSize: 12, fontWeight: isToday ? 800 : 600, color: fg }}>{d.getDate()}</div>
              {inMonth && hrs > 0 && <div style={{ marginTop: 'auto', fontSize: 10, fontWeight: 700, color: fg }}>{hrs}h</div>}
            </div>
          );
        })}
      </div>

      <div style={{ display: 'flex', gap: 14, marginTop: 8, fontSize: 10, color: '#605e5c', flexWrap: 'wrap' }}>
        <span><span style={swatch('#eef9f1', '#107c10')} />Free</span>
        <span><span style={swatch('#fff7d6', '#8a6000')} />Partly booked</span>
        <span><span style={swatch('#fde7e9', '#a4262c')} />Full day</span>
        <span style={{ color: '#888' }}>Use ‹ › to move month to month (~90 days). Hover a day for detail.</span>
      </div>
    </div>
  );
};
