import * as React from 'react';
import { ISseCommitment } from '../models/ICseRequest';
import { HOURS_PER_DAY } from '../models/ScheduleBlock';
import { HPE_NAVY } from '../styles/hpe';

const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
const keyOf = (d: Date): string => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

// A rolling month grid colored by daily load: green = free, yellow = partly booked
// (some hours but under a full day), red = full/over. ~90-day horizon by default.
export const AvailabilityCalendar: React.FC<{ commitments: ISseCommitment[]; weeks?: number }> = ({ commitments, weeks = 13 }) => {
  // Sum booked hours (and labels) per day across all commitments.
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
      (dayLabels[k] = dayLabels[k] || []).push(`${MON[d.getMonth()]} ${d.getDate()} — ${c.type === 'On-site' ? '📍 On-Site' : '💻 Remote/Prep'}${c.location ? ' · ' + c.location : ''} (${hpd}h)`);
    }
  }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const start = new Date(today); start.setDate(today.getDate() - today.getDay());   // back to Sunday
  const todayKey = keyOf(today);

  const cells: Date[] = [];
  for (let i = 0; i < weeks * 7; i++) { const d = new Date(start); d.setDate(start.getDate() + i); cells.push(d); }
  const first = cells[0], last = cells[cells.length - 1];

  const swatch = (bg: string, br: string): React.CSSProperties => ({ display: 'inline-block', width: 10, height: 10, background: bg, border: `1px solid ${br}`, borderRadius: 2, verticalAlign: 'middle', marginRight: 4 });

  return (
    <div>
      <div style={{ fontSize: 11, color: '#605e5c', marginBottom: 6 }}>
        {MON[first.getMonth()]} {first.getDate()} – {MON[last.getMonth()]} {last.getDate()} · next {weeks * 7} days
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
        {DOW.map((h, i) => <div key={'h' + i} style={{ textAlign: 'center', fontSize: 11, fontWeight: 700, color: '#888', padding: '2px 0' }}>{h}</div>)}
        {cells.map(d => {
          const k = keyOf(d);
          const hrs = dayHours[k] || 0;
          const isPast = d < today;
          const isToday = k === todayKey;
          const firstOfMonth = d.getDate() === 1;
          // green = free, yellow = partial (0 < h < full day), red = full/over
          const bg = isPast ? '#f7f6f5' : hrs === 0 ? '#eef9f1' : hrs < HOURS_PER_DAY ? '#fff7d6' : '#fde7e9';
          const fg = isPast ? '#c0bdba' : hrs === 0 ? '#107c10' : hrs < HOURS_PER_DAY ? '#8a6000' : '#a4262c';
          return (
            <div key={k} title={hrs ? `${hrs}h booked\n${dayLabels[k].join('\n')}` : (isPast ? '' : 'Free')}
              style={{ minHeight: 54, display: 'flex', flexDirection: 'column', padding: '4px 7px',
                background: bg, borderRadius: 5, border: isToday ? `2px solid ${HPE_NAVY}` : '1px solid #ececec',
                opacity: isPast ? 0.6 : 1 }}>
              <div style={{ fontSize: 12, fontWeight: (isToday || firstOfMonth) ? 800 : 600, color: fg }}>
                {firstOfMonth ? `${MON[d.getMonth()]} 1` : d.getDate()}
              </div>
              {hrs > 0 && <div style={{ marginTop: 'auto', fontSize: 10, fontWeight: 700, color: fg }}>{hrs}h</div>}
            </div>
          );
        })}
      </div>
      <div style={{ display: 'flex', gap: 14, marginTop: 8, fontSize: 10, color: '#605e5c', flexWrap: 'wrap' }}>
        <span><span style={swatch('#e8faf3', '#107c10')} />Free</span>
        <span><span style={swatch('#fff4ce', '#8a6000')} />Partly booked</span>
        <span><span style={swatch('#fde7e9', '#a4262c')} />Full day</span>
        <span style={{ color: '#888' }}>Hover a day for hours &amp; bookings.</span>
      </div>
    </div>
  );
};
