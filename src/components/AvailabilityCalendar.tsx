import * as React from 'react';
import { ISseCommitment } from '../models/ICseRequest';
import { HPE_NAVY } from '../styles/hpe';

const DOW = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
const keyOf = (d: Date): string => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

// A rolling green/red month grid: green = free, red = booked, for an at-a-glance availability read.
export const AvailabilityCalendar: React.FC<{ commitments: ISseCommitment[]; weeks?: number }> = ({ commitments, weeks = 6 }) => {
  // Expand each commitment's date range into per-day busy labels.
  const busy: Record<string, string[]> = {};
  for (const c of commitments) {
    if (!c.start) continue;
    const s = new Date(c.start.substring(0, 10) + 'T00:00:00');
    const e = new Date((c.end || c.start).substring(0, 10) + 'T00:00:00');
    const span = Math.max(0, Math.round((e.getTime() - s.getTime()) / 86400000));
    for (let i = 0; i <= span; i++) {
      const d = new Date(s); d.setDate(s.getDate() + i);
      const k = keyOf(d);
      (busy[k] = busy[k] || []).push(`${MON[d.getMonth()]} ${d.getDate()} — ${c.type === 'On-site' ? '📍 On-Site' : '💻 Remote/Prep'}${c.location ? ' · ' + c.location : ''}`);
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
        {MON[first.getMonth()]} {first.getDate()} – {MON[last.getMonth()]} {last.getDate()}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 3, maxWidth: 340 }}>
        {DOW.map((h, i) => <div key={'h' + i} style={{ textAlign: 'center', fontSize: 10, fontWeight: 700, color: '#888' }}>{h}</div>)}
        {cells.map(d => {
          const k = keyOf(d);
          const isBusy = !!busy[k];
          const isPast = d < today;
          const isToday = k === todayKey;
          const bg = isPast ? '#f3f2f1' : isBusy ? '#fde7e9' : '#e8faf3';
          const fg = isPast ? '#bbb' : isBusy ? '#a4262c' : '#107c10';
          return (
            <div key={k} title={isBusy ? busy[k].join('\n') : (isPast ? '' : 'Free')}
              style={{ height: 26, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11,
                background: bg, color: fg, borderRadius: 3, fontWeight: isToday ? 800 : 500,
                border: isToday ? `2px solid ${HPE_NAVY}` : '1px solid #fff', opacity: isPast ? 0.55 : 1 }}>
              {d.getDate()}
            </div>
          );
        })}
      </div>
      <div style={{ display: 'flex', gap: 14, marginTop: 8, fontSize: 10, color: '#605e5c' }}>
        <span><span style={swatch('#e8faf3', '#107c10')} />Free</span>
        <span><span style={swatch('#fde7e9', '#a4262c')} />Booked</span>
        <span style={{ color: '#888' }}>Hover a day to see what&rsquo;s booked.</span>
      </div>
    </div>
  );
};
