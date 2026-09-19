import * as React from 'react';
import { ISseCommitment } from '../models/ICseRequest';
import { HOURS_PER_DAY } from '../models/ScheduleBlock';
import { HPE_NAVY } from '../styles/hpe';

const DOW = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
const keyOf = (d: Date): string => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

// Amber hatch = "on hold" (dates proposed, pending both SED + SSE approval).
const HOLD_HATCH = 'repeating-linear-gradient(45deg, #fff3d6, #fff3d6 4px, #ffe4a3 4px, #ffe4a3 8px)';

// One month's grid. Firm bookings: green = free, yellow = partial, red = full/over.
// Days that are only TENTATIVE (no firm hours) render amber-hatched — "on hold", not yet booked.
const MonthGrid: React.FC<{
  base: Date; dayHours: Record<string, number>; dayTent: Record<string, number>; dayLabels: Record<string, string[]>;
  dayReq: Record<string, number>; onOpenRequest?: (requestId: number) => void; today: Date; todayKey: string;
}> = ({ base, dayHours, dayTent, dayLabels, dayReq, onOpenRequest, today, todayKey }) => {
  const gridStart = new Date(base); gridStart.setDate(1 - base.getDay());
  const lastDate = new Date(base.getFullYear(), base.getMonth() + 1, 0).getDate();
  const numRows = Math.ceil((base.getDay() + lastDate) / 7);
  const cells: Date[] = [];
  for (let i = 0; i < numRows * 7; i++) { const d = new Date(gridStart); d.setDate(gridStart.getDate() + i); cells.push(d); }

  return (
    <div style={{ flex: '1 1 250px', minWidth: 230 }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: HPE_NAVY, textAlign: 'center', marginBottom: 6 }}>{MON[base.getMonth()]} {base.getFullYear()}</div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 3 }}>
        {DOW.map((h, i) => <div key={'h' + i} style={{ textAlign: 'center', fontSize: 10, fontWeight: 700, color: '#888' }}>{h}</div>)}
        {cells.map(d => {
          const k = keyOf(d);
          if (d.getMonth() !== base.getMonth()) return <div key={k} style={{ minHeight: 40 }} />;   // blank pad for other months
          const hrs = dayHours[k] || 0;
          const tent = dayTent[k] || 0;
          const isPast = d < today;
          const isToday = k === todayKey;
          const holdOnly = hrs === 0 && tent > 0;   // only tentative time → "on hold" look
          let bg: string; let fg: string;
          if (isPast) { bg = '#f7f6f5'; fg = '#c0bdba'; }
          else if (hrs >= HOURS_PER_DAY) { bg = '#fde7e9'; fg = '#a4262c'; }
          else if (hrs > 0) { bg = '#fff7d6'; fg = '#8a6000'; }
          else if (holdOnly) { bg = '#fff7d6'; fg = '#8a6000'; }   // hatch drawn via backgroundImage below
          else { bg = '#eef9f1'; fg = '#107c10'; }
          const parts: string[] = [];
          if (hrs > 0) parts.push(`${hrs}h booked`);
          if (tent > 0) parts.push(`${tent}h on hold (pending approval)`);
          const reqId = dayReq[k];
          const clickable = !!onOpenRequest && !!reqId && !isPast;
          const title = (parts.length ? parts.join(' · ') + '\n' : '') + (dayLabels[k] ? dayLabels[k].join('\n') : (isPast ? '' : 'Free')) + (clickable ? '\n(click to open the engagement)' : '');
          return (
            <div key={k} title={title}
              onClick={clickable ? () => onOpenRequest!(reqId) : undefined}
              style={{ minHeight: 40, display: 'flex', flexDirection: 'column', padding: '3px 5px',
                background: bg, backgroundImage: (!isPast && holdOnly) ? HOLD_HATCH : undefined,
                borderRadius: 4, border: isToday ? `2px solid ${HPE_NAVY}` : (!isPast && holdOnly) ? '1px dashed #d0a000' : '1px solid #ececec',
                opacity: isPast ? 0.6 : 1, cursor: clickable ? 'pointer' : 'default' }}>
              <div style={{ fontSize: 11, fontWeight: isToday ? 800 : 600, color: fg }}>{d.getDate()}</div>
              {(hrs > 0 || tent > 0) && <div style={{ marginTop: 'auto', fontSize: 9, fontWeight: 700, color: fg }}>{hrs > 0 ? `${hrs}h` : `${tent}h hold`}</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
};

// Three months side by side (current → +2) — a ~90-day wall-calendar view across the full width.
export const AvailabilityCalendar: React.FC<{ commitments: ISseCommitment[]; onOpenRequest?: (requestId: number) => void }> = ({ commitments, onOpenRequest }) => {
  const dayHours: Record<string, number> = {};    // firm, both-approved bookings
  const dayTent: Record<string, number> = {};      // tentative "on hold" (pending approval)
  const dayLabels: Record<string, string[]> = {};
  const dayReq: Record<string, number> = {};       // a request to open when the day is clicked (first non-personal)
  for (const c of commitments) {
    if (!c.start) continue;
    const hpd = (typeof c.hoursPerDay === 'number') ? c.hoursPerDay : HOURS_PER_DAY;
    const s = new Date(c.start.substring(0, 10) + 'T00:00:00');
    const e = new Date((c.end || c.start).substring(0, 10) + 'T00:00:00');
    const span = Math.max(0, Math.round((e.getTime() - s.getTime()) / 86400000));
    for (let i = 0; i <= span; i++) {
      const d = new Date(s); d.setDate(s.getDate() + i);
      const k = keyOf(d);
      if (c.tentative) dayTent[k] = (dayTent[k] || 0) + hpd;
      else dayHours[k] = (dayHours[k] || 0) + hpd;
      // Personal time has no board row to open; link real engagements only (first one wins per day).
      if (!c.personal && c.requestId && !dayReq[k]) dayReq[k] = c.requestId;
      const who = c.personal ? '' : (c.customer ? ` · ${c.customer}` : '');
      const what = c.personal
        ? `🌴 ${c.label || 'Personal time'}`
        : `${c.tentative ? '⏳ On hold — ' : ''}${c.type === 'On-site' ? '📍 On-Site' : '💻 Remote/Prep'}${who}${c.location ? ' · ' + c.location : ''}`;
      (dayLabels[k] = dayLabels[k] || []).push(`${d.getMonth() + 1}/${d.getDate()} — ${what} (${hpd}h)`);
    }
  }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayKey = keyOf(today);
  const months = [0, 1, 2].map(off => new Date(today.getFullYear(), today.getMonth() + off, 1));

  const swatch = (bg: string, br: string): React.CSSProperties => ({ display: 'inline-block', width: 10, height: 10, background: bg, border: `1px solid ${br}`, borderRadius: 2, verticalAlign: 'middle', marginRight: 4 });

  return (
    <div>
      <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap', alignItems: 'flex-start' }}>
        {months.map((b, i) => <MonthGrid key={i} base={b} dayHours={dayHours} dayTent={dayTent} dayLabels={dayLabels} dayReq={dayReq} onOpenRequest={onOpenRequest} today={today} todayKey={todayKey} />)}
      </div>
      <div style={{ display: 'flex', gap: 14, marginTop: 10, fontSize: 10, color: '#605e5c', flexWrap: 'wrap' }}>
        <span><span style={swatch('#eef9f1', '#107c10')} />Free</span>
        <span><span style={swatch('#fff7d6', '#8a6000')} />Partly booked</span>
        <span><span style={swatch('#fde7e9', '#a4262c')} />Full day</span>
        <span><span style={{ ...swatch('#fff3d6', '#d0a000'), backgroundImage: HOLD_HATCH }} />On hold (pending approval)</span>
        <span style={{ color: '#888' }}>Hover a day for hours &amp; bookings.</span>
      </div>
    </div>
  );
};
