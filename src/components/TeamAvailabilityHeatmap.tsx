import * as React from 'react';
import { ISseCommitment } from '../models/ICseRequest';
import { HOURS_PER_DAY } from '../models/ScheduleBlock';
import { HPE_NAVY } from '../styles/hpe';

// Weekly load → color (≈40h = a full week).
const weekColor = (h: number): { bg: string; fg: string } => {
  if (h <= 0)  return { bg: '#f0f7f2', fg: '#9bb0a2' };
  if (h < 12)  return { bg: '#dff1e4', fg: '#107c10' };
  if (h < 24)  return { bg: '#fbf1c7', fg: '#8a6000' };
  if (h < 36)  return { bg: '#ffdcae', fg: '#a4560c' };
  return { bg: '#fbcaca', fg: '#a4262c' };
};

// One row per SSE, one column per week (~90-day horizon). Each cell = hours booked that week,
// colored by load — at-a-glance "who's slammed, who's free" for the whole team.
export const TeamAvailabilityHeatmap: React.FC<{ commitments: ISseCommitment[]; roster: { name: string; email: string }[]; weeks?: number }> = ({ commitments, roster, weeks = 13 }) => {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const start = new Date(today); start.setDate(today.getDate() - today.getDay());   // Sunday of this week
  const startMs = start.getTime();
  const weekStarts: Date[] = [];
  for (let w = 0; w < weeks; w++) { const d = new Date(start); d.setDate(start.getDate() + w * 7); weekStarts.push(d); }

  const load: Record<string, number[]> = {};
  const nameByEmail: Record<string, string> = {};
  roster.forEach(r => { nameByEmail[r.email] = r.name; load[r.email] = new Array(weeks).fill(0); });
  for (const c of commitments) {
    if (!c.start) continue;
    const email = (c.sseEmail || '').toLowerCase();
    if (!email) continue;
    if (!nameByEmail[email]) nameByEmail[email] = c.sseName || email;
    if (!load[email]) load[email] = new Array(weeks).fill(0);
    const hpd = typeof c.hoursPerDay === 'number' ? c.hoursPerDay : HOURS_PER_DAY;
    const s = new Date(c.start.substring(0, 10) + 'T00:00:00');
    const e = new Date((c.end || c.start).substring(0, 10) + 'T00:00:00');
    const span = Math.max(0, Math.round((e.getTime() - s.getTime()) / 86400000));
    for (let i = 0; i <= span; i++) {
      const d = new Date(s); d.setDate(s.getDate() + i); d.setHours(0, 0, 0, 0);
      const wi = Math.floor(Math.round((d.getTime() - startMs) / 86400000) / 7);
      if (wi >= 0 && wi < weeks) load[email][wi] += hpd;
    }
  }

  const rows = Object.keys(load)
    .map(email => ({ email, name: nameByEmail[email] || email, hrs: load[email], total: load[email].reduce((a, b) => a + b, 0) }))
    .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));

  if (rows.length === 0) return <div style={{ fontSize: 12, color: '#888' }}>No SSEs to show.</div>;

  const TH: React.CSSProperties = { fontSize: 9, fontWeight: 700, color: '#888', padding: '2px 3px', textAlign: 'center', whiteSpace: 'nowrap' };
  const NAME: React.CSSProperties = { fontSize: 11, fontWeight: 600, color: HPE_NAVY, padding: '3px 8px 3px 2px', whiteSpace: 'nowrap', position: 'sticky', left: 0, background: '#fff' };

  return (
    <div style={{ overflowX: 'auto' }}>
      <table style={{ borderCollapse: 'separate', borderSpacing: 2 }}>
        <thead>
          <tr>
            <th style={{ ...TH, textAlign: 'left', position: 'sticky', left: 0, background: '#fff' }}>SSE</th>
            {weekStarts.map((d, i) => <th key={i} style={TH}>{d.getMonth() + 1}/{d.getDate()}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={r.email}>
              <td style={NAME}>{r.name}</td>
              {r.hrs.map((h, i) => {
                const c = weekColor(h);
                return (
                  <td key={i} style={{ fontSize: 10, textAlign: 'center', padding: '3px 2px', borderRadius: 2, minWidth: 30, background: c.bg, color: c.fg }}
                    title={`${r.name} · week of ${weekStarts[i].getMonth() + 1}/${weekStarts[i].getDate()}: ${h}h booked`}>
                    {h || ''}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ marginTop: 8, fontSize: 10, color: '#605e5c' }}>
        Hours booked per week (≈40h = a full week). Green = open · yellow = filling · red = slammed. Hover a cell for detail.
      </div>
    </div>
  );
};
