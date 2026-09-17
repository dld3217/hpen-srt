// Flexible per-engagement scheduling. An engagement carries ANY number of independently-dated,
// typed time blocks (Remote / Prep / On-Site) — contiguous or not. Blocks power both PLANNING
// (future intent) and ACCOUNTING (flip Planned→Actual, optional hours override) with minimal input.

export type ScheduleBlockType = 'Remote' | 'Prep' | 'On-Site';

export const SCHEDULE_BLOCK_TYPES: ScheduleBlockType[] = ['Remote', 'Prep', 'On-Site'];

export const BLOCK_STYLE: Record<ScheduleBlockType, { bg: string; color: string; icon: string }> = {
  'Remote':  { bg: '#eff6fc', color: '#0078d4', icon: '💻' },
  'Prep':    { bg: '#f3f0fa', color: '#5b4b8a', icon: '📋' },
  'On-Site': { bg: '#e0f4f4', color: '#007a7a', icon: '📍' },
};

export interface IScheduleBlock {
  id: string;                    // stable id for edit / confirm / log
  type: ScheduleBlockType;
  start: string;                 // 'YYYY-MM-DD' ('' when tbd)
  end: string;                   // 'YYYY-MM-DD' ('' when tbd)
  location: string;              // On-Site destination; '' for Remote/Prep (home office)
  label: string;                 // optional note, e.g. 'Kickoff visit'
  tbd: boolean;                  // dates not yet set
  // Confirmation — customer-facing blocks only; Prep is self-booked and skips the handshake.
  proposedBy: '' | 'SE' | 'SSE';
  confirmed: boolean;
  // Accounting — Planned → Actual via one switch; optional hours override.
  logged: boolean;
  actualHours?: number;          // when logged & unset, actual == planned
}

export const HOURS_PER_DAY = 8;

// Prep is the SSE's own time — self-booked, never needs the SE/SSE confirmation handshake.
export const isSelfBooked = (t: ScheduleBlockType): boolean => t === 'Prep';

// Inclusive calendar-day span of a block's date range (0 when tbd / undated).
export function blockDays(b: IScheduleBlock): number {
  if (b.tbd || !b.start) return 0;
  const s = b.start.substring(0, 10);
  const e = (b.end || b.start).substring(0, 10);
  const [sy, sm, sd] = s.split('-').map(Number);
  const [ey, em, ed] = e.split('-').map(Number);
  const ms = new Date(ey, em - 1, ed).getTime() - new Date(sy, sm - 1, sd).getTime();
  return Math.max(0, Math.round(ms / 86400000)) + 1;
}

// Planned effort in hours (days × HOURS_PER_DAY).
export function plannedHours(b: IScheduleBlock): number {
  return blockDays(b) * HOURS_PER_DAY;
}

// Actual effort: explicit override if given, else the planned figure once logged, else 0.
export function actualHoursOf(b: IScheduleBlock): number {
  if (!b.logged) return 0;
  return (typeof b.actualHours === 'number') ? b.actualHours : plannedHours(b);
}

// A block occupies the SSE's calendar (for double-book avoidance) once it has real dates.
export function isBusyBlock(b: IScheduleBlock): boolean {
  return !b.tbd && !!b.start;
}

// Parse the JSON stored in the ScheduleBlocks column into a normalized array (tolerant of junk).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function parseScheduleBlocks(json: string | undefined): IScheduleBlock[] {
  let arr: any[]; // eslint-disable-line @typescript-eslint/no-explicit-any
  try { arr = JSON.parse(json || '[]'); } catch { return []; }
  if (!Array.isArray(arr)) return [];
  return arr.map(b => ({
    id: String(b.id || ''),
    type: (b.type === 'Prep' || b.type === 'On-Site') ? b.type : 'Remote',
    start: b.start || '',
    end: b.end || '',
    location: b.location || '',
    label: b.label || '',
    tbd: !!b.tbd,
    proposedBy: (b.proposedBy === 'SE' || b.proposedBy === 'SSE') ? b.proposedBy : '',
    confirmed: !!b.confirmed,
    logged: !!b.logged,
    actualHours: (typeof b.actualHours === 'number') ? b.actualHours : undefined,
  }));
}

let _seq = 0;
export function newBlock(type: ScheduleBlockType): IScheduleBlock {
  _seq += 1;
  return {
    id: 'b' + new Date().getTime().toString(36) + '-' + _seq.toString(36),
    type, start: '', end: '', location: '', label: '', tbd: true,
    proposedBy: '', confirmed: false, logged: false, actualHours: undefined,
  };
}
