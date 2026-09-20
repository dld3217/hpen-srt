import { ICseRequest, CseRequestStatus } from './ICseRequest';
import { IScheduleBlock, newBlock, ScheduleBlockType } from './ScheduleBlock';

// A throwaway demo SSE with an artificially loaded ~3-month calendar — for walking through the
// availability calendar + team heatmap. It intentionally covers EVERY availability state so the
// whole picture demos in one shot: firm booked (red), partial-hours (yellow), on-hold/tentative
// (amber-hatch, Pending/Accepted rows), and personal time (🌴). Every row is tagged [SAMPLE] so the
// dashboard 🧹 Clear [SAMPLE] button backs the whole thing out in one click.

const JOE = 'Joe Cool / joe.cool@hpe.com';

const p2 = (n: number): string => (n < 10 ? '0' + n : '' + n);
const dayOff = (off: number): string => { const t = new Date(); t.setDate(t.getDate() + off); return `${t.getFullYear()}-${p2(t.getMonth() + 1)}-${p2(t.getDate())}`; };

// A schedule block by day offsets from today. Pass `hours` for an hours-mode (partial-day) block.
function blk(type: ScheduleBlockType, startOff: number, endOff: number, label: string, location = '', hours?: number): IScheduleBlock {
  const b = newBlock(type);
  b.tbd = false;
  b.start = dayOff(startOff);
  if (hours !== undefined) { b.unit = 'hours'; b.hours = hours; b.end = dayOff(startOff); }
  else { b.end = dayOff(endOff); }
  b.label = label;
  b.location = type === 'On-Site' ? location : '';
  return b;
}

interface IDemoOpts { status?: CseRequestStatus; engagementType?: string; priority?: string; }

function req(customer: string, blocks: IScheduleBlock[], opts: IDemoOpts = {}): Omit<ICseRequest, 'id'> {
  return {
    title: `${customer} — SSE Support`,
    source: opts.engagementType === 'Personal' ? 'Personal Time' : 'Demo',
    linkedPocId: 0,
    // In Progress = FIRM (solid on the calendar). Pending/Accepted = on-hold (amber, not both-approved).
    requestStatus: opts.status || 'In Progress',
    scheduleStatus: 'TBD',
    requestedCse: JOE,
    sseManagerEmail: '',
    cseDescription: 'Demo engagement for Joe Cool.',
    csePriority: opts.priority || 'Medium',
    csePriorityReason: '',
    solutionsFocus: '',
    supportType: '',
    remoteTbd: true, remoteStart: '', remoteEnd: '', remoteDuration: '',
    onsiteTbd: true, onsiteStart: '', onsiteEnd: '', onsiteDuration: '', onsiteDestination: '',
    sePrimary: '',
    semPrimary: '',
    sedEmail: '',
    buRegion: '',
    hpenBusinessUnit: '',
    customerName: customer,
    pocName: '',
    opportunityAmount: 0,
    custTemp: 'Normal',
    signedOffBy: '',
    signOffDate: '',
    notes: '',
    engagementType: opts.engagementType || '',
    scheduleBlocks: blocks,
  };
}

// Personal (non-opportunity) time — PTO / holiday / appointment. Hidden from the board, shows 🌴 on
// the calendar. Pass `hours` for a partial-day appointment (yellow); omit for full day(s) (red).
function personal(label: string, startOff: number, endOff: number, hours?: number): Omit<ICseRequest, 'id'> {
  return req(`[SAMPLE] ${label}`, [blk('Prep', startOff, endOff, label, '', hours)], { engagementType: 'Personal', priority: 'Low' });
}

// ~3 months of varied load spanning the whole availability calendar, mixing every state so the
// picture is walkable end to end: firm on-site trips (red), partial prep/remote (yellow), on-hold
// proposals (amber), personal time (🌴), and green gaps in between.
export function demoJoeCoolRequests(): Omit<ICseRequest, 'id'>[] {
  return [
    // ── Month 1 ──
    req('[SAMPLE] Acme Corp',         [blk('Remote', 2, 3, 'Discovery & scoping'), blk('Prep', 5, 5, 'Lab build', '', 4), blk('On-Site', 8, 10, 'Kickoff visit', 'Acme HQ')]),
    req('[SAMPLE] Wayne Enterprises', [blk('On-Site', 15, 17, 'Proposed design workshop', 'Wayne HQ')], { status: 'Pending' }),   // on-hold (amber)
    personal('PTO — Vacation', 22, 24),                                                                                          // 🌴 full days
    // ── Month 2 ──
    req('[SAMPLE] Globex Fabric',     [blk('On-Site', 33, 35, 'DC fabric install', 'Globex DC'), blk('Remote', 38, 38, 'Follow-up', '', 2)]),
    personal('Doctor appointment', 43, 43, 2),                                                                                    // 🌴 partial (yellow)
    req('[SAMPLE] Oscorp NAC',        [blk('On-Site', 47, 49, 'Proposed rollout', 'Oscorp')], { status: 'Accepted' }),           // on-hold (amber)
    // ── Month 3 ──
    req('[SAMPLE] Initech SD-WAN',    [blk('Prep', 63, 63, 'Design session', '', 4), blk('On-Site', 66, 68, 'Solution deep-dive', 'Initech')]),
    req('[SAMPLE] Stark Campus',      [blk('On-Site', 75, 77, 'Proposed exec readout', 'Stark Tower')], { status: 'Pending' }),  // on-hold (amber)
    personal('Holiday', 81, 81),                                                                                                  // 🌴 full day
  ];
}
