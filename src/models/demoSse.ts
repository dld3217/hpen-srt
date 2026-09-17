import { ICseRequest } from './ICseRequest';
import { IScheduleBlock, newBlock, ScheduleBlockType } from './ScheduleBlock';

// A throwaway demo SSE with an artificially loaded calendar — for showing off the availability
// calendar + team heatmap. Every row is tagged [SAMPLE] so the dashboard 🧹 Clear [SAMPLE]
// button backs the whole thing out in one click.

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

function req(customer: string, blocks: IScheduleBlock[]): Omit<ICseRequest, 'id'> {
  return {
    title: `${customer} — SSE Support`,
    source: 'Demo',
    linkedPocId: 0,
    requestStatus: 'In Progress',
    scheduleStatus: 'TBD',
    requestedCse: JOE,
    sseManagerEmail: '',
    cseDescription: 'Demo engagement for Joe Cool.',
    csePriority: 'Medium',
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
    scheduleBlocks: blocks,
  };
}

// ~90 days of varied load: full-day on-site trips (red weeks), partial prep days (yellow), gaps (green).
export function demoJoeCoolRequests(): Omit<ICseRequest, 'id'>[] {
  return [
    req('[SAMPLE] Acme Corp',        [blk('Remote', 2, 3, 'Discovery & scoping'), blk('Prep', 6, 6, 'Lab build', '', 4), blk('On-Site', 9, 11, 'Kickoff visit', 'Acme HQ')]),
    req('[SAMPLE] Globex Fabric',    [blk('On-Site', 22, 24, 'DC fabric install', 'Globex DC'), blk('Remote', 26, 26, 'Follow-up', '', 2)]),
    req('[SAMPLE] Initech SD-WAN',   [blk('Prep', 33, 33, 'Design session', '', 4), blk('On-Site', 37, 40, 'Solution deep-dive', 'Initech')]),
    req('[SAMPLE] Umbrella Security',[blk('Remote', 51, 52, 'NAC review'), blk('On-Site', 55, 57, 'Rollout', 'Umbrella Corp')]),
    req('[SAMPLE] Stark Campus',     [blk('Prep', 68, 68, 'Exec deck prep', '', 6), blk('On-Site', 72, 73, 'Executive readout', 'Stark Tower')]),
  ];
}
