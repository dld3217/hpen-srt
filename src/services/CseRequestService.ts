import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ICseRequest, CseRequestStatus, ScheduleStatus, CustTemp, ISseCommitment } from '../models/ICseRequest';
import { IScheduleBlock, parseScheduleBlocks, HOURS_PER_DAY } from '../models/ScheduleBlock';

const LIST_NAME = 'CSERequests';

const SP_SELECT = [
  'Id', 'Title', 'Source', 'LinkedPOCId', 'RequestStatus', 'ScheduleStatus',
  'RequestedCSE', 'SSEManagerEmail', 'CSEDescription', 'CSEPriority', 'CSEPriorityReason', 'SolutionsFocus',
  'SupportType', 'RemoteTBD', 'RemoteStart', 'RemoteEnd', 'RemoteDuration',
  'OnsiteTBD', 'OnsiteStart', 'OnsiteEnd', 'OnsiteDuration', 'OnsiteDestination',
  'SEPrimary', 'SEMPrimary', 'SEDEmail', 'BURegion', 'HPENBusinessUnit',
  'CustomerName', 'POCName', 'OpportunityAmount',
  'CustTemp', 'SignedOffBy', 'SignOffDate', 'Opportunity', 'Notes', 'AdditionalResources', 'Modified', 'SpecialtyType',
  'EngagementType', 'EngagementPurpose', 'CurrentEnvironment', 'HasDisplacement', 'EngagementOutcome',
  'DesiredOutcome', 'DesiredOutcomeDetail', 'EngagementPurposeOther', 'DatesProposedBy',
  'SpecialProjectCategory', 'SpecialProjectInitiative', 'ScheduleBlocks',
];

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function unwrapChoice(val: unknown): string {
  if (!val) return '';
  if (typeof val === 'object' && val !== null && 'Value' in val) return (val as Record<string, string>).Value || '';
  if (typeof val === 'string' && val.startsWith('{')) {
    try { const p = JSON.parse(val); return p.Value || val; } catch { /* not JSON */ }
  }
  return String(val);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function mapToRequest(item: Record<string, any>): ICseRequest {
  return {
    id: item.Id,
    title: item.Title || '',
    source: item.Source || '',
    linkedPocId: item.LinkedPOCId || 0,
    requestStatus: (unwrapChoice(item.RequestStatus) || 'Pending') as CseRequestStatus,
    scheduleStatus: (unwrapChoice(item.ScheduleStatus) || 'TBD') as ScheduleStatus,
    requestedCse: item.RequestedCSE || '',
    sseManagerEmail: item.SSEManagerEmail || '',
    cseDescription: item.CSEDescription || '',
    csePriority: unwrapChoice(item.CSEPriority),
    csePriorityReason: item.CSEPriorityReason || '',
    solutionsFocus: unwrapChoice(item.SolutionsFocus),
    supportType: item.SupportType || '',
    remoteTbd: !!item.RemoteTBD,
    remoteStart: item.RemoteStart || '',
    remoteEnd: item.RemoteEnd || '',
    remoteDuration: item.RemoteDuration || '',
    onsiteTbd: !!item.OnsiteTBD,
    onsiteStart: item.OnsiteStart || '',
    onsiteEnd: item.OnsiteEnd || '',
    onsiteDuration: item.OnsiteDuration || '',
    onsiteDestination: item.OnsiteDestination || '',
    sePrimary: item.SEPrimary || '',
    semPrimary: item.SEMPrimary || '',
    sedEmail: item.SEDEmail || '',
    buRegion: unwrapChoice(item.BURegion),
    hpenBusinessUnit: unwrapChoice(item.HPENBusinessUnit),
    customerName: item.CustomerName || '',
    pocName: item.POCName || '',
    opportunityAmount: item.OpportunityAmount || 0,
    custTemp: (unwrapChoice(item.CustTemp) || 'Normal') as CustTemp,
    signedOffBy: item.SignedOffBy || '',
    signOffDate: item.SignOffDate || '',
    opportunity: item.Opportunity || '',
    notes: item.Notes || '',
    additionalResources: item.AdditionalResources || '',
    modified: item.Modified || '',
    specialtyType: item.SpecialtyType || '',
    engagementType: unwrapChoice(item.EngagementType),
    engagementPurpose: unwrapChoice(item.EngagementPurpose),
    currentEnvironment: item.CurrentEnvironment || '',
    hasDisplacement: !!item.HasDisplacement,
    engagementOutcome: unwrapChoice(item.EngagementOutcome),
    desiredOutcome: Array.isArray(item.DesiredOutcome) ? item.DesiredOutcome : (item.DesiredOutcome?.results || []),
    desiredOutcomeDetail: item.DesiredOutcomeDetail || '',
    engagementPurposeOther: item.EngagementPurposeOther || '',
    datesProposedBy: item.DatesProposedBy || '',
    specialProjectCategory: item.SpecialProjectCategory || '',
    specialProjectInitiative: item.SpecialProjectInitiative || '',
    scheduleBlocks: parseScheduleBlocks(item.ScheduleBlocks),
  };
}

export class CseRequestService {
  constructor(private sp: SPFI) {}

  async create(req: Omit<ICseRequest, 'id'>): Promise<number> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result: any = await this.sp.web.lists.getByTitle(LIST_NAME).items.add({
      Title: req.title,
      Source: req.source,
      LinkedPOCId: req.linkedPocId || 0,
      RequestStatus: req.requestStatus,
      ScheduleStatus: req.scheduleStatus || 'TBD',
      DatesProposedBy: req.datesProposedBy || '',
      RequestedCSE: req.requestedCse,
      SSEManagerEmail: req.sseManagerEmail || '',
      CSEDescription: req.cseDescription,
      CSEPriority: req.csePriority,
      CSEPriorityReason: req.csePriorityReason,
      SolutionsFocus: req.solutionsFocus,
      SupportType: req.supportType,
      RemoteTBD: req.remoteTbd,
      RemoteStart: req.remoteStart || null,
      RemoteEnd: req.remoteEnd || null,
      RemoteDuration: req.remoteDuration,
      OnsiteTBD: req.onsiteTbd,
      OnsiteStart: req.onsiteStart || null,
      OnsiteEnd: req.onsiteEnd || null,
      OnsiteDuration: req.onsiteDuration,
      OnsiteDestination: req.onsiteDestination,
      SEPrimary: req.sePrimary,
      SEMPrimary: req.semPrimary,
      SEDEmail: req.sedEmail,
      BURegion: req.buRegion,
      HPENBusinessUnit: req.hpenBusinessUnit,
      CustomerName: req.customerName,
      POCName: req.pocName,
      OpportunityAmount: req.opportunityAmount || 0,
      CustTemp: req.custTemp || 'Normal',
      Notes: req.notes,
      AdditionalResources: req.additionalResources || '',
      SpecialtyType: req.specialtyType || '',
      EngagementType: req.engagementType || '',
      EngagementPurpose: req.engagementPurpose || '',
      CurrentEnvironment: req.currentEnvironment || '',
      HasDisplacement: !!req.hasDisplacement,
      EngagementOutcome: req.engagementOutcome || '',
      DesiredOutcome: req.desiredOutcome || [],
      DesiredOutcomeDetail: req.desiredOutcomeDetail || '',
      EngagementPurposeOther: req.engagementPurposeOther || '',
      SpecialProjectCategory: req.specialProjectCategory || '',
      SpecialProjectInitiative: req.specialProjectInitiative || '',
      ScheduleBlocks: JSON.stringify(req.scheduleBlocks || []),
    });
    return result.Id || result.id || 0;
  }

  async reassign(id: number, requestedCse: string, sseManagerEmail: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      RequestedCSE: requestedCse,
      SSEManagerEmail: sseManagerEmail,
      RequestStatus: 'Accepted',
    });
  }

  async getByPocId(pocId: number): Promise<ICseRequest[]> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = await this.sp.web.lists.getByTitle(LIST_NAME).items
      .filter(`LinkedPOCId eq ${pocId}`)
      .select(...SP_SELECT)
      .orderBy('Created', true)();
    return items.map(mapToRequest);
  }

  async getById(id: number): Promise<ICseRequest> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item: any = await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).select(...SP_SELECT)();
    return mapToRequest(item);
  }

  async getStatus(id: number): Promise<CseRequestStatus> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item: any = await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).select('RequestStatus')();
    return (item.RequestStatus || 'Pending') as CseRequestStatus;
  }

  async getAll(): Promise<ICseRequest[]> {
    // Page fully so CSERequests over the page size neither truncates the Dashboard/Insights
    // counts nor throws the view threshold. Requires Created (sort) indexed on the list.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = [];
    const q = this.sp.web.lists.getByTitle(LIST_NAME).items
      .select(...SP_SELECT)
      .orderBy('Created', false)
      .top(2000);
    for await (const page of q) items.push(...page);
    return items.map(mapToRequest);
  }

  async updateStatus(id: number, status: CseRequestStatus, notes?: string): Promise<void> {
    const update: Record<string, unknown> = { RequestStatus: status };
    if (notes !== undefined) update.Notes = notes;
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update(update);
  }

  // Park an engagement: put it on hold AND free up the calendar. Clears the proposed/confirmed
  // dates back to TBD so the SSE's availability opens immediately (Parked is also excluded from
  // getSseCommitments). Schedule blocks are left intact — re-date on the dashboard when it un-parks.
  async parkRequest(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      RequestStatus: 'Parked',
      ScheduleStatus: 'TBD',
      DatesProposedBy: '',
      RemoteTBD: true,  RemoteStart: null, RemoteEnd: null, RemoteDuration: '',
      OnsiteTBD: true,  OnsiteStart: null, OnsiteEnd: null, OnsiteDuration: '', OnsiteDestination: '',
    });
  }

  async updateDates(id: number, dates: {
    remoteTbd?: boolean; remoteStart?: string; remoteEnd?: string; remoteDuration?: string;
    onsiteTbd?: boolean; onsiteStart?: string; onsiteEnd?: string; onsiteDuration?: string; onsiteDestination?: string;
    scheduleStatus?: ScheduleStatus; datesProposedBy?: string;
  }): Promise<void> {
    const update: Record<string, unknown> = {
      ScheduleStatus: dates.scheduleStatus || 'Dates Proposed',
    };
    if (dates.datesProposedBy !== undefined) update.DatesProposedBy = dates.datesProposedBy;
    if (dates.remoteTbd !== undefined) update.RemoteTBD = dates.remoteTbd;
    if (dates.remoteStart !== undefined) update.RemoteStart = dates.remoteStart || null;
    if (dates.remoteEnd !== undefined) update.RemoteEnd = dates.remoteEnd || null;
    if (dates.remoteDuration !== undefined) update.RemoteDuration = dates.remoteDuration;
    if (dates.onsiteTbd !== undefined) update.OnsiteTBD = dates.onsiteTbd;
    if (dates.onsiteStart !== undefined) update.OnsiteStart = dates.onsiteStart || null;
    if (dates.onsiteEnd !== undefined) update.OnsiteEnd = dates.onsiteEnd || null;
    if (dates.onsiteDuration !== undefined) update.OnsiteDuration = dates.onsiteDuration;
    if (dates.onsiteDestination !== undefined) update.OnsiteDestination = dates.onsiteDestination;
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update(update);
  }

  async confirmDates(id: number, requestStatus: CseRequestStatus = 'Scheduled'): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      ScheduleStatus: 'Dates Confirmed',
      RequestStatus: requestStatus,
      DatesProposedBy: '', // handshake complete — clear the ball-in-court marker
    });
  }

  async declineDates(id: number, notes?: string): Promise<void> {
    const update: Record<string, unknown> = { ScheduleStatus: 'Rescheduling' };
    if (notes) update.Notes = notes;
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update(update);
  }

  async cancelRequest(id: number, reason: string, note?: string): Promise<void> {
    const notesValue = [reason, note].filter(Boolean).join(' — ');
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      RequestStatus: 'Cancelled',
      Notes: notesValue,
    });
  }

  async updateCustTemp(id: number, custTemp: CustTemp): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({ CustTemp: custTemp });
  }

  async updatePriority(id: number, priority: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({ CSEPriority: priority });
  }

  async signOff(id: number, signedOffBy: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      SignedOffBy: signedOffBy,
      SignOffDate: new Date().toISOString(),
      RequestStatus: 'Complete',
    });
  }

  async delete(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).delete();
  }

  async updateSolutions(id: number, solutionsFocus: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({ SolutionsFocus: solutionsFocus });
  }

  // Persist the flexible schedule-block list for a request.
  async updateScheduleBlocks(id: number, blocks: IScheduleBlock[]): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({ ScheduleBlocks: JSON.stringify(blocks || []) });
  }

  // Set / change a Special Project's category or initiative on an existing row.
  async updateSpecialProject(id: number, fields: { category?: string; initiative?: string }): Promise<void> {
    const update: Record<string, unknown> = {};
    if (fields.category !== undefined) update.SpecialProjectCategory = fields.category;
    if (fields.initiative !== undefined) update.SpecialProjectInitiative = fields.initiative;
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update(update);
  }

  async updateOpportunityAndNotes(id: number, opportunity: string, notes: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      Opportunity: opportunity,
      Notes: notes,
    });
  }

  // Confirmed, upcoming on-site + remote busy blocks — for SSE availability / double-book avoidance.
  // Pass an sseEmail to scope to one SSE; omit for all SSEs (standalone availability view).
  async getSseCommitments(sseEmail?: string): Promise<ISseCommitment[]> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: Record<string, any>[] = [];
    const q = this.sp.web.lists.getByTitle(LIST_NAME).items.select(...SP_SELECT).top(2000);
    for await (const page of q) items.push(...page);
    const reqs = items.map(mapToRequest);
    const now = new Date();
    const pad = (n: number): string => (n < 10 ? '0' + n : '' + n);
    const todayStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const email = (sseEmail || '').toLowerCase();
    const out: ISseCommitment[] = [];
    for (const r of reqs) {
      // Done or freed → off the calendar entirely. Parked deliberately releases its dates.
      if (['Declined', 'Cancelled', 'Complete', 'Parked'].indexOf(r.requestStatus) !== -1) continue;
      if (email && !(r.requestedCse || '').toLowerCase().includes(email)) continue;
      // FIRM = both parties approved (Scheduled / In Progress). Otherwise the dates are still
      // TENTATIVE ("on hold") — surfaced so another SE won't book over a pending request, but
      // rendered distinctly and never counted as a hard commitment. Personal time is always firm.
      const tentative = r.requestStatus !== 'Scheduled' && r.requestStatus !== 'In Progress';
      const sseName = (r.requestedCse.split('/')[0] || '').trim() || r.requestedCse;
      const sseEmailVal = (r.requestedCse.split('/')[1] || '').trim();
      const isPersonal = r.engagementType === 'Personal';
      // New model: every dated schedule block (Remote/Prep/On-Site) is busy time.
      const busyBlocks = (r.scheduleBlocks || []).filter(b => !b.tbd && b.start);
      if (busyBlocks.length) {
        for (const b of busyBlocks) {
          const end = b.end || b.start;
          const hpd = b.unit === 'hours' ? (b.hours || 0) : HOURS_PER_DAY;
          if (end.substring(0, 10) >= todayStr) {
            out.push({ start: b.start, end, type: b.type === 'On-Site' ? 'On-site' : 'Remote', location: b.type === 'On-Site' ? (b.location || '') : '', sseEmail: sseEmailVal, sseName, requestId: r.id, hoursPerDay: hpd, label: b.label || '', personal: isPersonal, tentative });
          }
        }
        continue; // blocks supersede the legacy scalar fields for this request
      }
      // Legacy model: firm rows require confirmed dates; tentative rows surface any set dates as "on hold".
      if (!tentative && r.scheduleStatus !== 'Dates Confirmed') continue;
      if (!r.remoteTbd && r.remoteStart) {
        const end = r.remoteEnd || r.remoteStart;
        if (end.substring(0, 10) >= todayStr) out.push({ start: r.remoteStart, end, type: 'Remote', location: '', sseEmail: sseEmailVal, sseName, requestId: r.id, hoursPerDay: HOURS_PER_DAY, personal: isPersonal, tentative });
      }
      if (!r.onsiteTbd && r.onsiteStart) {
        const end = r.onsiteEnd || r.onsiteStart;
        if (end.substring(0, 10) >= todayStr) out.push({ start: r.onsiteStart, end, type: 'On-site', location: r.onsiteDestination || '', sseEmail: sseEmailVal, sseName, requestId: r.id, hoursPerDay: HOURS_PER_DAY, personal: isPersonal, tentative });
      }
    }
    out.sort((a, b) => (a.start.substring(0, 10) < b.start.substring(0, 10) ? -1 : 1));
    return out;
  }
}
