import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ICseRequest, CseRequestStatus, ScheduleStatus, CustTemp } from '../models/ICseRequest';

const LIST_NAME = 'CSERequests';

const SP_SELECT = [
  'Id', 'Title', 'Source', 'LinkedPOCId', 'RequestStatus', 'ScheduleStatus',
  'RequestedCSE', 'SSEManagerEmail', 'CSEDescription', 'CSEPriority', 'CSEPriorityReason', 'SolutionsFocus',
  'SupportType', 'RemoteTBD', 'RemoteStart', 'RemoteEnd', 'RemoteDuration',
  'OnsiteTBD', 'OnsiteStart', 'OnsiteEnd', 'OnsiteDuration', 'OnsiteDestination',
  'SEPrimary', 'SEMPrimary', 'SEDEmail', 'BURegion', 'HPENBusinessUnit',
  'CustomerName', 'POCName', 'OpportunityAmount',
  'CustTemp', 'SignedOffBy', 'SignOffDate', 'Notes',
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
    requestStatus: (item.RequestStatus || 'Pending') as CseRequestStatus,
    scheduleStatus: (item.ScheduleStatus || 'TBD') as ScheduleStatus,
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
    custTemp: (item.CustTemp || 'Normal') as CustTemp,
    signedOffBy: item.SignedOffBy || '',
    signOffDate: item.SignOffDate || '',
    notes: item.Notes || '',
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
    });
    return result.Id || result.id || 0;
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
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const items: any[] = await this.sp.web.lists.getByTitle(LIST_NAME).items
      .select(...SP_SELECT)
      .orderBy('Created', false)
      .top(2000)();
    return items.map(mapToRequest);
  }

  async updateStatus(id: number, status: CseRequestStatus, notes?: string): Promise<void> {
    const update: Record<string, unknown> = { RequestStatus: status };
    if (notes !== undefined) update.Notes = notes;
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update(update);
  }

  async updateDates(id: number, dates: {
    remoteTbd?: boolean; remoteStart?: string; remoteEnd?: string; remoteDuration?: string;
    onsiteTbd?: boolean; onsiteStart?: string; onsiteEnd?: string; onsiteDuration?: string; onsiteDestination?: string;
    scheduleStatus?: ScheduleStatus;
  }): Promise<void> {
    const update: Record<string, unknown> = {
      ScheduleStatus: dates.scheduleStatus || 'Dates Proposed',
    };
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

  async confirmDates(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      ScheduleStatus: 'Dates Confirmed',
      RequestStatus: 'Scheduled',
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

  async signOff(id: number, signedOffBy: string): Promise<void> {
    await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(id).update({
      SignedOffBy: signedOffBy,
      SignOffDate: new Date().toISOString(),
      RequestStatus: 'Complete',
    });
  }
}
