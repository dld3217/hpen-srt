export type CseRequestStatus = 'Pending' | 'Accepted' | 'Scheduled' | 'In Progress' | 'Parked' | 'Complete' | 'Declined' | 'Needs Info' | 'Cancelled';
export type ScheduleStatus = 'TBD' | 'Dates Proposed' | 'Dates Confirmed' | 'Rescheduling';
export type CustTemp = 'Low' | 'Normal' | 'High' | 'Critical';

export interface ICseRequest {
  id?: number;
  title: string;
  source: string;
  linkedPocId: number;
  requestStatus: CseRequestStatus;
  scheduleStatus: ScheduleStatus;
  requestedCse: string;
  sseManagerEmail: string;
  cseDescription: string;
  csePriority: string;
  csePriorityReason: string;
  solutionsFocus: string;
  supportType: string;
  remoteTbd: boolean;
  remoteStart: string;
  remoteEnd: string;
  remoteDuration: string;
  onsiteTbd: boolean;
  onsiteStart: string;
  onsiteEnd: string;
  onsiteDuration: string;
  onsiteDestination: string;
  sePrimary: string;
  semPrimary: string;
  sedEmail: string;
  buRegion: string;
  hpenBusinessUnit: string;
  customerName: string;
  pocName: string;
  opportunityAmount: number;
  custTemp: CustTemp;
  signedOffBy: string;
  signOffDate: string;
  opportunity?: string;
  notes: string;
  modified?: string;
  specialtyType?: string;
  // Strategic Engagement (Phase 1)
  engagementType?: string;        // 'POC Support' | 'Strategic Engagement'
  engagementPurpose?: string;     // see ENGAGEMENT_PURPOSES
  currentEnvironment?: string;    // JSON string of IEnvironmentRow[]
  hasDisplacement?: boolean;      // any competitor row present
  engagementOutcome?: string;     // ACTUAL result: 'Advisory Only' | 'Spawned New POC' | 'Modified Existing POC'
  desiredOutcome?: string[];      // GOAL(s), multi-choice selected outcomes (set at request time)
  desiredOutcomeDetail?: string;  // objection list + any 'Other' free text
  engagementPurposeOther?: string; // free text when EngagementPurpose = Other
}

export const CSE_STATUS_STYLE: Record<CseRequestStatus, { bg: string; color: string }> = {
  'Pending':     { bg: '#edebe9', color: '#605e5c' },
  'Accepted':    { bg: '#eff6fc', color: '#0078d4' },
  'Scheduled':   { bg: '#e0f4f4', color: '#007a7a' },
  'In Progress': { bg: '#ece1f7', color: '#4b2e83' },
  'Parked':      { bg: '#e6e2f0', color: '#5b4b8a' },
  'Complete':    { bg: '#e8faf3', color: '#107c10' },
  'Declined':    { bg: '#fde7e9', color: '#a4262c' },
  'Needs Info':  { bg: '#f0e6ff', color: '#6b2faf' },
  'Cancelled':   { bg: '#f0f0f0', color: '#a19f9d' },
};

export const CUST_TEMP_STYLE: Record<CustTemp, { bg: string; color: string; label: string }> = {
  'Low':      { bg: '#f0f0f0', color: '#605e5c', label: 'Low' },
  'Normal':   { bg: '#dceee1', color: '#1e6b3a', label: 'Normal' },
  'High':     { bg: '#fff4ce', color: '#8a6000', label: 'High' },
  'Critical': { bg: '#fde7e9', color: '#a4262c', label: 'Critical' },
};

export const SCHEDULE_STATUS_STYLE: Record<ScheduleStatus, { bg: string; color: string }> = {
  'TBD':             { bg: '#edebe9', color: '#605e5c' },
  'Dates Proposed':  { bg: '#fff4ce', color: '#8a6000' },
  'Dates Confirmed': { bg: '#e8faf3', color: '#107c10' },
  'Rescheduling':    { bg: '#fde7e9', color: '#a4262c' },
};
