// Strategic Engagement — shared types & taxonomy (Phase 1)

export type EngagementType = 'POC Support' | 'Strategic Engagement';

export type EngagementPurpose =
  | 'Strategic Briefing'
  | 'Roadmap / Product Direction'
  | 'EOL / Migration Planning'
  | 'Pre-Sales Advisory'
  | 'Architecture / Exec Conversation'
  | 'Other';

export const ENGAGEMENT_PURPOSES: EngagementPurpose[] = [
  'Strategic Briefing',
  'Roadmap / Product Direction',
  'EOL / Migration Planning',
  'Pre-Sales Advisory',
  'Architecture / Exec Conversation',
  'Other',
];

export type EngagementOutcome = 'Advisory Only' | 'Spawned New POC' | 'Modified Existing POC';

// One row of the customer's current environment / competitive landscape.
export interface IEnvironmentRow {
  area: string;     // gold-standard solution category (from SOLUTION_CATEGORIES)
  vendor: string;   // structured vendor (see below)
  product: string;  // free text: model / product line
  version: string;  // free text: code / firmware version
}

// Vendor taxonomy — structured for bulletproof displacement reporting.
export const OUR_VENDORS: string[] = ['HPE Aruba', 'Juniper (HPE)'];

export const COMPETITOR_VENDORS: string[] = [
  'Cisco',
  'Cisco Meraki',
  'Arista',
  'Extreme',
  'Fortinet',
  'Palo Alto',
  'VMware (VeloCloud)',
  'Versa',
  'Cato',
  'Forescout',
  'Dell',
  'Ubiquiti',
  'Ruckus / CommScope',
  'Nokia',
  'Huawei',
  'Other',
];

export const ALL_VENDORS: string[] = [...OUR_VENDORS, ...COMPETITOR_VENDORS];

// A row is a displacement opportunity when the incumbent vendor is NOT one of ours.
export function isDisplacement(vendor: string): boolean {
  return !!vendor && OUR_VENDORS.indexOf(vendor) === -1;
}

// True if any row in the environment is a competitor (drives the HasDisplacement flag).
export function environmentHasDisplacement(rows: IEnvironmentRow[]): boolean {
  return (rows || []).some(r => isDisplacement(r.vendor));
}
