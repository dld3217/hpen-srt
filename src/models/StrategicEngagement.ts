// Strategic Engagement — shared types & taxonomy (Phase 1 / v1.0.69)

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

// Desired Outcome — multi-select goal(s) for the engagement (blended audiences can have several).
export const OUTCOME_OBJECTION = 'Overcome a specific objection';
export const OUTCOME_OTHER = 'Other';
export const DESIRED_OUTCOMES: string[] = [
  'Prove the proposed solution',
  'Advance to a purchase order',
  'Educate the executive team',
  "Shape the requirements / get spec'd in",
  'Displace the incumbent',
  'Establish trusted-advisor relationship',
  OUTCOME_OBJECTION,
  OUTCOME_OTHER,
];

export type EngagementOutcome = 'Advisory Only' | 'Spawned New POC' | 'Modified Existing POC';

// Disposition of each solution line relative to the customer's current landscape.
export type Disposition = 'Expansion' | 'New' | 'Integrate' | 'Displace';

// One row of the Solution Landscape:
//   solution (what WE position) + current vendor (what they run) + disposition.
export interface IEnvironmentRow {
  solutionCode: string; // gold-standard solution code (drives SolutionsFocus)
  solution: string;     // solution name for display / JSON readability
  vendor: string;       // current incumbent vendor (see below), or greenfield
  product: string;      // free text: their current model / product line
  version: string;      // free text: their current code / firmware version
  disposition: string;  // Disposition
  detail: string;       // competitor coexistence (Integrate) / migration (Displace) detail
}

// Vendor taxonomy — structured for bulletproof displacement reporting.
export const OUR_VENDORS: string[] = ['HPE Aruba', 'Juniper (HPE)'];

export const GREENFIELD_VENDOR = 'None / Greenfield';

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

export const ALL_VENDORS: string[] = [...OUR_VENDORS, GREENFIELD_VENDOR, ...COMPETITOR_VENDORS];

export function isOurVendor(vendor: string): boolean {
  return OUR_VENDORS.indexOf(vendor) !== -1;
}

export function isCompetitor(vendor: string): boolean {
  return !!vendor && vendor !== GREENFIELD_VENDOR && OUR_VENDORS.indexOf(vendor) === -1;
}

// Auto-disposition for non-competitor vendors; competitors are set by the SE (Integrate/Displace).
export function autoDisposition(vendor: string): string {
  if (!vendor) return '';
  if (vendor === GREENFIELD_VENDOR) return 'New';
  if (isOurVendor(vendor)) return 'Expansion';
  return ''; // competitor — SE must pick Integrate or Displace
}

export const DISPOSITION_STYLE: Record<string, { bg: string; color: string; label: string }> = {
  'Expansion': { bg: '#e8faf3', color: '#107c10', label: '✅ Expansion' },
  'New':       { bg: '#eff6fc', color: '#0078d4', label: '🌱 New' },
  'Integrate': { bg: '#fff4ce', color: '#8a6000', label: '🤝 Integrate' },
  'Displace':  { bg: '#fde7e9', color: '#a4262c', label: '🎯 Displace' },
};

// True if any row is a Displace (drives the HasDisplacement flag).
export function environmentHasDisplacement(rows: IEnvironmentRow[]): boolean {
  return (rows || []).some(r => r.disposition === 'Displace');
}
