import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const SITE_URL  = 'https://hpe.sharepoint.com/teams/EnterpriseSalesTeamHome';
const LIST_NAME = 'Contacts';

// ── Territory matching: SRT BU/Region → Contact Directory Geography codes (SEGMENT|REGIONS) ──
export const SEGMENT_MAP: Record<string, string> = {
  'Enterprise-West': 'ENT', 'Enterprise-East': 'ENT',
  'Commercial-East': 'COMM', 'Commercial-Central': 'COMM', 'Commercial-West': 'COMM',
  'GMA': 'GMA',
};
export const REGION_CODE_MAP: Record<string, string> = {
  'Great Plains': 'GPLAINS', 'Midwest': 'MIDWEST', 'NorCal': 'NORCAL', 'NTOLA': 'NTOLA',
  'PacNW': 'PACNW', 'SoCal': 'SOCAL', 'Southwest': 'SWEST', 'STOLA': 'STOLA',
};
export const GENERALIST_CATEGORY = 'Strategic Systems Engineers';
export interface IStaffSlot { key: string; label: string; category: string; solutionCategories: string[] }
export const SPECIALIST_SLOTS: IStaffSlot[] = [
  { key: 'dc',      label: 'DC Specialist',      category: 'Datacenter Specialist Overlay',     solutionCategories: ['Data Center'] },
  { key: 'routing', label: 'Routing Specialist', category: 'Routing Specialist Overlay',        solutionCategories: ['Routing'] },
  { key: 'sase',    label: 'SASE Specialist',    category: 'SASE & Security Specialist Overlay', solutionCategories: ['SASE', 'NAC'] },
  { key: 'mist',    label: 'Mist Specialist',    category: 'Mist Specialist Overlay',            solutionCategories: ['Mist AI'] },
];

export function parseGeography(geo: string): { seg: string; regions: string[] }[] {
  return (geo || '').split(';').map(g => g.trim()).filter(Boolean).map(g => {
    const parts = g.split('|');
    return { seg: (parts[0] || '').trim().toUpperCase(), regions: (parts[1] || '').split(',').map(r => r.trim().toUpperCase()).filter(Boolean) };
  });
}
// 0 = out of territory · 1 = in the specific region · 2 = segment-wide (ALL)
export function territoryTier(geography: string, bu: string, region: string): number {
  const segCode = SEGMENT_MAP[bu] || '';
  const regCode = REGION_CODE_MAP[region] || '';
  if (!segCode) return 0;
  let tier = 0;
  for (const g of parseGeography(geography)) {
    if (g.seg !== segCode) continue;
    if (regCode && g.regions.indexOf(regCode) !== -1) return 1;
    if (g.regions.indexOf('ALL') !== -1) tier = 2;
  }
  return tier;
}

export interface IContact {
  id: number;
  name: string;
  email: string;
  category: string;
  businessUnit: string;
  geography: string;
  pocSolutions: string;
  phone: string;
}

export class ContactDirectoryService {
  constructor(private context: WebPartContext) {}

  async getAll(): Promise<IContact[]> {
    try {
      const sp = spfi(SITE_URL).using(SPFx(this.context));
      const items: { Id: number; Title: string; Email: string; Category: string; BusinessUnit: string; Geography: string; POCSolutions: string; Phone: string }[] =
        await sp.web.lists.getByTitle(LIST_NAME).items
          .select('Id', 'Title', 'Email', 'Category', 'BusinessUnit', 'Geography', 'POCSolutions', 'Phone')
          .orderBy('Title', true)
          .top(500)();
      return items.map(i => ({
        id: i.Id,
        name: i.Title || '',
        email: i.Email || '',
        category: i.Category || '',
        businessUnit: i.BusinessUnit || '',
        geography: i.Geography || '',
        pocSolutions: i.POCSolutions || '',
        phone: i.Phone || '',
      }));
    } catch {
      return [];
    }
  }

  suggestForSolutions(contacts: IContact[], solutionCodes: string[]): IContact[] {
    if (!solutionCodes.length) return [];
    return contacts.filter(c => {
      const cSols = c.pocSolutions.split(';').map(s => s.trim()).filter(Boolean);
      return solutionCodes.some(s => cSols.includes(s));
    });
  }
}
