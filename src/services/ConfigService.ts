import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { ISolutionDef, SOLUTIONS } from '../models/ISolution';

export interface IRegionConfig {
  semEmail?: string;
  sdEmail?: string;
  managerEmail?: string;
}

export interface IBUConfig {
  sedEmail?: string;
  vpGmEmail?: string;
  regions: Record<string, IRegionConfig>;
}

export type BURegionMap = Record<string, IBUConfig>;

const DEFAULT_BU_REGIONS: BURegionMap = {
  'Enterprise-West': {
    regions: { 'NorCal': {}, 'PacNW': {}, 'SoCal': {}, 'Southwest': {} }
  },
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function migrateBURegions(raw: Record<string, any>): BURegionMap {
  const result: BURegionMap = {};
  for (const [bu, val] of Object.entries(raw)) {
    if (Array.isArray(val)) {
      const regions: Record<string, IRegionConfig> = {};
      (val as string[]).forEach(r => { regions[r] = {}; });
      result[bu] = { regions };
    } else if (val && typeof val === 'object' && 'regions' in val) {
      result[bu] = val as IBUConfig;
    } else {
      result[bu] = { regions: {} };
    }
  }
  return result;
}

export class ConfigService {
  constructor(private _sp: SPFI) {}

  public async getSuperUsers(): Promise<string[]> {
    try {
      // Use SRTSuperUsers key; fall back to SuperUsers (POC Manager key) if not yet seeded
      let items: { Value: string }[] = await this._sp.web.lists
        .getByTitle('AppConfig').items
        .filter("Title eq 'SRTSuperUsers'").select('Value')();
      if (items.length === 0 || !items[0].Value) {
        items = await this._sp.web.lists
          .getByTitle('AppConfig').items
          .filter("Title eq 'SuperUsers'").select('Value')();
      }
      if (items.length > 0 && items[0].Value) {
        return items[0].Value.split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
      }
    } catch { /* fall through */ }
    return [];
  }

  public async isSuperUser(email: string): Promise<boolean> {
    const users = await this.getSuperUsers();
    return users.includes(email.trim().toLowerCase());
  }

  public async saveSuperUsers(users: string[]): Promise<void> {
    const items: { Id: number }[] = await this._sp.web.lists
      .getByTitle('AppConfig').items
      .filter("Title eq 'SRTSuperUsers'").select('Id')();
    if (items.length > 0) {
      await this._sp.web.lists.getByTitle('AppConfig').items
        .getById(items[0].Id).update({ Value: users.join(',') });
    } else {
      await this._sp.web.lists.getByTitle('AppConfig').items
        .add({ Title: 'SRTSuperUsers', Value: users.join(',') });
    }
  }

  public async getSolutions(): Promise<ISolutionDef[]> {
    try {
      const items: { Value: string }[] = await this._sp.web.lists
        .getByTitle('AppConfig').items
        .filter("Title eq 'Solutions'").select('Value')();
      if (items.length === 0 || !items[0].Value) return [...SOLUTIONS];
      return JSON.parse(items[0].Value) as ISolutionDef[];
    } catch {
      return [...SOLUTIONS];
    }
  }

  public async saveBURegions(buRegions: BURegionMap): Promise<void> {
    const items: { Id: number }[] = await this._sp.web.lists
      .getByTitle('AppConfig').items
      .filter("Title eq 'BURegions'").select('Id')();
    if (items.length > 0) {
      await this._sp.web.lists.getByTitle('AppConfig').items
        .getById(items[0].Id).update({ Value: JSON.stringify(buRegions) });
    } else {
      await this._sp.web.lists.getByTitle('AppConfig').items
        .add({ Title: 'BURegions', Value: JSON.stringify(buRegions) });
    }
  }

  public async getBURegions(): Promise<BURegionMap> {
    try {
      const items: { Value: string }[] = await this._sp.web.lists
        .getByTitle('AppConfig').items
        .filter("Title eq 'BURegions'").select('Value')();
      if (items.length > 0 && items[0].Value) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return migrateBURegions(JSON.parse(items[0].Value) as Record<string, any>);
      }
    } catch { /* fall through */ }
    return { ...DEFAULT_BU_REGIONS };
  }
}
