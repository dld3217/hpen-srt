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
