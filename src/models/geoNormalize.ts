// Gold-standard geo canonicalization.
//
// Charlie's source spreadsheet had baked-in variants ("Enterprise-West" vs "Enterprise West",
// "Midwest" vs "Midwest " with a trailing space, casing twins, etc.). Rather than rewrite the
// underlying CSERequests data, we snap values to the canonical Gold-standard form at READ time so
// the dashboard filters, matching, display, and reports all speak one vocabulary — and any future
// dirty entry self-heals to canonical.
//
// Canonical source = the AppConfig BURegions map (ConfigService.getBURegions). BU keys are the
// canonical Business Units; the region keys under each BU are the canonical Regions.

import { BURegionMap } from '../services/ConfigService';

// Fold a value to a comparison key: lowercase, strip everything but a–z/0–9.
// So "Enterprise-West", "Enterprise West", and "enterprise  west" all key to "enterprisewest".
export const normKey = (s: string): string => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

// Trim + collapse internal whitespace (kills the "Midwest " trailing-space twin) without
// changing the visible casing.
const tidy = (s: string): string => (s || '').trim().replace(/\s+/g, ' ');

export interface GeoCanonicalizer {
  bu: (v: string) => string;
  region: (v: string) => string;
}

export function buildGeoCanonicalizer(buRegions: BURegionMap): GeoCanonicalizer {
  const buMap = new Map<string, string>();      // normKey → canonical BU
  const regionMap = new Map<string, string>();  // normKey → canonical Region
  for (const [bu, cfg] of Object.entries(buRegions || {})) {
    if (bu) buMap.set(normKey(bu), bu);
    for (const region of Object.keys(cfg?.regions || {})) {
      if (region) regionMap.set(normKey(region), region);
    }
  }
  const snap = (map: Map<string, string>, v: string): string => {
    const t = tidy(v);
    if (!t) return '';
    return map.get(normKey(t)) || t; // canonical form when known, else the tidied value
  };
  return {
    bu: v => snap(buMap, v),
    region: v => snap(regionMap, v),
  };
}

// Canonicalize a list, then collapse case/whitespace duplicates (by normKey), keeping the first
// representative. Use for filter-dropdown options so variants show as a single entry.
export function dedupeByKey(values: string[]): string[] {
  const seen = new Map<string, string>();
  for (const v of values) {
    const k = normKey(v);
    if (k && !seen.has(k)) seen.set(k, v);
  }
  return Array.from(seen.values()).sort();
}
