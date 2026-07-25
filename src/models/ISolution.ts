export interface ISolutionDef {
  code: string;
  name: string;
  category: string;
}

export const SOLUTIONS: ISolutionDef[] = [
  { code: 'MWRL', name: 'Mist Wireless',               category: 'Mist AI' },
  { code: 'MWRD', name: 'Mist EX Switching',           category: 'Mist AI' },
  { code: 'MSRX', name: 'Mist SRX SD-WAN',             category: 'Mist AI' },
  { code: 'MSSR', name: 'Mist SSR SD-WAN',             category: 'Mist AI' },
  { code: 'CWRL', name: 'Central Wireless',             category: 'Aruba Central' },
  { code: 'CWRD', name: 'Central Wired',                category: 'Aruba Central' },
  { code: 'CTNL', name: 'Central AOS10 Tunneled',      category: 'Aruba Central' },
  { code: 'HWRL', name: 'HPE WLAN AOS10',              category: 'HPE WLAN' },
  { code: 'HMIG', name: 'HPE WLAN AOS8→AOS10 Mig.',   category: 'HPE WLAN' },
  { code: 'HCOP', name: 'HPE WLAN CoP',                category: 'HPE WLAN' },
  { code: 'HCXS', name: 'HPE CX Switching',            category: 'HPE Switching' },
  { code: 'HEXS', name: 'HPE EX Switching',            category: 'HPE Switching' },
  { code: 'DCJS', name: 'Juniper QFX Switching',       category: 'Data Center' },
  { code: 'DCAS', name: 'Aruba CX10K Switching',       category: 'Data Center' },
  { code: 'DCSA', name: 'Apstra',                      category: 'Data Center' },
  { code: 'DCSD', name: 'Data Center Assurance',       category: 'Data Center' },
  { code: 'RHMX', name: 'Routing MX',                  category: 'Routing' },
  { code: 'RHTX', name: 'Routing PTX',                 category: 'Routing' },
  { code: 'RHCX', name: 'Routing ACX',                 category: 'Routing' },
  { code: 'RSRD', name: 'Routing Director',             category: 'Routing' },
  { code: 'RSRA', name: 'Routing Assurance',           category: 'Routing' },
  { code: 'SSRX', name: 'SASE SRX Firewall',           category: 'SASE' },
  { code: 'SECW', name: 'SASE EdgeConnect SD-WAN',     category: 'SASE' },
  { code: 'SSSE', name: 'SASE Axis SSE',               category: 'SASE' },
  { code: 'NSAA', name: 'NAC Access Assurance',        category: 'NAC' },
  { code: 'NSCP', name: 'NAC ClearPass',               category: 'NAC' },
  { code: 'NSCN', name: 'NAC Central NAC',             category: 'NAC' },
  { code: 'AUXI', name: 'UXI (User Experience Insights)', category: 'AIOps' },
  { code: 'UD1',  name: 'User Defined 1',              category: 'User Defined' },
  { code: 'UD2',  name: 'User Defined 2',              category: 'User Defined' },
  { code: 'UD3',  name: 'User Defined 3',              category: 'User Defined' },
];

export const SOLUTION_CATEGORIES = Array.from(new Set(SOLUTIONS.map(s => s.category)));
