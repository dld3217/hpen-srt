import * as React from 'react';
import { useState, useRef } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphUser { displayName: string; mail: string; }

// Graph people search — "displayName:query" (needs the app's User.ReadBasic.All grant).
export async function searchGraphUsers(context: WebPartContext, query: string): Promise<IGraphUser[]> {
  try {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const response: any = await client.api('/users')
      .header('ConsistencyLevel', 'eventual')
      .search(`"displayName:${query.replace(/"/g, '')}"`)
      .select('displayName,mail').top(10).get();
    interface IGraphResult { displayName?: string; mail?: string; }
    return ((response.value || []) as IGraphResult[])
      .filter(u => !!u.mail)
      .map(u => ({ displayName: (u.displayName || '') as string, mail: (u.mail || '') as string }));
  } catch { return []; }
}

const INPUT: React.CSSProperties = { width: '100%', boxSizing: 'border-box', fontSize: 13, padding: '7px 10px', border: '1px solid #ccc', borderRadius: 4 };

// Self-contained people picker producing a "DisplayName / email" value. Shows a filled chip
// (mailto + ✎/✕) once picked, else a debounced Graph search input with a results dropdown.
export const PeoplePickerField: React.FC<{
  value: string;
  onChange: (val: string) => void;
  searchUsers: (query: string) => Promise<IGraphUser[]>;
  placeholder?: string;
}> = ({ value, onChange, searchUsers, placeholder }) => {
  const [editing, setEditing]       = useState(false);
  const [localValue, setLocalValue] = useState('');
  const [results, setResults]       = useState<IGraphUser[]>([]);
  const [searching, setSearching]   = useState(false);
  const debounceRef  = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const containerRef = useRef<HTMLDivElement>(null);

  const slashIdx    = value.indexOf(' / ');
  const showDisplay = !editing && slashIdx !== -1;
  const namePart    = showDisplay ? value.substring(0, slashIdx) : '';
  const emailPart   = showDisplay ? value.substring(slashIdx + 3) : '';

  const startEdit = (): void => { setLocalValue(slashIdx !== -1 ? value.substring(0, slashIdx) : value); setResults([]); setEditing(true); };
  const selectPerson = (p: IGraphUser): void => { onChange(`${p.displayName} / ${p.mail}`); setEditing(false); setResults([]); setLocalValue(''); };
  const handleLocalChange = (q: string): void => {
    setLocalValue(q);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (q.length < 3) { setResults([]); return; }
    debounceRef.current = setTimeout(() => {
      setSearching(true);
      searchUsers(q).then(r => setResults(r)).catch(() => setResults([])).finally(() => setSearching(false));
    }, 300);
  };
  const handleBlur = (e: React.FocusEvent<HTMLDivElement>): void => {
    if (containerRef.current?.contains(e.relatedTarget as Node)) return;
    if (!editing) return;
    onChange(localValue.trim()); setEditing(false); setResults([]); setLocalValue('');
  };

  return (
    <div ref={containerRef} onBlur={handleBlur} style={{ flex: 1, position: 'relative' }}>
      {showDisplay ? (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', background: '#f3f2f1', borderRadius: 4, border: '1px solid #ccc' }}>
          <a href={`mailto:${emailPart}`} style={{ fontWeight: 600, fontSize: 13, color: '#0078d4', textDecoration: 'none' }}>{namePart}</a>
          <span style={{ fontSize: 11, color: '#605e5c' }}>{emailPart}</span>
          <span style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
            <button type="button" onClick={startEdit} title="Change" style={{ fontSize: 11, color: '#0078d4', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✎</button>
            <button type="button" onClick={() => onChange('')} title="Clear" style={{ fontSize: 11, color: '#a4262c', background: 'none', border: 'none', cursor: 'pointer', padding: '0 4px' }}>✕</button>
          </span>
        </div>
      ) : (
        <>
          <input autoFocus={editing} value={localValue} onChange={e => handleLocalChange(e.target.value)}
            onFocus={() => { if (!editing) startEdit(); }} placeholder={placeholder || 'Search name…'} style={INPUT} />
          {searching && <span style={{ position: 'absolute', right: 8, top: 9, fontSize: 11, color: '#888' }}>…</span>}
          {results.length > 0 && (
            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #ccc', borderRadius: 4, zIndex: 1100, boxShadow: '0 4px 12px rgba(0,0,0,0.12)', maxHeight: 220, overflowY: 'auto' }}>
              {results.map(r => (
                <div key={r.mail} tabIndex={0} onMouseDown={() => selectPerson(r)} onKeyDown={e => e.key === 'Enter' && selectPerson(r)}
                  style={{ padding: '7px 12px', cursor: 'pointer', fontSize: 13, borderBottom: '1px solid #f0f0f0' }}
                  onMouseEnter={e => (e.currentTarget.style.background = '#f3f2f1')} onMouseLeave={e => (e.currentTarget.style.background = '#fff')}>
                  <span style={{ fontWeight: 600 }}>{r.displayName}</span>
                  <span style={{ color: '#888', marginLeft: 8, fontSize: 11 }}>{r.mail}</span>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
};
