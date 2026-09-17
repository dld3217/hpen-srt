import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ConfigService, BURegionMap, IBUConfig, ISSETeam, SpecialProjectsMap } from '../../../services/ConfigService';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISrtAdminPanelProps {
  sp: SPFI;
  context: WebPartContext;
  onClose: () => void;
}

interface IPdlMember { displayName: string; mail: string; jobTitle?: string; }

const TAB_BTN = (active: boolean): React.CSSProperties => ({
  flex: 1, padding: '7px 0', fontSize: 12, fontWeight: 600,
  background: active ? HPE_NAVY : '#f3f2f1',
  color: active ? '#fff' : '#605e5c',
  border: 'none', cursor: 'pointer',
  borderBottom: active ? `2px solid ${HPE_GREEN}` : '2px solid transparent',
});

export const SrtAdminPanel: React.FC<ISrtAdminPanelProps> = ({ sp, context, onClose }) => {
  const [tab, setTab]               = useState<'users' | 'bu' | 'teams' | 'special'>('users');
  const [message, setMessage]       = useState('');
  const [msgType, setMsgType]       = useState<'ok' | 'err'>('ok');

  // ── Super Users ──────────────────────────────────────────────────────────────
  const [superUsers, setSuperUsers] = useState<string[]>([]);
  const [newEmail, setNewEmail]     = useState('');
  const [usersLoading, setUsersLoading] = useState(true);
  const [saving, setSaving]         = useState(false);

  // ── SED Bypass ───────────────────────────────────────────────────────────────
  const [sedApprovalRequired, setSedApprovalRequired] = useState(true);
  const [sedSaving, setSedSaving]   = useState(false);

  // ── SSE Teams ────────────────────────────────────────────────────────────
  const [sseTeams, setSseTeams]         = useState<ISSETeam[]>([]);
  const [teamsLoading, setTeamsLoading] = useState(true);
  const [teamsSaving, setTeamsSaving]   = useState(false);
  const [editTeamIdx, setEditTeamIdx]   = useState<number | null>(null);
  const [editTeamEmail, setEditTeamEmail] = useState('');
  const [editTeamName, setEditTeamName] = useState('');
  const [showAddTeam, setShowAddTeam]   = useState(false);
  const [newTeamName, setNewTeamName]   = useState('');
  const [newTeamEmail, setNewTeamEmail] = useState('');

  // ── BU Config ─────────────────────────────────────────────────────────────
  const [buRegions, setBuRegions]   = useState<BURegionMap>({});
  const [buLoading, setBuLoading]   = useState(true);
  const [buSaving, setBuSaving]     = useState(false);
  const [editSedBu, setEditSedBu]   = useState<string | null>(null);
  const [editSedEmail, setEditSedEmail] = useState('');
  const [newBuName, setNewBuName]   = useState('');
  const [newBuSed, setNewBuSed]     = useState('');
  const [showAddBu, setShowAddBu]   = useState(false);

  // ── Special Projects (CIC / Marketing) ─────────────────────────────────────
  const [spMap, setSpMap]           = useState<SpecialProjectsMap>({});
  const [spLoading, setSpLoading]   = useState(true);
  const [spSaving, setSpSaving]     = useState(false);
  const [showAddCat, setShowAddCat] = useState(false);
  const [newCatName, setNewCatName] = useState('');
  const [editCat, setEditCat]       = useState<string | null>(null);
  const [editCatName, setEditCatName] = useState('');
  const [newInit, setNewInit]       = useState<Record<string, string>>({});
  const [editInit, setEditInit]     = useState<{ cat: string; idx: number } | null>(null);
  const [editInitName, setEditInitName] = useState('');

  // ── SRT SSEs (may create Special Projects) ─────────────────────────────────
  const [sseList, setSseList]       = useState<string[]>([]);
  const [sseLoading, setSseLoading] = useState(true);
  const [sseSaving, setSseSaving]   = useState(false);
  const [newSseEmail, setNewSseEmail] = useState('');

  // ── PDL Lookup ───────────────────────────────────────────────────────────
  const [pdlEmail, setPdlEmail]     = useState('');
  const [pdlResults, setPdlResults] = useState<IPdlMember[]>([]);
  const [pdlLoading, setPdlLoading] = useState(false);
  const [pdlError, setPdlError]     = useState('');

  const configSvc = React.useMemo(() => new ConfigService(sp), [sp]);
  const myEmail   = context.pageContext.user.email.toLowerCase();

  const showMsg = (text: string, type: 'ok' | 'err'): void => {
    setMessage(text); setMsgType(type);
    setTimeout(() => setMessage(''), 4000);
  };

  useEffect(() => {
    configSvc.getSuperUsers()
      .then(u => { setSuperUsers(u); setUsersLoading(false); })
      .catch(() => setUsersLoading(false));
    configSvc.getBURegions()
      .then(b => { setBuRegions(b); setBuLoading(false); })
      .catch(() => setBuLoading(false));
    configSvc.getSSETeams()
      .then(t => { setSseTeams(t); setTeamsLoading(false); })
      .catch(() => setTeamsLoading(false));
    configSvc.getSEDApprovalRequired()
      .then(r => setSedApprovalRequired(r))
      .catch(() => undefined);
    configSvc.getSpecialProjects()
      .then(m => { setSpMap(m); setSpLoading(false); })
      .catch(() => setSpLoading(false));
    configSvc.getSSEs()
      .then(u => { setSseList(u); setSseLoading(false); })
      .catch(() => setSseLoading(false));
  }, []);

  // ── Super User actions ────────────────────────────────────────────────────
  const saveUsers = async (users: string[]): Promise<void> => {
    setSaving(true);
    try {
      await configSvc.saveSuperUsers(users);
      setSuperUsers(users);
      showMsg('Saved.', 'ok');
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setSaving(false); }
  };

  const addUser = (): void => {
    const email = newEmail.trim().toLowerCase();
    if (!email || superUsers.includes(email)) return;
    saveUsers([...superUsers, email]).catch(() => undefined);
    setNewEmail('');
  };

  const removeUser = (email: string): void => {
    if (email === myEmail) { showMsg('You cannot remove yourself.', 'err'); return; }
    saveUsers(superUsers.filter(u => u !== email)).catch(() => undefined);
  };

  // ── SED Bypass actions ────────────────────────────────────────────────────
  const toggleSedApproval = async (required: boolean): Promise<void> => {
    setSedSaving(true);
    try {
      await configSvc.saveSEDApprovalRequired(required);
      setSedApprovalRequired(required);
      showMsg(
        required
          ? 'SED approval re-enabled — new requests will be held Pending.'
          : 'SED approval bypassed — new requests route directly to the SSE.',
        'ok'
      );
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setSedSaving(false); }
  };

  // ── BU Config actions ─────────────────────────────────────────────────────
  const saveBU = async (updated: BURegionMap): Promise<void> => {
    setBuSaving(true);
    try {
      await configSvc.saveBURegions(updated);
      setBuRegions(updated);
      showMsg('BU config saved.', 'ok');
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setBuSaving(false); }
  };

  const commitSedEdit = (): void => {
    if (!editSedBu) return;
    const updated: BURegionMap = {
      ...buRegions,
      [editSedBu]: { ...(buRegions[editSedBu] as IBUConfig), sedEmail: editSedEmail.trim().toLowerCase() },
    };
    saveBU(updated).catch(() => undefined);
    setEditSedBu(null);
  };

  const addBu = (): void => {
    const name = newBuName.trim();
    if (!name) return;
    const updated: BURegionMap = {
      ...buRegions,
      [name]: { sedEmail: newBuSed.trim().toLowerCase(), regions: {} },
    };
    saveBU(updated).catch(() => undefined);
    setNewBuName(''); setNewBuSed(''); setShowAddBu(false);
  };

  // ── SSE Teams actions ─────────────────────────────────────────────────────
  const saveTeams = async (teams: ISSETeam[]): Promise<void> => {
    setTeamsSaving(true);
    try {
      await configSvc.saveSSETeams(teams);
      setSseTeams(teams);
      showMsg('SSE Teams saved.', 'ok');
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setTeamsSaving(false); }
  };

  const commitTeamEdit = (): void => {
    if (editTeamIdx === null) return;
    const updated = sseTeams.map((t, i) =>
      i === editTeamIdx ? { name: editTeamName.trim() || t.name, managerEmail: editTeamEmail.trim().toLowerCase() } : t
    );
    saveTeams(updated).catch(() => undefined);
    setEditTeamIdx(null);
  };

  const addTeam = (): void => {
    const name = newTeamName.trim();
    if (!name) return;
    saveTeams([...sseTeams, { name, managerEmail: newTeamEmail.trim().toLowerCase() }]).catch(() => undefined);
    setNewTeamName(''); setNewTeamEmail(''); setShowAddTeam(false);
  };

  const removeTeam = (idx: number): void => {
    saveTeams(sseTeams.filter((_, i) => i !== idx)).catch(() => undefined);
  };

  // ── Special Projects actions ────────────────────────────────────────────────
  const saveSp = async (updated: SpecialProjectsMap): Promise<void> => {
    setSpSaving(true);
    try {
      await configSvc.saveSpecialProjects(updated);
      setSpMap(updated);
      showMsg('Special Projects saved.', 'ok');
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setSpSaving(false); }
  };

  const addCategory = (): void => {
    const name = newCatName.trim();
    if (!name) return;
    if (Object.keys(spMap).some(c => c.toLowerCase() === name.toLowerCase())) {
      showMsg('That category already exists.', 'err'); return;
    }
    saveSp({ ...spMap, [name]: [] }).catch(() => undefined);
    setNewCatName(''); setShowAddCat(false);
  };

  const commitCatRename = (): void => {
    if (!editCat) return;
    const name = editCatName.trim();
    if (!name || name === editCat) { setEditCat(null); return; }
    if (Object.keys(spMap).some(c => c !== editCat && c.toLowerCase() === name.toLowerCase())) {
      showMsg('A category with that name already exists.', 'err'); return;
    }
    // Rebuild preserving insertion order, swapping the renamed key.
    const updated: SpecialProjectsMap = {};
    for (const [c, inits] of Object.entries(spMap)) updated[c === editCat ? name : c] = inits;
    saveSp(updated).catch(() => undefined);
    setEditCat(null);
  };

  const retireCategory = (cat: string): void => {
    if (!window.confirm(`Retire "${cat}" from the picker? Existing records keep their values; this only removes it from the dropdown going forward.`)) return;
    const updated = { ...spMap }; delete updated[cat];
    saveSp(updated).catch(() => undefined);
  };

  const addInitiative = (cat: string): void => {
    const name = (newInit[cat] || '').trim();
    if (!name) return;
    const existing = spMap[cat] || [];
    if (existing.some(i => i.toLowerCase() === name.toLowerCase())) {
      setNewInit({ ...newInit, [cat]: '' }); return;   // already present — dedup silently
    }
    saveSp({ ...spMap, [cat]: [...existing, name] }).catch(() => undefined);
    setNewInit({ ...newInit, [cat]: '' });
  };

  const commitInitRename = (): void => {
    if (!editInit) return;
    const { cat, idx } = editInit;
    const name = editInitName.trim();
    const list = spMap[cat] || [];
    if (!name) { setEditInit(null); return; }
    // Renaming onto an existing sibling (case-insensitive) MERGES — collapse the duplicate.
    const dup = list.some((i, k) => k !== idx && i.toLowerCase() === name.toLowerCase());
    const updatedList = dup ? list.filter((_, k) => k !== idx) : list.map((i, k) => k === idx ? name : i);
    saveSp({ ...spMap, [cat]: updatedList }).catch(() => undefined);
    setEditInit(null);
  };

  const removeInitiative = (cat: string, idx: number): void => {
    const list = spMap[cat] || [];
    saveSp({ ...spMap, [cat]: list.filter((_, k) => k !== idx) }).catch(() => undefined);
  };

  // ── SRT SSEs actions ────────────────────────────────────────────────────────
  const saveSses = async (users: string[]): Promise<void> => {
    setSseSaving(true);
    try {
      await configSvc.saveSSEs(users);
      setSseList(users);
      showMsg('SSE list saved.', 'ok');
    } catch (e) {
      showMsg(`Save failed: ${(e as Error).message}`, 'err');
    } finally { setSseSaving(false); }
  };

  const addSse = (): void => {
    const email = newSseEmail.trim().toLowerCase();
    if (!email || sseList.includes(email)) return;
    saveSses([...sseList, email]).catch(() => undefined);
    setNewSseEmail('');
  };

  const removeSse = (email: string): void => {
    saveSses(sseList.filter(u => u !== email)).catch(() => undefined);
  };

  // ── PDL Lookup ────────────────────────────────────────────────────────────
  const handlePdlLookup = async (): Promise<void> => {
    if (!pdlEmail.trim()) return;
    setPdlLoading(true); setPdlError(''); setPdlResults([]);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const client: any = await context.msGraphClientFactory.getClient('3');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const groupsResp: any = await client.api('/groups')
        .filter(`mail eq '${pdlEmail.trim()}'`)
        .select('id,displayName,mail')
        .get();
      if (!groupsResp.value?.length) {
        setPdlError('No group found with that email. It may be a legacy Exchange DL not synced to Azure AD.');
        return;
      }
      const groupId: string = groupsResp.value[0].id;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const membersResp: any = await client.api(`/groups/${groupId}/members`)
        .select('displayName,mail,jobTitle')
        .get();
      setPdlResults(membersResp.value || []);
    } catch (e) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const err = e as any;
      if (err?.statusCode === 403 || err?.code === 'Authorization_RequestDenied') {
        setPdlError('Access denied (403) — GroupMember.Read.All permission not yet approved for this app. A tenant admin needs to approve it in SP Admin → API access.');
      } else if (err?.statusCode === 404) {
        setPdlError('Group not found — verify the email address is correct.');
      } else {
        setPdlError(`Error: ${err?.message || String(e)}`);
      }
    } finally { setPdlLoading(false); }
  };

  const buKeys = Object.keys(buRegions).sort();

  return (
    <>
      <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.35)', zIndex: 1000 }} />
      <div style={{
        position: 'fixed', top: 0, right: 0, bottom: 0, width: 420,
        background: '#fff', boxShadow: '-4px 0 20px rgba(0,0,0,0.18)',
        zIndex: 1001, display: 'flex', flexDirection: 'column', fontFamily: 'inherit',
      }}>
        {/* Header */}
        <div style={{ background: HPE_NAVY, color: '#fff', padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 3, height: 20, background: HPE_GREEN, borderRadius: 2 }} />
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>SRT Admin</div>
              <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>Configuration &amp; Access Control</div>
            </div>
          </div>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer', lineHeight: 1, padding: 4 }}>✕</button>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', borderBottom: '1px solid #edebe9' }}>
          <button style={TAB_BTN(tab === 'users')} onClick={() => setTab('users')}>Super Users</button>
          <button style={TAB_BTN(tab === 'bu')} onClick={() => setTab('bu')}>BU Config</button>
          <button style={TAB_BTN(tab === 'teams')} onClick={() => setTab('teams')}>SSE Teams</button>
          <button style={TAB_BTN(tab === 'special')} onClick={() => setTab('special')}>⭐ Special</button>
        </div>

        {/* Message bar */}
        {message && (
          <div style={{ margin: '10px 16px 0', padding: '8px 12px', borderRadius: 4, fontSize: 12,
            background: msgType === 'ok' ? '#dff6dd' : '#fde7e9',
            color: msgType === 'ok' ? '#107c10' : '#a4262c',
            border: `1px solid ${msgType === 'ok' ? '#107c10' : '#a4262c'}` }}>
            {message}
          </div>
        )}

        {/* Body */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px' }}>

          {/* ── Super Users tab ── */}
          {tab === 'users' && (
            <>
              <p style={{ fontSize: 12, color: '#605e5c', marginTop: 0 }}>
                Super users see all SSE requests and can Accept, Decline, edit Status and Solutions.
                Regular SEs see only their own submissions. Currently <strong>{superUsers.length}</strong> super user{superUsers.length !== 1 ? 's' : ''}.
              </p>
              {usersLoading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {superUsers.map(email => (
                    <div key={email} style={{ display: 'flex', alignItems: 'center', gap: 10,
                      padding: '8px 12px', background: '#faf9f8', border: '1px solid #edebe9', borderRadius: 4 }}>
                      <span style={{ flex: 1, fontSize: 13 }}>{email}</span>
                      {email === myEmail && <span style={{ fontSize: 11, color: '#888', fontStyle: 'italic' }}>(you)</span>}
                      <button onClick={() => removeUser(email)} disabled={saving}
                        style={{ background: 'none', border: 'none', color: '#d13438', fontSize: 16, cursor: 'pointer', lineHeight: 1, padding: '2px 4px' }}
                        title="Remove">✕</button>
                    </div>
                  ))}
                  {superUsers.length === 0 && <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No super users yet.</div>}
                </div>
              )}
              <div style={{ marginTop: 16, display: 'flex', gap: 8 }}>
                <input type="email" placeholder="user@hpe.com" value={newEmail}
                  onChange={e => setNewEmail(e.target.value)}
                  onKeyDown={e => { if (e.key === 'Enter') addUser(); }}
                  style={{ flex: 1, fontSize: 13, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4 }} />
                <button onClick={addUser} disabled={saving || !newEmail.trim()}
                  style={{ padding: '6px 16px', background: HPE_NAVY, color: '#fff', border: 'none', borderRadius: 4,
                    fontSize: 13, fontWeight: 600, cursor: saving || !newEmail.trim() ? 'not-allowed' : 'pointer',
                    opacity: saving || !newEmail.trim() ? 0.6 : 1 }}>
                  {saving ? '…' : 'Add'}
                </button>
              </div>

              {/* Workflow Settings */}
              <div style={{ marginTop: 24, paddingTop: 16, borderTop: '1px solid #edebe9' }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 10, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4 }}>
                  Workflow Settings
                </div>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: sedSaving ? 'not-allowed' : 'pointer' }}>
                  <input type="checkbox" checked={sedApprovalRequired} disabled={sedSaving}
                    onChange={e => toggleSedApproval(e.target.checked).catch(() => undefined)}
                    style={{ accentColor: HPE_NAVY, width: 14, height: 14, marginTop: 2, flexShrink: 0 }} />
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: '#323130' }}>Require SED/GM approval</div>
                    <div style={{ fontSize: 11, color: '#605e5c', marginTop: 2, lineHeight: 1.4 }}>
                      When enabled, new requests are held <strong>Pending</strong> until a SED accepts them.
                      When disabled, requests route directly to the SSE as <strong>Accepted</strong>.
                    </div>
                    {!sedApprovalRequired && (
                      <div style={{ marginTop: 6, fontSize: 11, color: '#d83b01', fontWeight: 600 }}>
                        ⚠ Bypass is active — SED approval is skipped for all new requests.
                      </div>
                    )}
                  </div>
                </label>
              </div>
            </>
          )}

          {/* ── BU Config tab ── */}
          {tab === 'bu' && (
            <>
              {/* PDL Lookup */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4 }}>
                  PDL Lookup
                </div>
                <p style={{ fontSize: 12, color: '#605e5c', margin: '0 0 10px' }}>
                  Enter a team PDL email to resolve its members via Microsoft Graph. Use this to verify SE/SEM rosters before wiring up BU configs.
                </p>
                <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                  <input type="email" placeholder="HPENEntWestSTOLASE@hpe.com" value={pdlEmail}
                    onChange={e => setPdlEmail(e.target.value)}
                    onKeyDown={e => { if (e.key === 'Enter') handlePdlLookup().catch(() => undefined); }}
                    style={{ flex: 1, fontSize: 12, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4 }} />
                  <button onClick={() => handlePdlLookup().catch(() => undefined)} disabled={pdlLoading || !pdlEmail.trim()}
                    style={{ padding: '6px 14px', background: HPE_NAVY, color: '#fff', border: 'none', borderRadius: 4,
                      fontSize: 12, fontWeight: 600, cursor: pdlLoading || !pdlEmail.trim() ? 'not-allowed' : 'pointer',
                      opacity: pdlLoading || !pdlEmail.trim() ? 0.6 : 1, whiteSpace: 'nowrap' }}>
                    {pdlLoading ? 'Looking up…' : 'Look Up'}
                  </button>
                </div>
                {pdlError && (
                  <div style={{ fontSize: 12, color: '#a4262c', background: '#fde7e9', border: '1px solid #a4262c', borderRadius: 4, padding: '8px 10px', marginBottom: 8 }}>
                    {pdlError}
                  </div>
                )}
                {pdlResults.length > 0 && (
                  <div style={{ border: '1px solid #edebe9', borderRadius: 4, overflow: 'hidden' }}>
                    <div style={{ padding: '6px 10px', background: '#f3f2f1', fontSize: 11, fontWeight: 700, color: '#323130' }}>
                      {pdlResults.length} member{pdlResults.length !== 1 ? 's' : ''} found
                    </div>
                    {pdlResults.map((m, i) => (
                      <div key={m.mail || i} style={{ padding: '7px 10px', borderTop: i === 0 ? 'none' : '1px solid #f0f0f0',
                        background: i % 2 === 0 ? '#fff' : '#faf9f8' }}>
                        <div style={{ fontSize: 12, fontWeight: 600 }}>{m.displayName}</div>
                        <div style={{ fontSize: 11, color: '#605e5c' }}>{m.mail}</div>
                        {m.jobTitle && <div style={{ fontSize: 10, color: '#888' }}>{m.jobTitle}</div>}
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* BU Regions */}
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span>BU Configuration</span>
                  <button onClick={() => setShowAddBu(v => !v)}
                    style={{ fontSize: 11, padding: '3px 10px', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600, textTransform: 'none', letterSpacing: 0 }}>
                    + Add BU
                  </button>
                </div>

                {showAddBu && (
                  <div style={{ marginBottom: 12, padding: '12px 14px', background: '#f0f9f4', border: `1px solid ${HPE_GREEN}`, borderRadius: 4 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8 }}>New Business Unit</div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                      <input placeholder="BU Name (e.g. Enterprise-East)" value={newBuName}
                        onChange={e => setNewBuName(e.target.value)}
                        style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                      <input placeholder="SED email (e.g. john.doe@hpe.com)" value={newBuSed}
                        onChange={e => setNewBuSed(e.target.value)}
                        style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button onClick={addBu} disabled={!newBuName.trim() || buSaving}
                          style={{ flex: 1, padding: '5px 0', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                          {buSaving ? '…' : 'Save BU'}
                        </button>
                        <button onClick={() => { setShowAddBu(false); setNewBuName(''); setNewBuSed(''); }}
                          style={{ flex: 1, padding: '5px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                          Cancel
                        </button>
                      </div>
                    </div>
                  </div>
                )}

                {buLoading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {buKeys.length === 0 && <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No BUs configured yet.</div>}
                    {buKeys.map(bu => {
                      const cfg = buRegions[bu] as IBUConfig;
                      const isEditing = editSedBu === bu;
                      const regionCount = Object.keys(cfg?.regions || {}).length;
                      return (
                        <div key={bu} style={{ border: '1px solid #edebe9', borderRadius: 4, overflow: 'hidden' }}>
                          <div style={{ padding: '10px 12px', background: '#f3f2f1', display: 'flex', alignItems: 'center', gap: 8 }}>
                            <div style={{ flex: 1 }}>
                              <div style={{ fontSize: 13, fontWeight: 700 }}>{bu}</div>
                              <div style={{ fontSize: 11, color: '#888' }}>{regionCount} region{regionCount !== 1 ? 's' : ''}</div>
                            </div>
                            {!isEditing && (
                              <button onClick={() => { setEditSedBu(bu); setEditSedEmail(cfg?.sedEmail || ''); }}
                                style={{ fontSize: 11, padding: '3px 10px', background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 3, cursor: 'pointer' }}>
                                Edit SED
                              </button>
                            )}
                          </div>
                          <div style={{ padding: '10px 12px' }}>
                            {isEditing ? (
                              <div>
                                <div style={{ fontSize: 11, color: '#555', marginBottom: 4 }}>SED / GM Email</div>
                                <div style={{ display: 'flex', gap: 6 }}>
                                  <input value={editSedEmail} onChange={e => setEditSedEmail(e.target.value)}
                                    placeholder="sed@hpe.com" autoFocus
                                    style={{ flex: 1, fontSize: 12, padding: '5px 8px', border: '1px solid #0078d4', borderRadius: 4 }} />
                                  <button onClick={commitSedEdit} disabled={buSaving}
                                    style={{ padding: '5px 12px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                    {buSaving ? '…' : 'Save'}
                                  </button>
                                  <button onClick={() => setEditSedBu(null)}
                                    style={{ padding: '5px 10px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                                    ✕
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <div style={{ fontSize: 12 }}>
                                <span style={{ color: '#888' }}>SED/GM: </span>
                                <span style={{ color: cfg?.sedEmail ? '#323130' : '#aaa', fontStyle: cfg?.sedEmail ? 'normal' : 'italic' }}>
                                  {cfg?.sedEmail || 'not set'}
                                </span>
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </>
          )}
          {/* ── SSE Teams tab ── */}
          {tab === 'teams' && (
            <>
              <p style={{ fontSize: 12, color: '#605e5c', marginTop: 0 }}>
                Map each SSE specialty type to the manager who oversees that team. The SE selects a specialty on the request form and the corresponding manager is notified.
              </p>
              <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>Specialty Teams</span>
                <button onClick={() => setShowAddTeam(v => !v)}
                  style={{ fontSize: 11, padding: '3px 10px', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600, textTransform: 'none', letterSpacing: 0 }}>
                  + Add Team
                </button>
              </div>

              {showAddTeam && (
                <div style={{ marginBottom: 12, padding: '12px 14px', background: '#f0f9f4', border: `1px solid ${HPE_GREEN}`, borderRadius: 4 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8 }}>New Specialty Team</div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <input placeholder="Specialty name (e.g. Mist)" value={newTeamName}
                      onChange={e => setNewTeamName(e.target.value)}
                      style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                    <input placeholder="Manager email (e.g. mike.bruno@hpe.com)" value={newTeamEmail}
                      onChange={e => setNewTeamEmail(e.target.value)}
                      style={{ fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                    <div style={{ display: 'flex', gap: 6 }}>
                      <button onClick={addTeam} disabled={!newTeamName.trim() || teamsSaving}
                        style={{ flex: 1, padding: '5px 0', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                        {teamsSaving ? '…' : 'Save'}
                      </button>
                      <button onClick={() => { setShowAddTeam(false); setNewTeamName(''); setNewTeamEmail(''); }}
                        style={{ flex: 1, padding: '5px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                        Cancel
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {teamsLoading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {sseTeams.length === 0 && <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No teams configured yet.</div>}
                  {sseTeams.map((team, idx) => {
                    const isEditing = editTeamIdx === idx;
                    return (
                      <div key={idx} style={{ border: '1px solid #edebe9', borderRadius: 4, overflow: 'hidden' }}>
                        <div style={{ padding: '10px 12px', background: '#f3f2f1', display: 'flex', alignItems: 'center', gap: 8 }}>
                          <div style={{ flex: 1, fontSize: 13, fontWeight: 700 }}>{team.name}</div>
                          {!isEditing && (
                            <>
                              <button onClick={() => { setEditTeamIdx(idx); setEditTeamName(team.name); setEditTeamEmail(team.managerEmail); }}
                                style={{ fontSize: 11, padding: '3px 10px', background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 3, cursor: 'pointer' }}>
                                Edit
                              </button>
                              <button onClick={() => removeTeam(idx)} disabled={teamsSaving}
                                style={{ background: 'none', border: 'none', color: '#d13438', fontSize: 16, cursor: 'pointer', lineHeight: 1, padding: '2px 4px' }}
                                title="Remove team">✕</button>
                            </>
                          )}
                        </div>
                        <div style={{ padding: '10px 12px' }}>
                          {isEditing ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                              <div>
                                <div style={{ fontSize: 11, color: '#555', marginBottom: 3 }}>Specialty Name</div>
                                <input value={editTeamName} onChange={e => setEditTeamName(e.target.value)}
                                  style={{ width: '100%', boxSizing: 'border-box', fontSize: 12, padding: '5px 8px', border: '1px solid #0078d4', borderRadius: 4 }} />
                              </div>
                              <div>
                                <div style={{ fontSize: 11, color: '#555', marginBottom: 3 }}>Manager Email</div>
                                <input value={editTeamEmail} onChange={e => setEditTeamEmail(e.target.value)}
                                  placeholder="manager@hpe.com" autoFocus
                                  style={{ width: '100%', boxSizing: 'border-box', fontSize: 12, padding: '5px 8px', border: '1px solid #0078d4', borderRadius: 4 }} />
                              </div>
                              <div style={{ display: 'flex', gap: 6 }}>
                                <button onClick={commitTeamEdit} disabled={teamsSaving}
                                  style={{ flex: 1, padding: '5px 0', background: '#107c10', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                                  {teamsSaving ? '…' : 'Save'}
                                </button>
                                <button onClick={() => setEditTeamIdx(null)}
                                  style={{ flex: 1, padding: '5px 0', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                                  Cancel
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div style={{ fontSize: 12 }}>
                              <span style={{ color: '#888' }}>Manager: </span>
                              <span style={{ color: team.managerEmail ? '#323130' : '#aaa', fontStyle: team.managerEmail ? 'normal' : 'italic' }}>
                                {team.managerEmail || 'not set'}
                              </span>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          )}

          {/* ── Special Projects tab ── */}
          {tab === 'special' && (
            <>
              {/* Who can create Special Projects */}
              <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4 }}>
                SSEs — can create Special Projects
              </div>
              <p style={{ fontSize: 11, color: '#605e5c', margin: '0 0 8px' }}>
                SSEs in the <strong>Contact Directory</strong> automatically get the <strong>⭐ New Special Project</strong> button — no need to list them here.
                Use this list only to grant it to someone <em>not</em> in the directory. Admins and SEDs always have it.
              </p>
              {sseLoading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {sseList.map(email => (
                    <div key={email} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: '#faf9f8', border: '1px solid #edebe9', borderRadius: 4 }}>
                      <span style={{ flex: 1, fontSize: 13 }}>{email}</span>
                      <button onClick={() => removeSse(email)} disabled={sseSaving}
                        style={{ background: 'none', border: 'none', color: '#d13438', fontSize: 16, cursor: 'pointer', lineHeight: 1, padding: '2px 4px' }} title="Remove">✕</button>
                    </div>
                  ))}
                  {sseList.length === 0 && <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No SSEs added yet — only admins/SEDs can create Special Projects.</div>}
                </div>
              )}
              <div style={{ marginTop: 10, marginBottom: 22, paddingBottom: 16, borderBottom: '1px solid #edebe9', display: 'flex', gap: 8 }}>
                <input type="email" placeholder="sse@hpe.com" value={newSseEmail}
                  onChange={e => setNewSseEmail(e.target.value)}
                  onKeyDown={e => { if (e.key === 'Enter') addSse(); }}
                  style={{ flex: 1, fontSize: 13, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4 }} />
                <button onClick={addSse} disabled={sseSaving || !newSseEmail.trim()}
                  style={{ padding: '6px 16px', background: HPE_NAVY, color: '#fff', border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: sseSaving || !newSseEmail.trim() ? 'not-allowed' : 'pointer', opacity: sseSaving || !newSseEmail.trim() ? 0.6 : 1 }}>
                  {sseSaving ? '…' : 'Add'}
                </button>
              </div>
              <p style={{ fontSize: 12, color: '#605e5c', marginTop: 0 }}>
                Non-geo engagements — CICs (Houston, San Jose…) and Marketing — that SSEs are pulled into
                across regions. Tracked by <strong>Category → Initiative</strong>; they never roll up into
                geo BU/Region reporting. Requesters can add a new initiative on the fly, and it appears here
                for you to rename, merge, or retire.
              </p>
              <div style={{ fontSize: 12, fontWeight: 700, color: HPE_NAVY, textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: 8, borderBottom: `2px solid ${HPE_GREEN}`, paddingBottom: 4, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span>Categories &amp; Initiatives</span>
                <button onClick={() => setShowAddCat(v => !v)}
                  style={{ fontSize: 11, padding: '3px 10px', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 3, cursor: 'pointer', fontWeight: 600, textTransform: 'none', letterSpacing: 0 }}>
                  + Add Category
                </button>
              </div>

              {showAddCat && (
                <div style={{ marginBottom: 12, padding: '12px 14px', background: '#f0f9f4', border: `1px solid ${HPE_GREEN}`, borderRadius: 4 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8 }}>New Category</div>
                  <div style={{ display: 'flex', gap: 6 }}>
                    <input placeholder="e.g. Austin CIC" value={newCatName}
                      onChange={e => setNewCatName(e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') addCategory(); }}
                      style={{ flex: 1, fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                    <button onClick={addCategory} disabled={!newCatName.trim() || spSaving}
                      style={{ padding: '5px 14px', background: HPE_GREEN, color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
                      {spSaving ? '…' : 'Add'}
                    </button>
                    <button onClick={() => { setShowAddCat(false); setNewCatName(''); }}
                      style={{ padding: '5px 10px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>
                      Cancel
                    </button>
                  </div>
                </div>
              )}

              {spLoading ? <div style={{ color: '#888', fontSize: 13 }}>Loading…</div> : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {Object.keys(spMap).length === 0 && <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No categories yet.</div>}
                  {Object.keys(spMap).map(cat => {
                    const inits = spMap[cat] || [];
                    const isEditingCat = editCat === cat;
                    return (
                      <div key={cat} style={{ border: '1px solid #edebe9', borderRadius: 4, overflow: 'hidden' }}>
                        <div style={{ padding: '10px 12px', background: '#f3f2f1', display: 'flex', alignItems: 'center', gap: 8 }}>
                          {isEditingCat ? (
                            <>
                              <input value={editCatName} onChange={e => setEditCatName(e.target.value)} autoFocus
                                onKeyDown={e => { if (e.key === 'Enter') commitCatRename(); }}
                                style={{ flex: 1, fontSize: 13, padding: '4px 8px', border: '1px solid #0078d4', borderRadius: 4 }} />
                              <button onClick={commitCatRename} disabled={spSaving}
                                style={{ padding: '4px 10px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Save</button>
                              <button onClick={() => setEditCat(null)}
                                style={{ padding: '4px 8px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>✕</button>
                            </>
                          ) : (
                            <>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 13, fontWeight: 700 }}>⭐ {cat}</div>
                                <div style={{ fontSize: 11, color: '#888' }}>{inits.length} initiative{inits.length !== 1 ? 's' : ''}</div>
                              </div>
                              <button onClick={() => { setEditCat(cat); setEditCatName(cat); }}
                                style={{ fontSize: 11, padding: '3px 10px', background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 3, cursor: 'pointer' }}>Rename</button>
                              <button onClick={() => retireCategory(cat)} disabled={spSaving}
                                style={{ background: 'none', border: 'none', color: '#d13438', fontSize: 16, cursor: 'pointer', lineHeight: 1, padding: '2px 4px' }}
                                title="Retire category">✕</button>
                            </>
                          )}
                        </div>
                        <div style={{ padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {inits.length === 0 && <div style={{ fontSize: 11, color: '#aaa', fontStyle: 'italic' }}>No initiatives yet.</div>}
                          {inits.map((init, idx) => {
                            const isEditingInit = !!editInit && editInit.cat === cat && editInit.idx === idx;
                            return (
                              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                {isEditingInit ? (
                                  <>
                                    <input value={editInitName} onChange={e => setEditInitName(e.target.value)} autoFocus
                                      onKeyDown={e => { if (e.key === 'Enter') commitInitRename(); }}
                                      style={{ flex: 1, fontSize: 12, padding: '4px 8px', border: '1px solid #0078d4', borderRadius: 4 }} />
                                    <button onClick={commitInitRename} disabled={spSaving}
                                      style={{ padding: '4px 10px', background: '#107c10', color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Save</button>
                                    <button onClick={() => setEditInit(null)}
                                      style={{ padding: '4px 8px', background: '#f3f2f1', color: '#323130', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, cursor: 'pointer' }}>✕</button>
                                  </>
                                ) : (
                                  <>
                                    <span style={{ flex: 1, fontSize: 12, padding: '3px 10px', background: '#eef4fa', border: '1px solid #dbeafe', borderRadius: 12, color: '#1e3a5f' }}>{init}</span>
                                    <button onClick={() => { setEditInit({ cat, idx }); setEditInitName(init); }}
                                      style={{ fontSize: 11, padding: '2px 8px', background: '#fff', color: HPE_NAVY, border: `1px solid ${HPE_NAVY}`, borderRadius: 3, cursor: 'pointer' }}>Rename</button>
                                    <button onClick={() => removeInitiative(cat, idx)} disabled={spSaving}
                                      style={{ background: 'none', border: 'none', color: '#d13438', fontSize: 15, cursor: 'pointer', lineHeight: 1, padding: '2px 4px' }}
                                      title="Remove initiative">✕</button>
                                  </>
                                )}
                              </div>
                            );
                          })}
                          <div style={{ display: 'flex', gap: 6, marginTop: 2 }}>
                            <input placeholder="+ Add initiative" value={newInit[cat] || ''}
                              onChange={e => setNewInit({ ...newInit, [cat]: e.target.value })}
                              onKeyDown={e => { if (e.key === 'Enter') addInitiative(cat); }}
                              style={{ flex: 1, fontSize: 12, padding: '5px 8px', border: '1px solid #ccc', borderRadius: 4 }} />
                            <button onClick={() => addInitiative(cat)} disabled={!(newInit[cat] || '').trim() || spSaving}
                              style={{ padding: '5px 12px', background: HPE_NAVY, color: '#fff', border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Add</button>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          )}

        </div>

        {/* Footer */}
        <div style={{ padding: '12px 20px', borderTop: '1px solid #edebe9', background: '#faf9f8' }}>
          <div style={{ fontSize: 11, color: '#888' }}>Logged in as <strong>{myEmail}</strong></div>
        </div>
      </div>
    </>
  );
};
