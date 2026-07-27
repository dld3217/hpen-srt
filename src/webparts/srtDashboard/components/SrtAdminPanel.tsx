import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ConfigService, BURegionMap, IBUConfig } from '../../../services/ConfigService';
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
  const [tab, setTab]               = useState<'users' | 'bu'>('users');
  const [message, setMessage]       = useState('');
  const [msgType, setMsgType]       = useState<'ok' | 'err'>('ok');

  // ── Super Users ──────────────────────────────────────────────────────────────
  const [superUsers, setSuperUsers] = useState<string[]>([]);
  const [newEmail, setNewEmail]     = useState('');
  const [usersLoading, setUsersLoading] = useState(true);
  const [saving, setSaving]         = useState(false);

  // ── BU Config ─────────────────────────────────────────────────────────────
  const [buRegions, setBuRegions]   = useState<BURegionMap>({});
  const [buLoading, setBuLoading]   = useState(true);
  const [buSaving, setBuSaving]     = useState(false);
  const [editSedBu, setEditSedBu]   = useState<string | null>(null);
  const [editSedEmail, setEditSedEmail] = useState('');
  const [newBuName, setNewBuName]   = useState('');
  const [newBuSed, setNewBuSed]     = useState('');
  const [showAddBu, setShowAddBu]   = useState(false);

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
        </div>

        {/* Footer */}
        <div style={{ padding: '12px 20px', borderTop: '1px solid #edebe9', background: '#faf9f8' }}>
          <div style={{ fontSize: 11, color: '#888' }}>Logged in as <strong>{myEmail}</strong></div>
        </div>
      </div>
    </>
  );
};
