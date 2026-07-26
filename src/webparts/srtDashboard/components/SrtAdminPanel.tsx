import * as React from 'react';
import { useState } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ConfigService } from '../../../services/ConfigService';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISrtAdminPanelProps {
  sp: SPFI;
  context: WebPartContext;
  onClose: () => void;
}

export const SrtAdminPanel: React.FC<ISrtAdminPanelProps> = ({ sp, context, onClose }) => {
  const [superUsers, setSuperUsers] = useState<string[]>([]);
  const [newEmail, setNewEmail]     = useState('');
  const [loading, setLoading]       = useState(true);
  const [saving, setSaving]         = useState(false);
  const [message, setMessage]       = useState('');
  const [msgType, setMsgType]       = useState<'ok' | 'err'>('ok');

  const configSvc = React.useMemo(() => new ConfigService(sp), [sp]);
  const myEmail   = context.pageContext.user.email.toLowerCase();

  React.useEffect(() => {
    configSvc.getSuperUsers()
      .then(users => { setSuperUsers(users); setLoading(false); })
      .catch(() => setLoading(false));
  }, []);

  const save = async (users: string[]): Promise<void> => {
    setSaving(true);
    try {
      await configSvc.saveSuperUsers(users);
      setSuperUsers(users);
      setMessage('Saved.');
      setMsgType('ok');
    } catch (e) {
      setMessage(`Save failed: ${(e as Error).message}`);
      setMsgType('err');
    } finally { setSaving(false); }
  };

  const addUser = (): void => {
    const email = newEmail.trim().toLowerCase();
    if (!email || superUsers.includes(email)) return;
    save([...superUsers, email]).catch(() => undefined);
    setNewEmail('');
  };

  const removeUser = (email: string): void => {
    if (email === myEmail) {
      setMessage('You cannot remove yourself.');
      setMsgType('err');
      return;
    }
    save(superUsers.filter(u => u !== email)).catch(() => undefined);
  };

  return (
    <>
      {/* Backdrop */}
      <div
        onClick={onClose}
        style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.35)', zIndex: 1000 }}
      />

      {/* Panel */}
      <div style={{
        position: 'fixed', top: 0, right: 0, bottom: 0, width: 380,
        background: '#fff', boxShadow: '-4px 0 20px rgba(0,0,0,0.18)',
        zIndex: 1001, display: 'flex', flexDirection: 'column',
        fontFamily: 'inherit',
      }}>
        {/* Header */}
        <div style={{ background: HPE_NAVY, color: '#fff', padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 3, height: 20, background: HPE_GREEN, borderRadius: 2 }} />
            <div>
              <div style={{ fontSize: 15, fontWeight: 700 }}>SRT Admin</div>
              <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>Super User Management</div>
            </div>
          </div>
          <button
            onClick={onClose}
            style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer', lineHeight: 1, padding: 4 }}>
            ✕
          </button>
        </div>

        {/* Body */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '20px' }}>
          {message && (
            <div style={{
              marginBottom: 14, padding: '8px 12px', borderRadius: 4, fontSize: 12,
              background: msgType === 'ok' ? '#dff6dd' : '#fde7e9',
              color: msgType === 'ok' ? '#107c10' : '#a4262c',
              border: `1px solid ${msgType === 'ok' ? '#107c10' : '#a4262c'}`,
            }}>
              {message}
            </div>
          )}

          <p style={{ fontSize: 12, color: '#605e5c', marginTop: 0 }}>
            Super users see all SSE requests and can Accept, Decline, or request more info.
            Regular SEs see only their own submissions (read-only).
            Currently <strong>{superUsers.length}</strong> super user{superUsers.length !== 1 ? 's' : ''}.
          </p>

          {loading ? (
            <div style={{ color: '#888', fontSize: 13 }}>Loading…</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {superUsers.map(email => (
                <div key={email} style={{
                  display: 'flex', alignItems: 'center', gap: 10,
                  padding: '8px 12px', background: '#faf9f8',
                  border: '1px solid #edebe9', borderRadius: 4,
                }}>
                  <span style={{ flex: 1, fontSize: 13 }}>{email}</span>
                  {email === myEmail && (
                    <span style={{ fontSize: 11, color: '#888', fontStyle: 'italic' }}>(you)</span>
                  )}
                  <button
                    onClick={() => removeUser(email)}
                    disabled={saving}
                    style={{
                      background: 'none', border: 'none', color: '#d13438',
                      fontSize: 16, cursor: 'pointer', lineHeight: 1, padding: '2px 4px',
                    }}
                    title="Remove">
                    ✕
                  </button>
                </div>
              ))}

              {superUsers.length === 0 && (
                <div style={{ color: '#aaa', fontSize: 13, fontStyle: 'italic' }}>No super users yet.</div>
              )}
            </div>
          )}

          {/* Add user */}
          <div style={{ marginTop: 16, display: 'flex', gap: 8 }}>
            <input
              type="email"
              placeholder="user@hpe.com"
              value={newEmail}
              onChange={e => setNewEmail(e.target.value)}
              onKeyDown={e => { if (e.key === 'Enter') addUser(); }}
              style={{
                flex: 1, fontSize: 13, padding: '6px 10px',
                border: '1px solid #ccc', borderRadius: 4,
              }}
            />
            <button
              onClick={addUser}
              disabled={saving || !newEmail.trim()}
              style={{
                padding: '6px 16px', background: HPE_NAVY, color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 13, fontWeight: 600,
                cursor: saving || !newEmail.trim() ? 'not-allowed' : 'pointer',
                opacity: saving || !newEmail.trim() ? 0.6 : 1,
              }}>
              {saving ? '…' : 'Add'}
            </button>
          </div>
        </div>

        {/* Footer */}
        <div style={{ padding: '12px 20px', borderTop: '1px solid #edebe9', background: '#faf9f8' }}>
          <div style={{ fontSize: 11, color: '#888' }}>
            Logged in as <strong>{myEmail}</strong>
          </div>
        </div>
      </div>
    </>
  );
};
