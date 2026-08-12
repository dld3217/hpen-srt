import * as React from 'react';
import { useState } from 'react';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { StrategicEngagementForm } from './StrategicEngagementForm';
import { HPE_GREEN, HPE_NAVY } from '../../../styles/hpe';

export interface ISseFrontDoorProps {
  sp: SPFI;
  context: WebPartContext;
}

const POC_HOME = 'https://hpe.sharepoint.com/teams/hpen-poc-manager/SitePages/Home.aspx';

const firstName = (displayName: string): string => {
  if (!displayName) return '';
  if (displayName.includes(',')) return displayName.split(',')[1].trim().split(' ')[0];
  return displayName.split(' ')[0];
};
const greeting = (): string => {
  const h = new Date().getHours();
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
};

interface ICard { key: string; icon: string; title: string; desc: string; cta: string; }
const CARDS: ICard[] = [
  { key: 'new',       icon: '▶',  title: 'Start a New POC',        desc: 'Kick off a proof-of-concept in POC Manager — solutions, test plans, tracking.', cta: 'Open POC Manager →' },
  { key: 'modify',    icon: '✎',  title: 'Modify an Existing POC', desc: 'Jump into POC Manager and update one of your active POCs.',                    cta: 'Open My POCs →' },
  { key: 'strategic', icon: '🎯', title: 'Strategic Engagement',   desc: 'Request SSE time for a briefing, roadmap talk, or advisory — outside an active POC.', cta: 'Start request →' },
];

export const SseFrontDoor: React.FC<ISseFrontDoorProps> = ({ sp, context }) => {
  const [mode, setMode] = useState<'router' | 'strategic'>('router');
  const userDisplayName = context.pageContext.user.displayName || context.pageContext.user.email;

  if (mode === 'strategic') {
    return <StrategicEngagementForm sp={sp} context={context} onBack={() => setMode('router')} />;
  }

  const go = (key: string): void => {
    if (key === 'new') window.location.href = `${POC_HOME}?new=1`;
    else if (key === 'modify') window.location.href = POC_HOME;
    else setMode('strategic');
  };

  return (
    <div style={{ maxWidth: 820, margin: '0 auto', padding: '20px 16px', fontFamily: 'inherit' }}>
      {/* Header */}
      <div style={{ background: HPE_NAVY, color: '#fff', padding: '10px 18px', borderRadius: '6px 6px 0 0', marginBottom: 18, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 3, height: 24, background: HPE_GREEN, borderRadius: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, lineHeight: 1.2 }}>SSE Support Request</div>
            <div style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)' }}>HPE Networking — Strategic Systems Engineer</div>
          </div>
        </div>
        <div style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)', whiteSpace: 'nowrap' }}>{greeting()}, {firstName(userDisplayName)}</div>
      </div>

      <div style={{ fontSize: 14, color: '#605e5c', marginBottom: 16, textAlign: 'center' }}>What do you need to do?</div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 14 }}>
        {CARDS.map(c => (
          <button key={c.key} type="button" onClick={() => go(c.key)}
            style={{ textAlign: 'left', background: '#fff', border: '1px solid #e0e0e0', borderRadius: 8, padding: '20px 18px', cursor: 'pointer', display: 'flex', flexDirection: 'column', gap: 8, boxShadow: '0 1px 3px rgba(0,0,0,0.06)', transition: 'all .15s' }}
            onMouseEnter={e => { e.currentTarget.style.borderColor = HPE_GREEN; e.currentTarget.style.boxShadow = '0 4px 14px rgba(0,0,0,0.12)'; e.currentTarget.style.transform = 'translateY(-2px)'; }}
            onMouseLeave={e => { e.currentTarget.style.borderColor = '#e0e0e0'; e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.06)'; e.currentTarget.style.transform = 'none'; }}>
            <div style={{ fontSize: 26 }}>{c.icon}</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: HPE_NAVY }}>{c.title}</div>
            <div style={{ fontSize: 12.5, color: '#605e5c', lineHeight: 1.5 }}>{c.desc}</div>
            <div style={{ marginTop: 'auto', paddingTop: 8, fontSize: 12, fontWeight: 700, color: HPE_GREEN }}>{c.cta}</div>
          </button>
        ))}
      </div>
    </div>
  );
};
