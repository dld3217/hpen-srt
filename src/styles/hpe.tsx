import * as React from 'react';
import { IDropdownStyles, ICheckboxStyles } from '@fluentui/react';

export const HPE_NAVY  = '#13294B';
export const HPE_GREEN = '#01A982';
export const HPE_BORDER = '#d0d8d4';

const HPE_LOGO_B64 = 'data:image/png;base64,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';

export const HpeAppHeader: React.FC<{ title?: string; version?: string; greeting?: string }> = ({
  title = 'SSE Support Request Tracker', version, greeting
}) => (
  <div style={{
    background: HPE_NAVY, padding: '8px 16px',
    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
  }}>
    <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
      <img src={HPE_LOGO_B64} alt="HPE" style={{ height: 24 }} />
      <span style={{ color: '#fff', fontWeight: 300, fontSize: 15, letterSpacing: '0.3px' }}>{title}</span>
    </div>
    <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
      {greeting && (
        <span style={{ color: '#fff', fontSize: 18, fontWeight: 400, whiteSpace: 'nowrap', flexShrink: 0 }}>
          {greeting}
        </span>
      )}
      {version && <span style={{ color: 'rgba(255,255,255,0.75)', fontSize: 14, fontWeight: 500 }}>v{version}</span>}
    </div>
  </div>
);

export const SectionHeader: React.FC<{ title: string }> = ({ title }) => (
  <div style={{
    background: HPE_GREEN,
    color: '#fff',
    fontWeight: 700,
    fontSize: 14,
    padding: '5px 8px',
    textTransform: 'uppercase',
    letterSpacing: '0.6px',
    marginBottom: 8,
  } as React.CSSProperties}>
    {title}
  </div>
);

export const FieldRow: React.FC<{
  label: string;
  required?: boolean;
  wide?: boolean;
  narrow?: boolean;
  children: React.ReactNode;
}> = ({ label, required, wide, narrow, children }) => (
  <div style={{
    display: 'flex', alignItems: 'center', minHeight: 26, marginBottom: 4, gap: 6,
    background: required ? '#fffbe6' : undefined,
    borderRadius: required ? 3 : undefined,
    padding: required ? '1px 4px' : undefined,
  }}>
    <label style={{
      whiteSpace: 'nowrap',
      fontWeight: 700,
      color: HPE_NAVY,
      fontSize: 14,
      minWidth: narrow ? 130 : wide ? 210 : 165,
      maxWidth: narrow ? 130 : wide ? 210 : 165,
      flexShrink: 0,
    } as React.CSSProperties}>
      {label}{required && <span style={{ color: '#c00', marginLeft: 2 }}>*</span>}
    </label>
    <div style={{ flex: 1, minWidth: 0 }}>{children}</div>
  </div>
);

export const HpeDivider: React.FC = () => (
  <hr style={{ border: 'none', borderTop: `1px solid ${HPE_BORDER}`, margin: '6px 0' }} />
);

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
export const hpeTfStyles = (labelWidth = 165) => ({
  root: { marginBottom: 4 },
  subComponentStyles: {
    label: {
      root: {
        color: HPE_NAVY,
        fontWeight: 700,
        fontSize: 14,
        minWidth: labelWidth,
        maxWidth: labelWidth,
        lineHeight: '28px',
        padding: '0 4px 0 0',
        whiteSpace: 'nowrap',
      }
    }
  },
  field: { fontSize: 14, padding: '1px 2px' },
  fieldGroup: {
    selectors: { '::after': { borderColor: HPE_GREEN } }
  },
});

export const hpeMultilineStyles = {
  root: { marginBottom: 8 },
  subComponentStyles: {
    label: { root: { color: HPE_NAVY, fontWeight: 700, fontSize: 14 } }
  },
  field: { fontSize: 14 },
  fieldGroup: {
    border: `1px solid ${HPE_BORDER}`,
    selectors: { '::after': { borderColor: HPE_GREEN } }
  }
};

export const hpeDropStyles: Partial<IDropdownStyles> = {
  root: { width: '100%' },
  title: {
    border: 'none',
    borderBottom: `1px solid #aaa`,
    borderRadius: 0,
    height: 28,
    minHeight: 28,
    lineHeight: '28px',
    fontSize: 14,
    padding: '0 28px 0 2px',
  },
  caretDownWrapper: { height: 28, lineHeight: '28px', top: 0 },
  dropdownItem: { fontSize: 14, minHeight: 30 },
  dropdownItemSelected: { fontSize: 14, minHeight: 30 },
};

export const hpeCbStyles: Partial<ICheckboxStyles> = {
  root: { marginBottom: 4 },
  checkbox: { width: 16, height: 16 },
  text: { fontSize: 14, color: '#333' },
};
