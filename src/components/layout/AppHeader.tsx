import { ReactNode } from "react";

interface Props {
  rightSlot?: ReactNode;
}

export function AppHeader({ rightSlot }: Props) {
  return (
    <header className="shell-header">
      <div className="brand">
        <div className="logo" />
        <div>
          <div style={{ fontWeight: 700, letterSpacing: 0.2 }}>Microsoft 365 SPA Blueprint</div>
          <div className="tag">React + Vite + MSAL + Graph</div>
        </div>
      </div>
      <div className="button-row">
        {rightSlot}
      </div>
    </header>
  );
}
