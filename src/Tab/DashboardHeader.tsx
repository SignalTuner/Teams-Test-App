import React from "react";
import type { CurrentUser } from "./dashboardTypes";
import signalTunerDarkLogo from "./assets/signaltuner-logo-horizontal-darkmode.png";
import signalTunerLogo from "./assets/signaltuner-logo-horizontal.png";
export function SignalTunerLogo({ className }: { className: string }) {
  return (
    <span className={`${className} logoSwap`}>
      <img className="logoSwapLight" src={signalTunerLogo} alt="SignalTuner" />
      <img className="logoSwapDark" src={signalTunerDarkLogo} alt="SignalTuner" />
    </span>
  );
}

export function getInitials(name: string | null | undefined, email: string | null | undefined): string {
  const source = name?.trim() || email?.trim() || "SignalTuner User";
  const parts = source
    .replace(/@.*/, "")
    .split(/\s|[._-]/)
    .filter(Boolean);

  return parts
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join("") || "ST";
}

export default function DashboardHeader({ user, activePage, onNavigate, onSignOut }: {
 user: CurrentUser; activePage: "dashboard" | "settings" | "support";
 onNavigate: (page: "dashboard" | "settings" | "support") => void; onSignOut: () => void;
}) {
 const [accountOpen, setAccountOpen] = React.useState(false);
 const accountMenuRef = React.useRef<HTMLDivElement | null>(null);
  React.useEffect(() => {
    if (!accountOpen) {
      return;
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (!accountMenuRef.current?.contains(event.target as Node)) {
        setAccountOpen(false);
      }
    };

    window.addEventListener("pointerdown", handlePointerDown);
    return () => window.removeEventListener("pointerdown", handlePointerDown);
  }, [accountOpen]);

 return (
      <header className="appHeader">
        <SignalTunerLogo className="appLogo" />
        <div className="accountMenuWrap" ref={accountMenuRef}>
          <button
            aria-expanded={accountOpen}
            aria-haspopup="menu"
            className="accountButton"
            onClick={() => setAccountOpen((current) => !current)}
            type="button"
          >
            <span className="avatarBubble" aria-hidden="true">
              {getInitials(user.displayName, user.email)}
            </span>
            <span>{user.displayName ?? user.email ?? "SignalTuner account"}</span>
            <span className="chevron chevronDown" aria-hidden="true" />
          </button>
          {accountOpen && (
            <div className="accountDropdown" role="menu">
              <div className="accountDropdownHeader">
                <strong>{user.displayName ?? user.email ?? "SignalTuner account"}</strong>
                <span>{user.email}</span>
              </div>
              <button
                className={activePage === "dashboard" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("dashboard");
                  setAccountOpen(false);
                }}
              >
                Dashboard
              </button>
              <button
                className={activePage === "settings" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("settings");
                  setAccountOpen(false);
                }}
              >
                Settings
              </button>
              <button
                className={activePage === "support" ? "activeMenuItem" : ""}
                role="menuitem"
                type="button"
                onClick={() => {
                  onNavigate("support");
                  setAccountOpen(false);
                }}
              >
                Support
              </button>
              <button role="menuitem" type="button" onClick={onSignOut}>
                Sign out
              </button>
            </div>
          )}
        </div>
      </header>
 );
}
