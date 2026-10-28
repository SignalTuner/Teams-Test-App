import React, { createContext, useContext, useEffect, useState } from "react";
import { allowsExternalCommerce, initializeTeams } from "./teamsHost";

const CommerceContext = createContext(false);
export const useExternalCommerce = () => useContext(CommerceContext);

export function TeamsHostProvider({ children }: { children: React.ReactNode }) {
  const [allowed, setAllowed] = useState(false);
  useEffect(() => {
    let active = true;
    void initializeTeams().then(context => { if (active) setAllowed(allowsExternalCommerce(context)); })
      .catch(() => { if (active) setAllowed(false); });
    return () => { active = false; };
  }, []);
  return <CommerceContext.Provider value={allowed}>{children}</CommerceContext.Provider>;
}

// These destinations can expose commerce directly or through their navigation.
export function ExternalWebsiteLink({ href, children, className }: {
  href: string; children: string; className?: string;
}) {
  const allowed = useExternalCommerce();
  if (!allowed) return null;
  return <a href={href} className={className} target="_blank" rel="noopener noreferrer"
    aria-label={`${children} (opens a website outside Microsoft Teams)`}
    title="Opens a website outside Microsoft Teams">
    {children} <span aria-hidden="true">↗</span> <span>(outside Teams)</span>
  </a>;
}

export function PaymentDisclosure({ className = "authNotice" }: { className?: string }) {
  const allowed = useExternalCommerce();
  return allowed ? <p className={className} data-payment-disclosure>
    Purchase links open SignalTuner's website outside Microsoft Teams. Checkout and billing
    use Stripe, where you enter payment details. Review the price and terms there before confirming a purchase.
  </p> : null;
}
