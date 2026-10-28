import React, { useEffect, useRef, useState } from "react";
import AppearanceSelector from "./AppearanceSelector";
import { loadReportAccount, periodDate, subscriptionStatus, type ReportAccount } from "./settingsAccount";
import "./Settings.css";
import { ExternalWebsiteLink, PaymentDisclosure, useExternalCommerce } from "./TeamsHostProvider";

type SettingsUser = { userId: number; firstName?: string | null; lastName?: string | null; email: string | null; hasPassword?: boolean | null };
type Feedback = { text: string; error?: boolean } | null;
const BILLING_PORTAL = "https://billing.stripe.com/p/login/8wMdUQfLt1PjdpK5kk";

export default function SettingsPage({ user, apiBaseUrl, token, themePreference, onThemePreferenceChange,
  onUpdateProfile, onUpdatePassword, validatePassword, onOpenSupport }: {
  user: SettingsUser; apiBaseUrl: string; token: string; themePreference: "light" | "dark";
  onThemePreferenceChange: (value: "light" | "dark") => void;
  onUpdateProfile: (firstName: string, lastName: string, email: string) => Promise<void>;
  onUpdatePassword: (password: string) => Promise<string>;
  validatePassword: (password: string) => string | undefined;
  onOpenSupport: () => void;
}) {
  const allowCommerce = useExternalCommerce();
  const [account, setAccount] = useState<ReportAccount | null>(null);
  const [accountError, setAccountError] = useState("");
  const [loading, setLoading] = useState(true);
  const [attempt, setAttempt] = useState(0);
  const [userFirstName, setUserFirstName] = useState(user.firstName || "");
  const [userLastName, setUserLastName] = useState(user.lastName || "");
  const [userEmail, setUserEmail] = useState(user.email || "");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [profileMessage, setProfileMessage] = useState<Feedback>(null);
  const [passwordMessage, setPasswordMessage] = useState<Feedback>(null);
  const [savingProfile, setSavingProfile] = useState(false);
  const [savingPassword, setSavingPassword] = useState(false);
  const profilePending = useRef(false);
  const passwordPending = useRef(false);

  useEffect(() => {
    setUserFirstName(user.firstName || ""); setUserLastName(user.lastName || ""); setUserEmail(user.email || "");
  }, [user.userId, user.firstName, user.lastName, user.email]);

  useEffect(() => {
    const controller = new AbortController();
    let current = true;
    const timeout = setTimeout(() => controller.abort(), 20000);
    setLoading(true); setAccountError(""); setAccount(null);
    loadReportAccount(apiBaseUrl, token, controller.signal)
      .then(data => { if (current) setAccount(data); })
      .catch(error => { if (current) setAccountError(controller.signal.aborted
        ? "Your account took too long to load. Please try again."
        : error instanceof Error && ["Your session has expired. Sign in again to view your account.", "Your account could not be loaded. Please try again."].includes(error.message)
          ? error.message : "Your account could not be loaded. Please try again."); })
      .finally(() => { clearTimeout(timeout); if (current) setLoading(false); });
    return () => { current = false; clearTimeout(timeout); controller.abort(); };
  }, [apiBaseUrl, token, user.userId, attempt]);

  const updateUser = async (event: React.FormEvent) => {
    event.preventDefault();
    if (profilePending.current) return;
    setProfileMessage(null);
    const firstName = userFirstName.trim(), lastName = userLastName.trim(), email = userEmail.trim();
    if (!firstName || !lastName) { setProfileMessage({ error: true, text: "Enter your first and last name." }); return; }
    if (!user.hasPassword && email.toLowerCase() !== (user.email || "").toLowerCase()) {
      setProfileMessage({ error: true, text: "Create a password before updating your email." }); return;
    }
    profilePending.current = true; setSavingProfile(true);
    try {
      await onUpdateProfile(firstName, lastName, email);
      setProfileMessage({ text: "Profile saved." });
    } catch (error) {
      setProfileMessage({ error: true, text: error instanceof Error ? error.message : "Your profile could not be saved. Please try again." });
    } finally { profilePending.current = false; setSavingProfile(false); }
  };

  const updatePassword = async (event: React.FormEvent) => {
    event.preventDefault();
    if (passwordPending.current) return;
    setPasswordMessage(null);
    if (newPassword !== confirmPassword) { setPasswordMessage({ error: true, text: "Passwords do not match." }); return; }
    const error = validatePassword(newPassword);
    if (error) { setPasswordMessage({ error: true, text: error }); return; }
    passwordPending.current = true; setSavingPassword(true);
    try {
      await onUpdatePassword(newPassword);
      setNewPassword(""); setConfirmPassword(""); setPasswordMessage({ text: "Password updated." });
    } catch {
      setPasswordMessage({ error: true, text: "Your password could not be updated. Please try again." });
    } finally { passwordPending.current = false; setSavingPassword(false); }
  };

  const subscription = account?.subscription;
  const renewalDate = periodDate(subscription?.renewalDate);
  const monthlyActive = subscription?.status === "active";
  const message = (value: Feedback) => value && <p className={`settings-message${value.error ? " settings-message--error" : ""}`} role={value.error ? "alert" : "status"}>{value.text}</p>;

  return <section className="settings-page" aria-labelledby="settings-heading">
    <div className="settings-inner">
      <header className="settings-header settings-header--appearance">
        <div className="settings-header-copy">
          <p className="settings-eyebrow">YOUR ACCOUNT</p>
          <h1 id="settings-heading">Settings</h1>
          <p>Your reports, your details, your preferences.</p>
        </div>
        <AppearanceSelector value={themePreference} onChange={onThemePreferenceChange} />
      </header>

      <section className="settings-card settings-account" aria-labelledby="settings-account-heading" aria-busy={loading}>
        <div className="settings-section-heading">
          <div><h2 id="settings-account-heading">Account overview</h2><p>Everything you need for your next Connectivity Health Report.</p></div>
          <button type="button" className="settings-button settings-button--quiet" disabled={loading} onClick={() => setAttempt(value => value + 1)}>{loading ? "Refreshing..." : "Refresh balance"}</button>
        </div>
        {loading && <p role="status" className="settings-loading">Loading your report allowance...</p>}
        {accountError && <div className="settings-account-error"><p role="alert">{accountError}</p><button type="button" className="settings-button" onClick={() => setAttempt(value => value + 1)}>Try again</button></div>}
        {account && <>
          <div className="settings-account-grid">
            <div className="settings-credit-summary">
              <p className="settings-metric-label">{allowCommerce ? "Credit balance" : "Available report allowance"}</p>
              <p className="settings-credit-value">{(allowCommerce ? account.purchasedCredits : account.availableReports).toLocaleString()} <span>{allowCommerce ? "credits" : "reports"}</span></p>
              <p>Each report uses one unit of your report allowance.</p>
              <p>{account.teamAllocatedCredits || 0} allocated team credits &middot; {account.teamFundAvailable || 0} general fund credits available</p>
              {account.availableReports === 0 && <p>No report allowance is available. Test Analysis remains available on supported connected desktops.</p>}
              {allowCommerce && <div className="settings-actions">
                <ExternalWebsiteLink className="settings-button" href="https://www.signaltuner.com/team">Manage team credits on website</ExternalWebsiteLink>
                <ExternalWebsiteLink className="settings-button settings-button--primary settings-add-credits" href="https://www.signaltuner.com/plans">Buy credits on website</ExternalWebsiteLink>
              </div>}
            </div>
            {allowCommerce && <div className="settings-plan-summary">
              <p className="settings-metric-label">Your plan</p>
              <div className="settings-plan-title"><h3>{subscription?.name || "Pay as you go"}</h3>{subscription && <span className="settings-plan-status">{subscriptionStatus(subscription)}</span>}</div>
              {subscription ? <>
                <p><strong>{subscription.remaining} of {subscription.allowance}</strong> monthly reports available</p>
                {monthlyActive && !subscription.paidPeriod && <p className="settings-plan-note">Your paid report allowance is awaiting confirmation. Refresh your balance shortly.</p>}
                {monthlyActive && renewalDate && <p>{subscription.cancelAtPeriodEnd ? "Plan ends" : "Renews"} {renewalDate}</p>}
                {!monthlyActive && <p>Monthly reports are unavailable. Your credits are still available for reports.</p>}
              </> : <p>Get credits when you need them, or choose a monthly report allowance.</p>}
              <div className="settings-actions">
                <ExternalWebsiteLink className="settings-button" href="https://www.signaltuner.com/plans">View plans on website</ExternalWebsiteLink>
                {account.isCustomer && <ExternalWebsiteLink className="settings-billing-link" href={BILLING_PORTAL}>Manage billing on Stripe</ExternalWebsiteLink>}
              </div>
            </div>}
          </div>
          <PaymentDisclosure className="settings-field-help" />
          <div className="settings-account-footer"><span><strong>{account.availableReports}</strong> reports available in total</span><span>Team credits are used first, then your monthly allowance, then personal credits.</span></div>
        </>}
      </section>

      <div className="settings-details-grid">
        <section className="settings-card" aria-labelledby="settings-profile-heading">
          <h2 id="settings-profile-heading">Profile Information</h2>
          <p>Keep your name and email up to date.</p>
          <form className="settings-form" onSubmit={updateUser}>
            <fieldset disabled={savingProfile}>
              <div className="settings-name-fields">
                <div><label htmlFor="settings-first-name">First name</label><input id="settings-first-name" autoComplete="given-name" maxLength={255} required value={userFirstName} onChange={event => setUserFirstName(event.target.value)} /></div>
                <div><label htmlFor="settings-last-name">Last name</label><input id="settings-last-name" autoComplete="family-name" maxLength={255} required value={userLastName} onChange={event => setUserLastName(event.target.value)} /></div>
              </div>
              <label htmlFor="settings-email">Email address</label><input id="settings-email" readOnly={!user.hasPassword} aria-describedby={!user.hasPassword ? "settings-email-help" : undefined} type="email" autoComplete="email" required value={userEmail} onChange={event => setUserEmail(event.target.value)} />
              {!user.hasPassword && <p id="settings-email-help" className="settings-field-help">Create a password before updating your email.</p>}
              <button className="settings-button settings-button--primary" type="submit">{savingProfile ? "Saving..." : "Save profile"}</button>
            </fieldset>
            {message(profileMessage)}
          </form>
        </section>
        <section className="settings-card" aria-labelledby="settings-password-heading">
          <h2 id="settings-password-heading">{user.hasPassword ? "Change password" : "Create password"}</h2>
          <p>Choose a strong password to protect your account.</p>
          <form className="settings-form" onSubmit={updatePassword}>
            <fieldset disabled={savingPassword}>
              <label htmlFor="settings-new-password">New password</label><input id="settings-new-password" type="password" autoComplete="new-password" required aria-describedby="settings-password-help" value={newPassword} onChange={event => setNewPassword(event.target.value)} />
              <label htmlFor="settings-confirm-password">Confirm new password</label><input id="settings-confirm-password" type="password" autoComplete="new-password" required value={confirmPassword} onChange={event => setConfirmPassword(event.target.value)} />
              <p id="settings-password-help" className="settings-field-help">At least 8 characters with uppercase and lowercase letters, a number, and a symbol.</p>
              <button className="settings-button" type="submit">{savingPassword ? "Updating..." : user.hasPassword ? "Update password" : "Create password"}</button>
            </fieldset>
            {message(passwordMessage)}
          </form>
        </section>
      </div>
      <section className="settings-card settings-delete-account" aria-labelledby="settings-delete-account-heading">
        <h2 id="settings-delete-account-heading">Delete account</h2>
        {allowCommerce ? <>
          <p>Permanently remove your account. Review the details and confirm on the website outside Microsoft Teams.</p>
          <ExternalWebsiteLink className="settings-button" href="https://www.signaltuner.com/account/delete">Delete my account on website</ExternalWebsiteLink>
        </> : <>
          <p>To request permanent account deletion, contact SignalTuner support.</p>
          <button type="button" className="settings-button" onClick={onOpenSupport}>Contact support about account deletion</button>
        </>}
      </section>
    </div>
  </section>;
}
