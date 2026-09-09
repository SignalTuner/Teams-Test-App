import React from "react";

export const ActivationContext = React.createContext({ apiBaseUrl: "", token: "" });

type CodeState = { activationCode: string | null; expiresAtUtc: string | null; serverNowUtc?: string; expiresAtLocal?: number; sessionToken?: string };

export default function TemporaryActivationCode({ onCopied }: { onCopied?: () => void }) {
  const { apiBaseUrl, token } = React.useContext(ActivationContext);
  const [code, setCode] = React.useState<CodeState | null>(null);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState("");
  const [copied, setCopied] = React.useState(false);
  const [now, setNow] = React.useState(Date.now());
  const requestRef = React.useRef<AbortController | null>(null);
  const generatingRef = React.useRef(false);

  const refresh = React.useCallback(async (generate = false) => {
    if (!token || (!generate && requestRef.current)) return;
    requestRef.current?.abort();
    const controller = new AbortController();
    requestRef.current = controller;
    try {
      const response = await fetch(`${apiBaseUrl.replace(/\/$/, "")}/api/User/activation-code`, {
        method: generate ? "POST" : "GET",
        headers: { Authorization: `Bearer ${token}` },
        cache: "no-store",
        signal: controller.signal,
      });
      if (!response.ok) throw new Error(response.status === 429
        ? "Too many attempts. Please wait a minute and try again."
        : "Unable to check your activation code. Please try again.");
      const result = await response.json() as CodeState;
      if (!controller.signal.aborted) {
        const serverOffset = result.serverNowUtc ? Date.parse(result.serverNowUtc) - Date.now() : 0;
        setCode({ ...result, sessionToken: token, expiresAtLocal: Date.parse(result.expiresAtUtc ?? "") - serverOffset });
        setNow(Date.now());
        setError("");
      }
    } catch (caught) {
      if (!controller.signal.aborted) {
        setCode(null);
        setError(caught instanceof Error ? caught.message : "Unable to check your activation code.");
      }
    } finally {
      if (requestRef.current === controller) requestRef.current = null;
    }
  }, [apiBaseUrl, token]);

  React.useEffect(() => {
    setCode(null);
    setCopied(false);
    void refresh();
    const poll = window.setInterval(() => void refresh(), 3000);
    const clock = window.setInterval(() => setNow(Date.now()), 250);
    const onFocus = () => void refresh();
    window.addEventListener("focus", onFocus);
    document.addEventListener("visibilitychange", onFocus);
    return () => {
      window.clearInterval(poll);
      window.clearInterval(clock);
      window.removeEventListener("focus", onFocus);
      document.removeEventListener("visibilitychange", onFocus);
      requestRef.current?.abort();
      requestRef.current = null;
    };
  }, [refresh]);

  const remaining = code?.expiresAtLocal ? Math.max(0, Math.ceil((code.expiresAtLocal - now) / 1000)) : 0;
  const activeCode = remaining > 0 && code?.sessionToken === token ? code.activationCode : null;
  React.useEffect(() => setCopied(false), [activeCode]);
  const generate = async () => {
    if (generatingRef.current) return;
    generatingRef.current = true;
    setBusy(true);
    setCopied(false);
    await refresh(true);
    generatingRef.current = false;
    setBusy(false);
  };
  const copy = async () => {
    if (!activeCode || (code?.expiresAtLocal ?? 0) <= Date.now()) return;
    try {
      await navigator.clipboard.writeText(activeCode);
      setCopied(true);
      onCopied?.();
    } catch { setError("Could not copy. Select and copy the code below."); }
  };

  return <div className="temporaryActivationCode">
    {activeCode ? <>
      <strong className="temporaryActivationCodeValue">{activeCode}</strong>
      <button type="button" onClick={() => void copy()}>{copied ? "Copied" : "Copy code"}</button>
      <p>Valid for 5 minutes. Single use. {Math.floor(remaining / 60)}:{String(remaining % 60).padStart(2, "0")} remaining.</p>
    </> : <button type="button" disabled={busy || !token} onClick={() => void generate()}>
      {busy ? "Generating…" : "Generate activation code"}
    </button>}
    <p>Enter the code in the latest SignalTuner desktop client.</p>
    {error && <p role="alert">{error}</p>}
  </div>;
}
