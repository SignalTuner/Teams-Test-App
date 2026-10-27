export type PollingState = { state: string; deliveries?: Array<{ state: string }> };
export const activeAnalysisStates = ["Pending", "Approved", "Running"];
export const analysisStatusPath = (id: string) => `/${encodeURIComponent(id)}?includeLive=true`;

// Same transport as the website: one request at a time, immediately then two
// seconds after each active reply. Only email delivery uses the slower cadence.
export function pollAnalysisRequest<T extends PollingState>({ read, onData, onError, onAccessLost,
  version, isMutating, visibility = document }: {
  read: (signal: AbortSignal) => Promise<T>;
  onData: (data: T) => void; onError: () => void; onAccessLost: () => void;
  version: () => number; isMutating: () => boolean;
  visibility?: Pick<Document, "hidden" | "addEventListener" | "removeEventListener">;
}) {
  let stopped = false;
  let fetching = false;
  let failures = 0;
  let interval = 2000;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let controller: AbortController | undefined;
  const schedule = (delay: number) => {
    if (timer !== undefined) { clearTimeout(timer); timer = undefined; }
    if (!stopped && !visibility.hidden) timer = setTimeout(() => { timer = undefined; void poll(); }, delay);
  };
  async function poll() {
    if (stopped || fetching || visibility.hidden) return;
    if (isMutating()) { schedule(interval); return; }
    fetching = true;
    const revision = version();
    controller = new AbortController();
    const timeout = setTimeout(() => controller?.abort(), 8000);
    let keepPolling = true;
    try {
      const data = await read(controller.signal);
      if (stopped || revision !== version() || visibility.hidden) return;
      failures = 0;
      const active = activeAnalysisStates.includes(data.state);
      interval = active ? 2000 : 10000;
      keepPolling = active || Boolean(data.deliveries?.some(d => ["Queued", "Sending"].includes(d.state)));
      onData(data);
    } catch (error) {
      if (stopped || revision !== version() || visibility.hidden) return;
      const status = (error as { status?: number })?.status;
      if (status && [401, 403, 404].includes(status)) {
        stopped = true;
        onAccessLost();
      } else { failures++; onError(); }
    } finally {
      clearTimeout(timeout);
      fetching = false;
      if (keepPolling) schedule(Math.min(30000, interval * 2 ** failures));
      else stopped = true;
    }
  }
  const changed = () => {
    if (timer !== undefined) { clearTimeout(timer); timer = undefined; }
    if (visibility.hidden) controller?.abort();
    else void poll();
  };
  visibility.addEventListener("visibilitychange", changed);
  void poll();
  return () => {
    stopped = true;
    if (timer !== undefined) { clearTimeout(timer); timer = undefined; }
    controller?.abort();
    visibility.removeEventListener("visibilitychange", changed);
  };
}
