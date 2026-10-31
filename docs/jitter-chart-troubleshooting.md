# Jitter chart shows only the latest measurement

The Teams dashboard uses the same chart component in a browser and inside Teams.
The deployed frontend checked on 2026-10-27 supports `jitterHistoryMs`. A scalar-only
response reproduces the reported single-point chart; this does not establish which
deployed upstream component is missing history support.

Resolution: on 2026-10-27, the user confirmed that an older build had been running
and that the chart now works. No application-code change or deployment was performed
as part of this investigation.

During a fresh Test Analysis, inspect the response to the existing authenticated
`GET /api/remote-analysis/{id}?includeLive=true` request in browser developer tools.
Check only `liveTelemetry.snapshot.jitterHistoryMs`, `jitterMs` and
`elapsedSeconds`; do not copy session headers, request IDs, device IDs or complete
responses into logs or issue reports.

- A history array with multiple valid numbers should render multiple points.
  Zero is valid. Null entries are gaps and are not joined across.
- An absent/null history array with a valid `jitterMs` renders one latest point.
  Repeated polling cannot recover the missing probe history.
- An explicit empty array renders no samples. An array with only one valid number
  renders one point, even if its other entries are null.

The Windows client owns measurement history. Current source uploads its latest ten
probe slots, and the backend validates, stores and returns that array. An older
desktop supplies only the scalar; an older backend can discard the history even
when the desktop sends it. See the backend's
[live telemetry contract and rollout](../../signaltuner-back/docs/remote-analysis-live.md#jitter-history-deployment-impact-additive).

If the live response lacks history, verify that every backend instance has the
current implementation, then rebuild/release the current Windows client. Fully exit
the old client, including its tray process, start the updated build and begin a new
Test Analysis. Deployment is a separate authorized operation. The backend target is
the existing Azure App Service `signaltuner-back` (Azure portal > App Services >
signaltuner-back > Deployment Center for deployment status); this fix requires no
new App Service setting, database migration, permission or Teams package. Verify
the actual deployment target/slot before publishing. Retain the scalar for older
clients; do not derive jitter from latency values that may represent timeouts.

Validation: `npm run test:analysis-live` covers windows, colors, zeros and gaps.
`npm run test:dreamhost-browser` includes a compiled-dashboard regression with
synthetic history updates, reload recovery and the scalar-only reproduction. These
tests do not confirm the contents of a live authenticated telemetry response.
