# Temporary activation controls

Client setup and account settings use `TemporaryActivationCode.tsx` with the
current SignalTuner bearer session from `ActivationContext`. They explicitly POST
to generate a code and GET status every three seconds. The code is single-use,
valid for five minutes, displayed with a countdown, and cleared after expiry,
redemption, or a failed status check. Merely loading account/dashboard data no
longer controls code generation or display.

See [the authoritative backend contract](../../signaltuner-back/docs/temporary-activation-codes.md)
for the migration, endpoint shapes, compatibility and verification checklist.
Deploy the database migration and backend, release the compatible desktop, then
deploy this frontend and the website. Existing manifests and Teams/Entra
configuration do not change. Manual Teams tenant validation remains required.

Run `npm run build:frontend` and `node_modules/.bin/tsc -p tsconfig.app.json`.
The production frontend build passes. Strict TypeScript validation currently
reports the pre-existing unused `getSignalTone` function in `App.tsx`; no activation
type errors remain.
