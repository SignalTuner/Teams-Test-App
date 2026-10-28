console.error("SignalTuner production frontend uses DreamHost static hosting. This Toolkit action is disabled: it must not provision Azure, create/change app IDs, or generate/publish Teams packages. Run npm run build:dreamhost and follow docs/dreamhost-hosting.md. Local Teams debugging uses m365agents.local.yml and the existing VS Code tasks.");
process.exitCode = 1;
