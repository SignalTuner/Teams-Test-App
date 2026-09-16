# Windows RSSI telemetry compatibility

The Windows client now sends SignalWifiStrength as an integer percentage derived
only from numeric connected-BSSID RSSI: round((RSSI + 90) * 2.5), midpoint rounding
away from zero, clamped to 0?100. Wi-Fi tiers are Healthy 61?100, Fair 46?60,
Poor 31?45, Critical 0?30. The top telemetry status is now Healthy. Overall score
thresholds stay 90/80/70 and Wi-Fi retains its maximum 35-point contribution.

Accept legacy Excellent values from existing clients and historical data. Prefer
producer Wi-Fi status when rendering: older percentages use a different scale.
A zero strength with Unknown/Disconnected status is unavailable, not a measured
Critical reading. Ethernet retains 100%. No API field names or database schema
change; no migration or historical rewrite is required.

Deploy backend compatibility first, website and Teams consumers next, then the
Windows client. All four projects require release; old consumers may miscolor new
values and an old backend may treat Healthy as an issue. Roll back Windows first;
keep compatibility handling while Healthy records remain. Test mixed old/new
statuses, real 0%, unavailable readings, and Ethernet after deployment.
No production deployment or new logging of personal data is included.
