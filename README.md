# BigFix Action Reporter ⚡

A PowerShell/WPF dashboard for visualizing BigFix action deployment status.

## Features
- 🍩 **Donut chart** — Fixed/Failed/Running/Pending/Not Relevant/Expired breakdown
- 📈 **S-curve timeline** — Cumulative completions over time
- 📋 **Sortable endpoint table** — Computer name, status, timestamps, apply/retry counts
- 📄 **CSV export** — One-click export for management reports
- 🔄 **Live refresh** — Re-pull status without re-entering the action ID
- 🎨 **Dark theme** — Catppuccin Mocha, easy on the eyes

## Requirements
- Windows PowerShell 5.1+ or PowerShell 7+
- .NET Framework (for WPF) — built into Windows
- BigFix server with REST API enabled (port 52311)
- Credentials with API access

## Usage
```powershell
# Just run it
.\BigFixActionReporter.ps1
```

1. Enter your BigFix server URL (e.g. `https://bigfix-server:52311`)
2. Enter credentials
3. Click **Connect**
4. Enter an Action ID and click **📊 Fetch Status**

## API Endpoints Used
- `GET /api/action/{id}` — Action name/details
- `GET /api/action/{id}/status` — Per-computer results
- `GET /api/login` — Connection test

## Notes
- Self-signed certs are handled automatically (common in BigFix deployments)
- Completion % excludes "Not Relevant" endpoints from the denominator
- Timeline chart needs 2+ completed endpoints to render the S-curve
