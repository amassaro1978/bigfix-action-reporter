# BigFix Action Reporter ⚡

A PowerShell/WPF dashboard for visualizing BigFix action deployment status with single-action lookups and automated weekly reports.

## Features
- 🍩 **Donut chart** — Fixed/Failed/Running/Pending/Not Relevant/Expired breakdown
- 📈 **S-curve timeline** — Cumulative completions over time with Fixed/Failed toggle
- 📊 **Completion gauge** — Large percentage display with color-coded status pills
- 📋 **Sortable endpoint table** — Computer name, status, timestamps, apply/retry counts
- 📑 **Multi-action tabs** — Enter comma-separated Action IDs to compare side-by-side
- 📆 **Weekly Report** — Automatically discovers and aggregates all deployment actions from a site over a date range
- 📄 **CSV export** — One-click export for management reports
- 🔄 **Live refresh** — Re-pull status without re-entering the action ID
- 🎨 **Dark theme** — Catppuccin Mocha, easy on the eyes
- 📝 **CMTrace logging** — Full CMTrace-compatible log at `C:\temp\BigFixActionReporter.log`

## Requirements
- Windows PowerShell 5.1+ or PowerShell 7+
- .NET Framework (for WPF and System.Windows.Forms.DataVisualization) — built into Windows
- BigFix server with REST API enabled (port 52311)
- Credentials with API access

## Usage

```powershell
# Just run it
.\BigFixActionReporter.ps1
```

1. Enter your BigFix server URL (e.g. `https://bigfix-server:52311`)
2. Enter credentials and click **Connect**
3. Enter one or more Action IDs (comma-separated) and click **Fetch Status**
4. Or click **Weekly Report** to generate an automated site-wide deployment summary

## Single Action Lookup

Enter an Action ID to get a full dashboard with:
- Status donut chart showing the distribution across all endpoints
- Completion rate gauge (excludes "Not Relevant" from the denominator)
- Timeline chart showing the S-curve of completions over time
- Full endpoint detail table with sorting and CSV export

Enter multiple comma-separated IDs (e.g. `12345, 12346, 12347`) to get tabbed views for side-by-side comparison. Tab labels are auto-extracted from the action name's phase/group suffix.

## Weekly Report

Click **Weekly Report** to generate a comprehensive deployment summary across an entire BigFix site.

### How It Works

1. A dialog prompts for **Site Name**, **Start Date**, and **End Date** (max 7-day window)
2. The tool queries the BigFix server using a **Session Relevance** query via `POST /api/query`:

```
(id of it, name of it) of bes actions whose (
    name of it starts with "Update:" 
    AND site of source fixlet of it as string contains "<SiteName>" 
    AND time issued of it > (now - <DaysBack> * day)
)
```

This server-side filter efficiently returns only matching actions without pulling the full action list. The relevance expression filters by:
- **Action name prefix** — Only actions starting with `"Update:"` (standard deployment naming convention)
- **Source site** — Matches the site name you specify (partial match supported)  
- **Time window** — Only actions issued within the selected date range

3. For each discovered action, the tool fetches per-endpoint status via the REST API
4. Actions are parsed from the naming convention `Update: <Package> <Version>: <Phase>` to extract package names and deployment phases (Pilot, Prod, Force, Conference/Training Rooms, etc.)
5. Results are displayed in a summary table showing per-action success rates, endpoint counts, and phase breakdowns — plus a package-level rollup chart

### Weekly Report Output

- **Summary DataGrid** — Every action with: Action ID, Name, Package, Phase, Total/Relevant/Fixed/Failed/Running/Pending/Expired/Not Relevant counts, Success Rate %
- **Package rollup bar chart** — Aggregated success rates grouped by software package
- **Color-coded success rates** — Green (≥80%), Yellow (≥50%), Red (<50%)
- **CSV export** — Export the full summary table for stakeholder reporting

## API Endpoints Used

| Method | Endpoint | Purpose |
|--------|----------|---------|
| `GET` | `/api/login` | Connection test |
| `GET` | `/api/action/{id}` | Action name/details |
| `GET` | `/api/action/{id}/status` | Per-computer deployment results |
| `POST` | `/api/query` | Session Relevance queries (weekly report action discovery) |

### About `POST /api/query`

The BigFix REST API uses `POST` for relevance queries because the relevance expression is sent in the request body (as `application/x-www-form-urlencoded`). This is a **read-only operation** — nothing is created or modified on the server. POST is used instead of GET because relevance expressions can be complex and exceed URL length limits.

## Status Mapping

The tool maps BigFix's various status strings to standardized categories:

| Category | Matches |
|----------|---------|
| ✅ Fixed | Fixed, executed successfully, completed, succeeded |
| ❌ Failed | Failed, error |
| 🔄 Running | Running, Evaluating, executing |
| ⏳ Pending | Waiting, Pending, locked |
| ⬜ Not Relevant | Not Relevant, not applicable |
| ⏰ Expired | Expired |

## Notes
- Self-signed certs are handled automatically (common in BigFix deployments)
- Completion % always excludes "Not Relevant" endpoints from the denominator
- Timeline chart needs 2+ completed endpoints to render the S-curve
- All API calls and errors are logged in CMTrace format for troubleshooting
