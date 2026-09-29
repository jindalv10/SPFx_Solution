# Site Migration Tracker – list provisioning

Creates the six SharePoint lists for the Site Migration Tracker (Site Register, Update Queue,
Outreach Batches, Audit Log, Import Log, Configuration) on any site, with columns, indexes,
views and default configuration.

## Files

| File | Purpose |
|---|---|
| `Deploy-SiteMigrationTracker.ps1` | Provisioning script (safe to re-run) |
| `site-migration-tracker.schema.json` | Lists, columns, indexes, views and default configuration |

## Prerequisites

- PowerShell 7.4+
- PnP.PowerShell (`Install-Module PnP.PowerShell -Scope CurrentUser`)
- An Entra ID app (client) ID for PnP sign-in – follow the team's PnP setup instructions
- Owner (Full Control) on the target site

## Run – Canada dev

```powershell
.\Deploy-SiteMigrationTracker.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/SPDM-CA-Tracking-DEV" `
  -ClientId "<app-client-id>" `
  -DeptCode CA -BusinessGroup "SLF Canada" `
  -EnterpriseTeamEmail "<enterprise team address>" `
  -SharedMailbox "<shared migration inbox>" `
  -SupportTeamEmail "<SharePoint Canada support address>" `
  -SeedConfig
```

## Run – another department

Same script, their site and values:

```powershell
.\Deploy-SiteMigrationTracker.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/SPDM-US-Tracking" `
  -ClientId "<app-client-id>" -DeptCode US -BusinessGroup "SLF US" -SeedConfig
```

Two departments sharing one site: add `-ListPrefix US` so lists become `Lists/USSiteRegister`, "US Site Register", and so on.

## Changing the design

Edit the schema JSON and run the script again:

- **New column:** add an entry to the list's `fields`. It is created on the next run.
- **New choice value:** add it to `choices`. It is appended to the existing column.
- **New index:** set `"indexed": true`.
- **New view:** add an entry to `views`. Existing views with the same title are left alone.

The script never deletes lists, columns, views or data. Removals are done by hand after review.
Every run writes a log file (`provision-<date>.log`) next to the script.

## Notes

- `SiteKey` is a single line of text (255 characters max) so it can be indexed and unique. Keys longer
  than 255 characters are handled by the import flow (truncated and suffixed with the reference number).
- Choice values use a plain hyphen (`Hold - PII`) to avoid encoding problems in flows and filters.
- Site references in Update Queue, Audit Log and Outreach Batches are text (`SiteRefNo`), not lookups,
  so the lists stay portable between sites.
