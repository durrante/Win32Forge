# Bulk Upload

The Bulk Upload Manager is a full app catalogue editor. It lets you build a queue of apps — each with its own full set of Intune properties — and upload them all in one run. It is accessible from the **Bulk** button in the main Win32Forge window.

---

## What you can configure per row

Every row in the grid is a complete app record. The same fields available in the single upload form are available here, row by row:

| Field | Notes |
| --- | --- |
| Source Folder | Path to the app package — PSADT or standard Win32 |
| Template | Applied per row; auto-fills commands and defaults |
| Display Name | Auto-filled for PSADT packages |
| Version | Auto-filled for PSADT packages |
| Publisher | Auto-filled for PSADT packages |
| Setup File | Auto-detected; can be overridden |
| Install Command | From template; editable for non-PSADT apps |
| Uninstall Command | From template; editable for non-PSADT apps |
| Description | Optional free-text |
| Information URL | Optional URL |
| Privacy URL | Optional URL |
| Category | Loaded from your tenant |
| Detection Method | Configured via the full edit form (see below) |
| Assignment | Configured via the full edit form (see below) |
| Logo Path | PNG or JPG only |
| Status | Live upload status per row |

**Using templates reduces the number of fields you need to fill** — assignment, commands, return codes, architecture, and restart behaviour all come from the template, leaving only app-specific values to enter.

---

## Adding apps

### Scan a folder

Click **Scan Folder** and select a parent directory. Win32Forge scans for subfolders that look like app packages and adds each one as a row, auto-detecting PSADT packages and extracting their metadata.

### Add individually

Click **+ New Row** to add a blank row, then browse to the source folder in the cell.

### Import from JSON

Click **Load JSON** to import a previously saved queue. This is useful for repeatable deployments or sharing a queue between team members.

---

## Auto-detection when a source folder is set

Whenever a source folder is selected or scanned, Win32Forge automatically looks for two things:

**Detection script**
Scans the root of the source folder for any `.ps1` file with "detection" in its name. If found and no detection method has already been set for the row, it is automatically set as the PowerShell detection script.

**Logo**
Scans only the root of the source folder for a PNG, JPG, or JPEG file. If found and no logo has been set, the first match is automatically used as the app logo.

A notification is shown each time so you can confirm or override the auto-detected values. Both can be changed at any point by editing the row.

---

## Editing detection and assignment

Detection method and assignment cannot be fully configured in the grid cells alone. Click **Edit Selected** (or double-click a row and use the Edit Full button) to open the complete single-app upload form for that row. Changes saved in the form are written back to the row.

You can also right-click any row for a context menu with Edit Full, Delete, and Run Now options.

---

## Uploading

Click **Start Upload** to process all rows with status *Ready* (or only the selected rows if you have a selection). Win32Forge processes apps sequentially, updating the Status column as it goes:

- **Ready** — waiting to upload
- **Uploading...** — in progress
- **OK** — completed successfully
- **FAILED: \<error\>** — failed; error message shown in the cell

Errors do not stop the queue — the next row continues automatically. Once the run finishes, use **Clear Completed** to remove successful rows and retry any failures.

---

## Export / save

Click **Save Selected** to export the selected rows to a JSON file. The JSON can be loaded back later with **Load JSON**, or used for headless (unattended) bulk uploads:

```powershell
pwsh .\Invoke-Win32Forge.ps1 -BulkFile "C:\apps\upload-queue.json"
```

---

## Documentation

After each successful upload, Win32Forge writes a Markdown documentation file to the `DocumentationPath` folder from your config. Filename format:

```text
AppName_Version_YYYYMMDD.md
```

---

## Tips

- Use **Scan Folder** on a folder of PSADT packages to queue an entire app catalogue in seconds
- Assign different templates per row — mix `PSADT-Required` for system apps and `PSADT-Available` for optional tools in the same run
- Enable [[Verbose Logging|Configuration]] before a large bulk run to capture a full diagnostic log
