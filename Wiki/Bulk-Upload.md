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

### Import Subfolders (1 level)

Click **+ Import Subfolders (1 level)** and select a parent directory. Win32Forge imports **every immediate subfolder** (one level down — no recursion) as a row, auto-detecting PSADT packages, detection scripts, logos, and metadata files for each.

When it finishes, a **summary window** lists every imported app and shows, per app, whether a PSADT package, detection script, logo, and metadata file were found — and calls out any app that still needs a detection rule before it can upload. This is the fastest way to queue a whole catalogue of packages laid out as one-folder-per-app.

### Add individually

Click **+ New Row** to add a blank row, then browse to the source folder in the cell.

### Import from JSON

Click **Load JSON** to import a previously saved queue. This is useful for repeatable deployments or sharing a queue between team members.

---

## Auto-detection when a source folder is set

Whenever a source folder is selected or scanned, Win32Forge automatically looks in the **root** of that folder for:

**Detection script**
Any `.ps1` file with "detection" in its name. If found and no detection method has already been set for the row, it is automatically set as the PowerShell detection script.

**Logo**
A PNG, JPG, or JPEG file. If found and no logo has been set, the first match is automatically used as the app logo.

**Metadata file**
An optional `metadata.txt` or `metadata.json`. If present, it fills the row's **Description**, **Information URL**, **Privacy URL**, and **Categories** (only the fields left blank). See [[PSADT Support|PSADT-Support]] / the README for the `metadata.txt` format.

A notification is shown each time so you can confirm or override the auto-detected values. All can be changed at any point by editing the row.

---

## Editing detection and assignment

Detection method and assignment cannot be fully configured in the grid cells alone. Click **Edit Selected** (or double-click a row and use the Edit Full button) to open the complete single-app upload form for that row. Changes saved in the form are written back to the row.

You can also right-click any row for a context menu with Edit Full, Delete, and Run Now options.

---

## Changing the template on multiple rows

To switch several rows to a different template at once, select them and click **Set Template ▾**, then pick the template. It is applied to the selected rows only (with a confirmation prompt), overwriting their template-driven settings — assignment, architecture, install/uninstall commands, return codes, restart behaviour — while leaving each row's source folder, detection, logo, description, and URLs untouched.

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
