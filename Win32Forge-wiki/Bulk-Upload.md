# Bulk Upload

The Bulk Upload Manager lets you queue multiple apps and upload them all in one run. It is accessible from the **Bulk** button in the main Win32Forge window.

---

## Overview

The bulk manager displays a grid where each row represents one app. You can:

- Add app source folders individually or scan a parent folder to discover multiple apps at once
- Assign a different template and logo to each app
- Review and edit settings per-row before uploading
- Upload all queued apps sequentially with one click

---

## Adding apps

### Scan a folder

Click **Scan Folder** and select a parent directory. Win32Forge scans for subfolders that look like app packages (PSADT packages or folders containing an installer). Each discovered app is added as a row.

### Add individually

Click **Add App** to browse to a single app source folder.

---

## Grid columns

| Column | Description |
| --- | --- |
| **Source Path** | Path to the app source folder |
| **App Name** | Populated automatically for PSADT packages; edit manually for generic apps |
| **Version** | Populated automatically for PSADT packages; edit manually for generic apps |
| **Template** | Template to apply — defaults to the `DefaultTemplate` from config |
| **Logo Path** | Optional PNG or JPG logo for the app tile in Company Portal |
| **Status** | Upload status: Pending / Uploading / Done / Failed |

---

## Logos in bulk upload

- Accepted formats: **PNG, JPG, JPEG only**
- You can type a path directly, browse using the button in the cell, or drag and drop an image file onto the Logo Path cell
- Win32Forge validates the extension when you set a value — unsupported formats are rejected with a message

---

## Uploading

Click **Upload All** to process all rows with status *Pending*. Win32Forge processes apps sequentially, updating the Status column as it goes.

Rows with status *Done* or *Failed* are skipped on subsequent runs. To retry a failed row, reset its status to *Pending* by right-clicking the row.

---

## Documentation

After each successful upload, Win32Forge writes a Markdown documentation file to the `DocumentationPath` folder configured in `config.json`. The filename format is:

```
AppName_Version_YYYYMMDD.md
```

---

## Tips

- Use **Scan Folder** on a folder full of PSADT packages to queue an entire app catalogue in seconds
- Templates can be mixed — assign `PSADT-Required` to system apps and `PSADT-Available` to optional tools in the same bulk run
- Set up [[Verbose Logging|Configuration]] before a large bulk run to capture a full log for review
