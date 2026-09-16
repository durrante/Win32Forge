# Changelog

All notable changes to Win32Forge will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Fixed
- **Selecting x64 + arm64 (or x86 + arm64) no longer fails the upload.** The GUI builds a combined
  architecture value (`x64arm64` / `x86arm64`), but the IntuneWin32App module's `ValidateSet` only
  accepted `x64`, `x86`, `arm64`, `x64x86` and `AllWithARM64`, so every affected app died with
  *"Cannot validate argument on parameter 'Architecture'"*. A new module patch (**Patch F**) adds
  both mixed combinations and their Graph values (`x64,arm64` and `x86,arm64`) — the Graph
  `allowedArchitectures` property is a comma-separated flag list, so these were always valid; only
  the module's lookup table was missing them.
- **x86 + arm64 no longer silently becomes "all architectures".** The upload form previously mapped
  that combination to `AllWithARM64`, quietly adding x64. The Template Editor had the opposite bug —
  it dropped arm64 and saved plain `x86`.
- **Added a fallback if the module patch can't be applied** (module not writable, future module
  version). Instead of failing the upload, Win32Forge reads the `ValidateSet` the installed module
  actually exposes and widens to the nearest supported superset, with a warning in the log.
- **OneDrive "Files On-Demand" no longer breaks packaging.** `IntuneWinAppUtil.exe` cannot read a
  dehydrated cloud placeholder and aborts mid-zip with
  *"System.IO.IOException: The cloud operation is invalid"* (Win32 error 362). Win32Forge now
  downloads any cloud-only files in the source folder before packaging, always against the real
  path rather than through the long-path junction, which is where the recall is most likely to be
  rejected. If the error still occurs, it retries once after a full read of every source file, and
  failing that reports **which files** could not be read plus how to fix it — instead of the raw
  .NET stack trace.
- **OneDrive files with paths over 260 characters now download too.** With Windows long path
  support off (`LongPathsEnabled=0`, the default), OneDrive rejects the on-read download for those
  files. Win32Forge now temporarily marks them *Always keep on this device*, waits for OneDrive to
  sync them, then clears the setting again (the files stay local for packaging).

---

## [1.1.0] — 2026-06-19

A reliability + bulk-workflow release. Adds folder-driven metadata, faster bulk
operations, and a set of automatic compatibility patches for the IntuneWin32App
module that fix real-world failures on non-US, large-app, and multi-word-category
deployments.

### Added
- **Bulk: Import Subfolders (1 level)** — pick a parent folder and import every immediate subfolder as an app row in one click (no recursion). A single summary window reports, per app, whether a PSADT package, detection script, logo, and metadata file were found, and lists any apps still needing a detection rule.
- **Bulk: Set Template** — apply a template to the currently selected rows only (with confirmation), instead of having to change each row individually.
- **Per-app metadata files** — Win32Forge now auto-detects an optional `metadata.json` or `metadata.txt` in the root of a source folder and maps it onto the app's **Description**, **Information URL**, **Privacy URL**, and **Categories**. Works in both the single-upload form and the bulk manager (fills blank fields only).
- **Acquisition-method token in package names** — `.intunewin` files now carry the `Evergreen` / `WinGet` / `URLFallback` / `URL` token from the source folder name, e.g. `Google_Chrome_Latest_Evergreen_PSADT.intunewin`.
- **Template Editor: PSADT auto-fill** — ticking *PSADT Package* now pre-fills the install/uninstall commands with the framework defaults.
- **Long path (>260 char) support** — packaging a deep source tree that exceeds the Windows `MAX_PATH` limit now automatically retries via a short directory junction, so deeply-nested PSADT payloads on long OneDrive paths package successfully.
- **Tool version** is now shown in the window title bar and the generated-documentation footer, and stamped on every script.

### Changed
- Bulk manager window is larger (1800×840) and the toolbar wraps, so buttons are never pushed off-screen.
- Bulk manager Description cells now accept multi-line / pasted text.
- JSON export now always includes **Categories**, and writes additional requirement rules as a consistent array.
- **Generated documentation:** the app description is now its own section (long markdown descriptions no longer break the metadata table); the footer links to the GitHub repo and shows the tool version; the footer logo image was removed.
- The IntuneWin32App module compatibility patches now run automatically on every launch (idempotent and non-destructive).

### Fixed
- **Windows 11 24H2 / 23H2** can now be selected as the minimum supported OS (previously failed parameter validation).
- **App categories containing a space** (e.g. *Data management*, *Computer management*) now apply correctly instead of failing the upload — the category lookup is now resolved locally rather than via an unreliable server-side filter.
- **Large-app uploads** no longer flood SAS-token-renewal warnings (and the renewal now actually succeeds), preventing failed uploads of big packages.
- **Non-US locale (e.g. en-GB) DateTime crash** during upload is resolved.
- **Deep source paths over 260 characters** now package successfully (see Added → long path support).
- Importing a queue whose app had an **additional requirement rule** no longer errors when opening that app's full configuration.
- Categories now survive **export → import** round-trips.
- Underlying module errors are now surfaced to the activity log instead of a generic "no App ID returned" message.

> **Note on the bundled module patches:** Win32Forge patches the installed
> [IntuneWin32App](https://github.com/MSEndpointMgr/IntuneWin32App) module in place
> (idempotently, on launch) to fix the locale, Windows 11 24H2, large-upload, and
> category issues above. The module itself is unmodified upstream — credit and thanks
> to MSEndpointMgr; these are local compatibility shims.

---

## [1.0.0] — 2026-04-18

### Initial public release

#### Features
- WPF GUI for packaging and uploading Win32 apps to Microsoft Intune
- JSON template system — define deployment settings once, reuse across apps
- PSAppDeployToolkit (PSADT) v4 support with auto-detection of app name and version
- Bulk upload manager — queue multiple apps in a grid for batch processing
- Automatic Markdown documentation generation for each uploaded app
- Settings wizard for configuring tenant, auth, paths, template, and logging
- Verbose logging with structured log file output
- Support for Microsoft Graph CLI and custom Entra ID app registration auth
- Assignment types: All Devices, All Users, specific Azure AD groups
- Intune filter support on group assignments
- Detection rule builder (registry, file, MSI product code)
- Requirement rule support (OS version, architecture)
- Logo support for app icons (PNG/JPG/JPEG)
- One-time setup script that installs modules and downloads IntuneWinAppUtil.exe

#### Included templates
- `PSADT-Required` — PSADT v4, Required, All Devices
- `PSADT-Available` — PSADT v4, Available, All Users
- `PSADT-Groups` — PSADT v4, Required, specific Azure AD group (placeholders)
- `Generic-Required` — Standard Win32, Required, All Devices
- `Generic-Available` — Standard Win32, Available, All Users
