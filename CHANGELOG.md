# Changelog

All notable changes to Win32Forge will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

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
