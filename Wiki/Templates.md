# Templates

Templates are JSON files in the `Templates\` folder that define how an app is deployed to Intune. Rather than configuring install commands, assignments, and return codes for every app individually, you define them once in a template and reuse them.

---

## Included templates

The templates below are included as **examples only** to give you a starting point. They are not intended to be used as-is in production — build your own using the Template Editor to match your environment, naming convention, and assignment strategy.

| Template | Assignment | Intent | PSADT |
| --- | --- | --- | --- |
| `PSADT-Required` | All Devices | Required | Yes |
| `PSADT-Available` | All Users | Available | Yes |
| `PSADT-Groups` | Specific Entra ID group | Required | Yes |
| `Generic-Required` | All Devices | Required | No |
| `Generic-Available` | All Users | Available | No |

---

## Template structure

```json
{
  "TemplateName": "PSADT-Required",
  "Description": "Human-readable description shown in the template picker",
  "IsPSADT": true,

  "InstallCommandLine": "Invoke-AppDeployToolkit.exe -DeployMode Silent",
  "UninstallCommandLine": "Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent",

  "InstallExperience": "system",
  "RestartBehavior": "suppress",

  "Architecture": "x64",
  "MinimumSupportedWindowsRelease": "W10_2004",

  "MaximumInstallationTimeInMinutes": 60,
  "AllowAvailableUninstall": false,

  "ReturnCodes": [
    { "ReturnCode": 0,    "Type": "success"    },
    { "ReturnCode": 1707, "Type": "success"    },
    { "ReturnCode": 3010, "Type": "softReboot" },
    { "ReturnCode": 1641, "Type": "hardReboot" },
    { "ReturnCode": 1618, "Type": "retry"      }
  ],

  "Assignment": {
    "Type": "AllDevices",
    "Intent": "required",
    "Notification": "showAll"
  }
}
```

---

## Field reference

### Top-level fields

| Field | Values | Description |
| --- | --- | --- |
| `TemplateName` | string | Name shown in the template picker — should match the filename |
| `IsPSADT` | `true` / `false` | Enables PSADT auto-detection of app name and version |
| `InstallCommandLine` | string | Install command passed to Intune |
| `UninstallCommandLine` | string | Uninstall command passed to Intune |
| `InstallExperience` | `system` / `user` | Whether the app installs as System or the logged-in user |
| `RestartBehavior` | `suppress` / `allow` / `basedOnReturnCode` / `force` | Restart behaviour after install |
| `Architecture` | `x64` / `x86` / `arm64` | Target architecture |
| `MinimumSupportedWindowsRelease` | e.g. `W10_2004` | Minimum Windows version requirement |
| `MaximumInstallationTimeInMinutes` | integer | Timeout before Intune marks the install as failed |
| `AllowAvailableUninstall` | `true` / `false` | Show uninstall option in Company Portal |

### Assignment — All Devices / All Users

```json
"Assignment": {
  "Type": "AllDevices",
  "Intent": "required",
  "Notification": "showAll"
}
```

| Field | Values |
| --- | --- |
| `Type` | `AllDevices` / `AllUsers` |
| `Intent` | `required` / `available` / `uninstall` |
| `Notification` | `showAll` / `showReboot` / `hideAll` |

### Assignment — Entra ID group

```json
"Assignment": {
  "Type": "Group",
  "Groups": [
    {
      "GroupName": "YOUR-GROUP-NAME",
      "GroupID": "YOUR-GROUP-OBJECT-ID",
      "Intent": "required",
      "Notification": "showAll",
      "FilterName": "",
      "FilterID": "",
      "FilterIntent": "include"
    }
  ]
}
```

Replace `GroupName` and `GroupID` with your Entra ID group details. Multiple groups can be added to the `Groups` array.

Leave `FilterName` and `FilterID` empty if you are not using an Intune assignment filter. Set `FilterIntent` to `include` or `exclude`.

---

## Creating a custom template

Use the **Template Editor** inside Win32Forge (the Templates button in the main window) to create and edit templates without touching any JSON. Click **New**, fill in the form, and click **Save Template**. The new template appears in all template pickers immediately — no restart needed.

You can also duplicate an existing template as a starting point using the **Duplicate** button in the Template Editor.

If you prefer to edit JSON directly, templates are plain `.json` files in the `Templates\` folder and can be opened in any text editor.

---

## Setting a default template

The default template is set in the **Settings** window (the Settings button in the main Win32Forge window). It applies to all new uploads unless overridden.

The default can be overridden:

- **Per app** — change the template in the single upload form before uploading
- **Per row** — change the template in the bulk manager grid, row by row

You can also set `DefaultTemplate` directly in `Config\config.json` (filename without `.json`), but the Settings window is the recommended way.
