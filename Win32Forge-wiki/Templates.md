# Templates

Templates are JSON files in the `Templates\` folder that define how an app is deployed to Intune. Rather than configuring install commands, assignments, and return codes for every app individually, you define them once in a template and reuse them.

---

## Included templates

| Template | Assignment | Intent | PSADT |
| --- | --- | --- | --- |
| `PSADT-Required` | All Devices | Required | Yes |
| `PSADT-Available` | All Users | Available | Yes |
| `PSADT-Groups` | Specific Azure AD group | Required | Yes |
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

### Assignment — Azure AD group

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

Replace `GroupName` and `GroupID` with your Azure AD group details. Multiple groups can be added to the `Groups` array.

Leave `FilterName` and `FilterID` empty if you are not using an Intune assignment filter. Set `FilterIntent` to `include` or `exclude`.

---

## Creating a custom template

1. Copy one of the existing templates in `Templates\` as a starting point
2. Edit the JSON fields as needed
3. Save with a descriptive name, e.g. `PSADT-Required-Servers.json`
4. The new template appears in the template picker immediately — no restart needed

---

## Setting a default template

Set `DefaultTemplate` in `Config\config.json` to the template filename without the `.json` extension:

```json
"DefaultTemplate": "PSADT-Required"
```

This template is pre-selected for every new upload. You can override it per-app in the upload form or bulk manager.
