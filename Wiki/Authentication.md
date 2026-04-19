# Authentication

Win32Forge supports two authentication methods. Choose one during setup — it can be changed later in the Settings window or by editing `Config\config.json`.

---

## MicrosoftGraphCLI (recommended)

Uses the **Microsoft Graph Command Line Tools** public client application — a multi-tenant app registered by Microsoft that allows delegated access without you needing your own app registration.

**Pros:**

- No app registration required
- Fastest to set up
- Works for most organisations

**Cons:**

- Requires interactive browser login each session (no background/unattended support)
- Uses Microsoft's shared public client ID

**Config:**

```json
"AuthMethod": "MicrosoftGraphCLI",
"ClientID": "14d82eec-204b-4c2f-b7e8-296a70dab67e"
```

The Client ID above is Microsoft's official Graph Command Line Tools app — leave it as-is.

---

## CustomApp

Uses your own **Entra ID app registration**. Useful if your organisation restricts use of Microsoft's public client, or if you want a named, auditable app in your tenant.

### Required delegated permissions

| Permission | Purpose |
| --- | --- |
| `DeviceManagementApps.ReadWrite.All` | Upload and assign Win32 apps |
| `DeviceManagementConfiguration.Read.All` | Load Intune assignment filters |
| `Group.Read.All` | Search and resolve Entra ID groups for assignments |

> `DeviceManagementConfiguration.Read.All` is optional — if missing, Intune filters will not load, but all other features work normally.

### Creating the app registration

1. Go to **Entra ID → App registrations → New registration**
2. Name it `Win32Forge` (or anything you prefer)
3. Set **Supported account types** to *Accounts in this organizational directory only*
4. Set **Redirect URI** to `Public client/native` → `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**
6. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions**
7. Add the three permissions listed above
8. Click **Grant admin consent**
9. Copy the **Application (client) ID** — paste this as `ClientID` in your config

**Config:**

```json
"AuthMethod": "CustomApp",
"ClientID": "your-app-registration-client-id-here"
```

---

## How authentication works at runtime

When Win32Forge launches, it calls `Connect-MSIntuneGraph` from the IntuneWin32App module. This opens a browser window for interactive delegated login. The resulting token is stored in `$Global:AuthenticationHeader` and reused for all Graph API calls during the session.

You will be prompted to sign in once per session. Tokens are not persisted to disk.
