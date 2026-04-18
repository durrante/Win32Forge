# Tools

This folder is where `IntuneWinAppUtil.exe` should be placed.

`IntuneWinAppUtil.exe` is a Microsoft tool used to package Win32 apps into the `.intunewin` format required by Intune. It is **not included in this repository** — run `Setup-Win32Forge.ps1` to download it automatically:

```powershell
pwsh .\Setup-Win32Forge.ps1
```

The setup script downloads the latest release directly from the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) GitHub repository and places it here.

You can also download it manually from:  
https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
