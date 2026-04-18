<#
.SYNOPSIS
    WPF Settings editor — configure paths, defaults, and authentication.

.DESCRIPTION
    Opens a dialog allowing the user to edit config.json settings:
      - Default output path for .intunewin files
      - Documentation output path
      - IntuneWinAppUtil.exe path (with browse + auto-download)
      - Default template selection
      - Default author name (used in documentation)
      - Auth method (Microsoft Graph CLI vs Custom App Registration)
      - Tenant ID and Client ID

    Changes are saved to config.json and applied in-memory immediately.
#>

function Show-SettingsEditor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$ConfigPath = '',
        [string]$TemplateFolder = ''
    )

    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms | Out-Null

    # Enumerate available templates
    $availableTemplates = @()
    if ($TemplateFolder -and (Test-Path $TemplateFolder)) {
        $availableTemplates = @(Get-ChildItem -Path $TemplateFolder -Filter '*.json' -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty BaseName | Sort-Object)
    }

    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Win32Forge — Settings"
    Width="640" Height="760"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize">

  <Window.Resources>
    <Style x:Key="Section" TargetType="TextBlock">
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#333"/>
      <Setter Property="Margin" Value="0,18,0,10"/>
    </Style>
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Width" Value="150"/>
      <Setter Property="Foreground" Value="#444"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="5,4"/>
      <Setter Property="BorderBrush" Value="#CCC"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="52"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0.5" EndPoint="1,0.5">
          <GradientStop Color="#0693E3" Offset="0"/>
          <GradientStop Color="#9B51E0" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <StackPanel VerticalAlignment="Center" Margin="18,0">
        <TextBlock Text="Settings" FontSize="17" FontWeight="Light" Foreground="White"/>
        <TextBlock Text="Configure paths, defaults, and authentication" FontSize="11" Foreground="#D4C5F9" Margin="0,1,0,0"/>
      </StackPanel>
    </Border>

    <!-- Content -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="18,4,18,12">
      <StackPanel>

        <!-- ── PATHS ── -->
        <TextBlock Style="{StaticResource Section}" Text="Paths"/>

        <!-- Output Path -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Output Path"/>
          <TextBox   Grid.Column="1" x:Name="TxtOutputPath"/>
          <Button    Grid.Column="2" x:Name="BtnBrowseOutput" Content="Browse" Margin="6,0,0,0"/>
        </Grid>

        <!-- Docs Path -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Documentation Path"/>
          <TextBox   Grid.Column="1" x:Name="TxtDocsPath"/>
          <Button    Grid.Column="2" x:Name="BtnBrowseDocs" Content="Browse" Margin="6,0,0,0"/>
        </Grid>

        <!-- IntuneWinAppUtil -->
        <Grid Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="IntuneWinAppUtil"/>
          <TextBox   Grid.Column="1" x:Name="TxtUtilPath"/>
          <Button    Grid.Column="2" x:Name="BtnBrowseUtil" Content="Browse" Margin="6,0,6,0"/>
          <Button    Grid.Column="3" x:Name="BtnDownloadUtil" Content="Download"
                     Background="#5BA3E8" Foreground="White" BorderThickness="0"/>
        </Grid>
        <TextBlock Text="IntuneWinAppUtil.exe packages source folders into .intunewin files required for Intune uploads."
                   Foreground="#888" FontSize="11" Margin="150,2,0,0" TextWrapping="Wrap"/>

        <Separator Margin="0,14,0,0" Background="#DDD"/>

        <!-- ── DEFAULTS ── -->
        <TextBlock Style="{StaticResource Section}" Text="Defaults"/>

        <!-- Default Template -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Default Template"/>
          <ComboBox  Grid.Column="1" x:Name="CmbDefaultTemplate"/>
        </Grid>


        <Separator Margin="0,14,0,0" Background="#DDD"/>

        <!-- ── AUTHENTICATION ── -->
        <TextBlock Style="{StaticResource Section}" Text="Authentication"/>

        <!-- Auth Method -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Auth Method"/>
          <StackPanel Grid.Column="1">
            <RadioButton x:Name="RdoAuthCLI" Margin="0,0,0,4"
                         Content="Microsoft Graph CLI (no app registration required)"/>
            <RadioButton x:Name="RdoAuthApp"
                         Content="Custom App Registration (requires Azure AD app)"/>
          </StackPanel>
        </Grid>

        <!-- Tenant ID -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Tenant ID"/>
          <TextBox   Grid.Column="1" x:Name="TxtTenantID"/>
        </Grid>

        <!-- Client ID (Custom only) -->
        <Grid x:Name="PanelClientID" Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Client ID"/>
          <TextBox   Grid.Column="1" x:Name="TxtClientID"/>
        </Grid>

        <!-- Re-authenticate -->
        <Grid Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0"/>
          <Button Grid.Column="1" x:Name="BtnReauth" HorizontalAlignment="Left"
                  Content="Re-authenticate / Refresh Permissions"
                  Background="#4A2B8F" Foreground="White" BorderThickness="0" Padding="12,6"/>
        </Grid>
        <TextBlock Margin="150,4,0,0" Foreground="#888" FontSize="11" TextWrapping="Wrap"
                   Text="Changing Tenant ID or Auth Method requires signing out and reconnecting."/>

        <Separator Margin="0,14,0,0" Background="#DDD"/>

        <!-- ── LOGGING ── -->
        <TextBlock Style="{StaticResource Section}" Text="Logging"/>

        <!-- Enable Verbose Logging -->
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Verbose Logging"/>
          <CheckBox  Grid.Column="1" x:Name="ChkVerboseLog" VerticalAlignment="Center"
                     Content="Log commands and outputs to file"/>
        </Grid>

        <!-- Log File Path -->
        <Grid x:Name="PanelLogPath" Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Style="{StaticResource Lbl}" Text="Log File Path"/>
          <TextBox   Grid.Column="1" x:Name="TxtLogPath"/>
          <Button    Grid.Column="2" x:Name="BtnBrowseLog" Content="Browse" Margin="6,0,0,0"/>
        </Grid>
        <TextBlock x:Name="TxtLogNote"
                   Text="Logs packaging, upload, Graph API calls, and errors — useful for troubleshooting failures."
                   Foreground="#888" FontSize="11" Margin="150,2,0,0" TextWrapping="Wrap"/>

      </StackPanel>
    </ScrollViewer>

    <!-- Footer -->
    <Border Grid.Row="2" Background="#F5F5F5" BorderBrush="#DDD" BorderThickness="0,1,0,0" Padding="16,10">
      <Grid>
        <TextBlock x:Name="TxtStatus" Text="" VerticalAlignment="Center" Foreground="#C00" FontSize="11"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancel" Content="Cancel" Width="80" Margin="0,0,8,0"/>
          <Button x:Name="BtnSave"   Content="Save"   Width="80"
                  Background="#5BA3E8" Foreground="White" BorderThickness="0"/>
        </StackPanel>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    function Find { param($n) $window.FindName($n) }

    $txtOutputPath      = Find 'TxtOutputPath'
    $btnBrowseOutput    = Find 'BtnBrowseOutput'
    $txtDocsPath        = Find 'TxtDocsPath'
    $btnBrowseDocs      = Find 'BtnBrowseDocs'
    $txtUtilPath        = Find 'TxtUtilPath'
    $btnBrowseUtil      = Find 'BtnBrowseUtil'
    $btnDownloadUtil    = Find 'BtnDownloadUtil'
    $cmbDefaultTemplate = Find 'CmbDefaultTemplate'
    $rdoAuthCLI         = Find 'RdoAuthCLI'
    $rdoAuthApp         = Find 'RdoAuthApp'
    $txtTenantID        = Find 'TxtTenantID'
    $panelClientID      = Find 'PanelClientID'
    $txtClientID        = Find 'TxtClientID'
    $btnReauth          = Find 'BtnReauth'
    $chkVerboseLog      = Find 'ChkVerboseLog'
    $txtLogPath         = Find 'TxtLogPath'
    $btnBrowseLog       = Find 'BtnBrowseLog'
    $panelLogPath       = Find 'PanelLogPath'
    $txtLogNote         = Find 'TxtLogNote'
    $btnSave            = Find 'BtnSave'
    $btnCancel          = Find 'BtnCancel'
    $txtStatus          = Find 'TxtStatus'

    #region Populate template list
    foreach ($tpl in $availableTemplates) {
        $cmbDefaultTemplate.Items.Add($tpl) | Out-Null
    }
    #endregion

    #region Load current values
    $txtOutputPath.Text = $Config.DefaultOutputPath    ?? ''
    $txtDocsPath.Text   = $Config.DocumentationPath    ?? ''
    $txtUtilPath.Text   = $Config.IntuneWinAppUtilPath ?? ''
    $txtTenantID.Text   = $Config.TenantID             ?? ''
    $txtClientID.Text   = $Config.ClientID             ?? ''

    # Auth method radio
    if ($Config.AuthMethod -and $Config.AuthMethod -ne 'MicrosoftGraphCLI') {
        $rdoAuthApp.IsChecked = $true
    } else {
        $rdoAuthCLI.IsChecked = $true
    }

    # Default template
    $curTpl = $Config.DefaultTemplate ?? ''
    if ($curTpl -and $cmbDefaultTemplate.Items.Contains($curTpl)) {
        $cmbDefaultTemplate.SelectedItem = $curTpl
    } elseif ($cmbDefaultTemplate.Items.Count -gt 0) {
        $cmbDefaultTemplate.SelectedIndex = 0
    }

    # Logging
    $chkVerboseLog.IsChecked = ($Config.VerboseLogging -eq $true)
    $txtLogPath.Text          = $Config.LogPath ?? ''
    #endregion

    #region Logging panel visibility
    $updateLogPanel = {
        $vis = if ($chkVerboseLog.IsChecked) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
        $panelLogPath.Visibility = $vis
        $txtLogNote.Visibility   = $vis
    }
    $chkVerboseLog.Add_Checked($updateLogPanel)
    $chkVerboseLog.Add_Unchecked($updateLogPanel)
    & $updateLogPanel
    #endregion

    #region Auth method visibility
    $updateAuthPanel = {
        $panelClientID.Visibility = if ($rdoAuthApp.IsChecked) {
            [System.Windows.Visibility]::Visible
        } else {
            [System.Windows.Visibility]::Collapsed
        }
    }
    $rdoAuthCLI.Add_Checked($updateAuthPanel)
    $rdoAuthApp.Add_Checked($updateAuthPanel)
    & $updateAuthPanel
    #endregion

    #region Browse buttons
    $btnBrowseOutput.Add_Click({
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description = 'Select default output folder for .intunewin files'
        if ($txtOutputPath.Text -and (Test-Path $txtOutputPath.Text)) {
            $dlg.SelectedPath = $txtOutputPath.Text
        }
        if ($dlg.ShowDialog() -eq 'OK') { $txtOutputPath.Text = $dlg.SelectedPath }
    })

    $btnBrowseDocs.Add_Click({
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description = 'Select folder for generated documentation files'
        if ($txtDocsPath.Text -and (Test-Path $txtDocsPath.Text)) {
            $dlg.SelectedPath = $txtDocsPath.Text
        }
        if ($dlg.ShowDialog() -eq 'OK') { $txtDocsPath.Text = $dlg.SelectedPath }
    })

    $btnBrowseUtil.Add_Click({
        $dlg = [System.Windows.Forms.OpenFileDialog]::new()
        $dlg.Title  = 'Select IntuneWinAppUtil.exe'
        $dlg.Filter = 'IntuneWinAppUtil|IntuneWinAppUtil.exe|Executable|*.exe'
        if ($txtUtilPath.Text) {
            $parent = Split-Path $txtUtilPath.Text -Parent -ErrorAction SilentlyContinue
            if ($parent -and (Test-Path $parent)) { $dlg.InitialDirectory = $parent }
        }
        if ($dlg.ShowDialog() -eq 'OK') { $txtUtilPath.Text = $dlg.FileName }
    })

    $btnBrowseLog.Add_Click({
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title      = 'Choose log file location'
        $dlg.Filter     = 'Log file|*.log|Text file|*.txt|All files|*.*'
        $dlg.DefaultExt = 'log'
        if ($txtLogPath.Text) {
            $parent = Split-Path $txtLogPath.Text -Parent -ErrorAction SilentlyContinue
            if ($parent -and (Test-Path $parent)) { $dlg.InitialDirectory = $parent }
            $dlg.FileName = Split-Path $txtLogPath.Text -Leaf -ErrorAction SilentlyContinue
        }
        if ($dlg.ShowDialog() -eq 'OK') { $txtLogPath.Text = $dlg.FileName }
    })
    #endregion

    #region Download IntuneWinAppUtil
    $btnDownloadUtil.Add_Click({
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description = 'Choose folder to save IntuneWinAppUtil.exe'
        if ($txtUtilPath.Text) {
            $parent = Split-Path $txtUtilPath.Text -Parent -ErrorAction SilentlyContinue
            if ($parent -and (Test-Path $parent)) { $dlg.SelectedPath = $parent }
        }
        if ($dlg.ShowDialog() -ne 'OK') { return }

        $destPath = Join-Path $dlg.SelectedPath 'IntuneWinAppUtil.exe'
        $btnDownloadUtil.IsEnabled = $false
        $btnDownloadUtil.Content   = 'Downloading...'
        $txtStatus.Text            = ''

        try {
            $url = 'https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe'
            Invoke-WebRequest -Uri $url -OutFile $destPath -UseBasicParsing -ErrorAction Stop
            $txtUtilPath.Text = $destPath
            $txtStatus.Text   = ''
            [System.Windows.MessageBox]::Show(
                "Downloaded successfully:`n$destPath",
                'Download Complete', 'OK', 'Information')
        }
        catch {
            $txtStatus.Text = "Download failed: $_"
        }
        finally {
            $btnDownloadUtil.IsEnabled = $true
            $btnDownloadUtil.Content   = 'Download'
        }
    })
    #endregion

    #region Re-authenticate
    $btnReauth.Add_Click({
        [System.Windows.MessageBox]::Show(
            "To refresh permissions or re-authenticate:`n`n" +
            "1. Click 'Sign Out' on the main window`n" +
            "2. Click 'Connect to Intune'`n`n" +
            "This starts a new interactive login where you can consent to any updated or additional permissions.",
            'Re-authenticate', 'OK', 'Information')
    })
    #endregion

    $btnCancel.Add_Click({ $window.DialogResult = $false; $window.Close() })

    #region Save
    $btnSave.Add_Click({
        $txtStatus.Text = ''
        $outPath      = $txtOutputPath.Text.Trim()
        $docsPath     = $txtDocsPath.Text.Trim()
        $utilPath     = $txtUtilPath.Text.Trim()
        $tenantID     = $txtTenantID.Text.Trim()
        $clientID     = $txtClientID.Text.Trim()
        $verboseLog   = [bool]$chkVerboseLog.IsChecked
        $logFilePath  = $txtLogPath.Text.Trim()
        # Validation
        $errors = [System.Collections.Generic.List[string]]::new()
        if (-not $tenantID)                              { $errors.Add('Tenant ID is required.') }
        if ($rdoAuthApp.IsChecked -and -not $clientID)  { $errors.Add('Client ID is required for Custom App Registration.') }
        if ($verboseLog -and -not $logFilePath)          { $errors.Add('Log File Path is required when verbose logging is enabled.') }

        if ($errors.Count -gt 0) {
            $txtStatus.Text = $errors -join '  |  '
            return
        }

        # Create folders if needed
        foreach ($p in @($outPath, $docsPath) | Where-Object { $_ }) {
            if (-not (Test-Path $p)) {
                try { New-Item -ItemType Directory -Path $p -Force | Out-Null }
                catch { $txtStatus.Text = "Warning: could not create directory: $p" }
            }
        }

        $authMethod = if ($rdoAuthApp.IsChecked) { 'CustomApp' } else { 'MicrosoftGraphCLI' }
        $selTpl     = if ($null -ne $cmbDefaultTemplate.SelectedItem) { [string]$cmbDefaultTemplate.SelectedItem } else { '' }

        # Update $Config in-place (Add-Member -Force handles both new and existing properties)
        $updates = @{
            DefaultOutputPath    = $outPath
            DocumentationPath    = $docsPath
            IntuneWinAppUtilPath = $utilPath
            DefaultTemplate      = $selTpl
            TenantID             = $tenantID
            ClientID             = $clientID
            AuthMethod           = $authMethod
            VerboseLogging       = $verboseLog
            LogPath              = $logFilePath
        }
        foreach ($kv in $updates.GetEnumerator()) {
            $Config | Add-Member -NotePropertyName $kv.Key -NotePropertyValue $kv.Value -Force
        }

        # Persist to config.json
        if ($ConfigPath) {
            try {
                $Config | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Settings applied in memory but could not save to file:`n$_",
                    'Save Warning', 'OK', 'Warning')
            }
        }

        $window.DialogResult = $true
        $window.Close()
    })
    #endregion

    $window.ShowDialog() | Out-Null
}
