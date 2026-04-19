<#
.SYNOPSIS
    WPF main dashboard window for the Intune Win32 App Uploader.

.DESCRIPTION
    Shows the main application window with:
      - Connection status (tenant + user)
      - Action buttons: Upload App, Bulk Upload, Templates, Settings
      - Activity log showing results of uploads
      - Recent uploads list

    Calls Show-AppUploadForm for single-app uploads,
    and Invoke-ProcessApp to run the actual package/upload/document pipeline.
#>

function Show-MainWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$TemplateFolder,
        [string]$ToolRoot
    )

    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms | Out-Null

    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Win32Forge"
    Width="1000" Height="720"
    WindowStartupLocation="CenterScreen"
    MinWidth="700" MinHeight="520">

  <Window.Resources>
    <Style x:Key="TileBtn" TargetType="Button">
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="16,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="border" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="border" Property="Opacity" Value="0.7"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="64"/>   <!-- Header -->
      <RowDefinition Height="Auto"/> <!-- Status bar -->
      <RowDefinition Height="Auto"/> <!-- Action buttons -->
      <RowDefinition Height="*"/>    <!-- Log area -->
      <RowDefinition Height="Auto"/> <!-- Footer -->
    </Grid.RowDefinitions>

    <!-- ═══ HEADER ═══ -->
    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0.5" EndPoint="1,0.5">
          <GradientStop Color="#0693E3" Offset="0"/>
          <GradientStop Color="#9B51E0" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <Grid Margin="20,0">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="38" Height="38" Margin="0,0,12,0" VerticalAlignment="Center"
                 RenderOptions.BitmapScalingMode="HighQuality"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Win32Forge" FontSize="20" FontWeight="Light" Foreground="White"/>
            <TextBlock Text="Upload, Automate &amp; Document Win32 Apps in Intune  •  modernworkspacehub.com"
                       FontSize="11" Foreground="#D4C5F9" Margin="0,1,0,0"/>
          </StackPanel>
        </StackPanel>
        <Button x:Name="BtnSignOut" Content="Sign Out" HorizontalAlignment="Right"
                VerticalAlignment="Center" Padding="10,4" Background="#2D1B69"
                Foreground="White" BorderThickness="0" Cursor="Hand"/>
      </Grid>
    </Border>

    <!-- ═══ CONNECTION STATUS ═══ -->
    <Border Grid.Row="1" Background="#F0F0F0" BorderBrush="#DDD" BorderThickness="0,0,0,1" Padding="20,7">
      <Grid>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="ConnDot" Width="10" Height="10" Fill="#D32F2F" Margin="0,0,8,0"/>
          <TextBlock x:Name="TxtStatus" Text="Not connected" VerticalAlignment="Center" FontSize="12"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnRefresh" Content="Refresh Data" Visibility="Collapsed"
                  Padding="10,4" FontSize="12" Background="#5BA3E8" Foreground="White" BorderThickness="0"
                  Cursor="Hand" Margin="0,0,8,0"/>
          <Button x:Name="BtnConnect" Content="Connect to Intune"
                  Padding="10,4" FontSize="12" Background="#4A2B8F" Foreground="White" BorderThickness="0"
                  Cursor="Hand"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ═══ ACTION BUTTONS ═══ -->
    <Grid Grid.Row="2" Margin="16,14,16,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Button x:Name="BtnUploadSingle" Grid.Column="0" Height="72"
              Background="#5BA3E8" Foreground="White" Style="{StaticResource TileBtn}">
        <StackPanel>
          <TextBlock Text="&#xE898;" FontFamily="Segoe MDL2 Assets" FontSize="22" HorizontalAlignment="Center" Margin="0,0,0,4"/>
          <TextBlock Text="Upload App" FontSize="12"/>
        </StackPanel>
      </Button>

      <Button x:Name="BtnBulkUpload" Grid.Column="2" Height="72"
              Background="#3A2673" Foreground="White" Style="{StaticResource TileBtn}">
        <StackPanel>
          <TextBlock Text="&#xE838;" FontFamily="Segoe MDL2 Assets" FontSize="22" HorizontalAlignment="Center" Margin="0,0,0,4"/>
          <TextBlock Text="Bulk Upload" FontSize="12"/>
        </StackPanel>
      </Button>

      <Button x:Name="BtnTemplates" Grid.Column="4" Height="72"
              Background="#4A2B8F" Foreground="White" Style="{StaticResource TileBtn}">
        <StackPanel>
          <TextBlock Text="&#xE70B;" FontFamily="Segoe MDL2 Assets" FontSize="22" HorizontalAlignment="Center" Margin="0,0,0,4"/>
          <TextBlock Text="Templates" FontSize="12"/>
        </StackPanel>
      </Button>

      <Button x:Name="BtnSettings" Grid.Column="6" Height="72"
              Background="#2D1B69" Foreground="White" Style="{StaticResource TileBtn}">
        <StackPanel>
          <TextBlock Text="&#xE713;" FontFamily="Segoe MDL2 Assets" FontSize="22" HorizontalAlignment="Center" Margin="0,0,0,4"/>
          <TextBlock Text="Settings" FontSize="12"/>
        </StackPanel>
      </Button>
    </Grid>

    <!-- ═══ LOG AREA ═══ -->
    <Grid Grid.Row="3" Margin="16,12,16,0">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0" Margin="0,0,0,6">
        <TextBlock Text="Activity Log" FontWeight="SemiBold" FontSize="13" VerticalAlignment="Center"/>
        <Button x:Name="BtnClearLog" Content="Clear" HorizontalAlignment="Right"
                Padding="8,2" FontSize="11" Background="Transparent" BorderBrush="#CCC"
                Cursor="Hand"/>
      </Grid>

      <Border Grid.Row="1" BorderBrush="#CCC" BorderThickness="1" CornerRadius="3">
        <ScrollViewer x:Name="LogScroller" VerticalScrollBarVisibility="Auto">
          <TextBox x:Name="TxtLog"
                   IsReadOnly="True"
                   TextWrapping="Wrap"
                   AcceptsReturn="True"
                   BorderThickness="0"
                   Padding="8"
                   FontFamily="Consolas"
                   FontSize="12"
                   Background="Transparent"
                   VerticalAlignment="Top"/>
        </ScrollViewer>
      </Border>
    </Grid>

    <!-- ═══ FOOTER ═══ -->
    <Border Grid.Row="4" Background="#F5F5F5" BorderBrush="#DDD" BorderThickness="0,1,0,0" Padding="16,5">
      <Grid>
        <TextBlock x:Name="TxtFooter" Text="Ready" FontSize="11" Foreground="#666" VerticalAlignment="Center"/>
        <TextBlock Text="Win32Forge — Free &amp; Open Source — modernworkspacehub.com — Provided without warranty"
                   FontSize="10" Foreground="#AAA" HorizontalAlignment="Right" VerticalAlignment="Center"/>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    function Find { param($n) $window.FindName($n) }

    $connDot        = Find 'ConnDot'
    $txtStatus      = Find 'TxtStatus'
    $btnConnect     = Find 'BtnConnect'
    $btnRefresh     = Find 'BtnRefresh'
    $btnSignOut     = Find 'BtnSignOut'
    $btnUploadSingle = Find 'BtnUploadSingle'
    $btnBulkUpload  = Find 'BtnBulkUpload'
    $btnTemplates   = Find 'BtnTemplates'
    $btnSettings    = Find 'BtnSettings'
    $txtLog         = Find 'TxtLog'
    $logScroller    = Find 'LogScroller'
    $btnClearLog    = Find 'BtnClearLog'
    $txtFooter      = Find 'TxtFooter'

    # Load logo (embedded as base64 ÔÇö no dependency on Assets folder)
    $imgLogo = Find 'ImgLogo'
    try {
        $script:toolLogoBase64 = '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'
        $logoBytes  = [Convert]::FromBase64String($script:toolLogoBase64)
        $logoStream = [System.IO.MemoryStream]::new($logoBytes)
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.StreamSource = $logoStream
        $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $logoStream.Dispose()
        $imgLogo.Source = $bmp
    } catch {}
    }

    $script:connected          = $false
    $script:availableCategories = @()
    $script:availableFilters    = @()

    #region Helpers

    function Write-Log {
        param([string]$Text, [string]$Level = 'Info')
        $prefix = switch ($Level) {
            'OK'   { '[OK]   ' }
            'Warn' { '[WARN] ' }
            'Fail' { '[FAIL] ' }
            default{ '[INFO] ' }
        }
        $timestamp = Get-Date -Format 'HH:mm:ss'
        $line = "$timestamp  $prefix $Text`n"
        $txtLog.Dispatcher.Invoke([action]{
            $txtLog.AppendText($line)
            $logScroller.ScrollToEnd()
        })
    }

    function Set-Connected {
        param([string]$UserDisplay = '')
        $connDot.Fill              = [System.Windows.Media.Brushes]::Green
        $txtStatus.Text            = if ($UserDisplay) { "Connected as $UserDisplay" } else { 'Connected' }
        $txtFooter.Text            = 'Connected to Intune'
        $script:connected          = $true
        $btnRefresh.Visibility     = [System.Windows.Visibility]::Visible
    }

    function Set-Disconnected {
        $connDot.Fill              = [System.Windows.Media.Brushes]::Red
        $txtStatus.Text            = 'Not connected'
        $txtFooter.Text            = 'Not connected — click Connect to sign in'
        $script:connected          = $false
        $btnRefresh.Visibility     = [System.Windows.Visibility]::Collapsed
    }

    # Fetch categories and filters — shared by connect and the Refresh button
    function Invoke-RefreshTenantData {
        try {
            $catResp = Get-TenantGraphCollection -Url 'https://graph.microsoft.com/v1.0/deviceAppManagement/mobileAppCategories?$select=id,displayName'
            $script:availableCategories = @($catResp | ForEach-Object { $_.displayName } | Sort-Object)
            Write-Log "Loaded $($script:availableCategories.Count) app categories" 'OK'
        } catch {
            Write-Log "Could not load categories: $_" 'Warn'
        }

        try {
            # Assignment filters are beta-only — v1.0 returns 'segment not found'.
            # Fetch platform and assignmentFilterManagementType so we can restrict to
            # Windows 10 and later / Managed Devices only (matching what Intune shows
            # when assigning Win32 apps to device groups).
            $filterUrl  = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters' +
                          '?$select=id,displayName,platform,assignmentFilterManagementType'
            $filterResp = Get-TenantGraphCollection -Url $filterUrl

            $script:availableFilters = @(
                $filterResp |
                Where-Object {
                    $_.platform                       -eq 'windows10AndLater' -and
                    $_.assignmentFilterManagementType -eq 'devices'
                } |
                ForEach-Object { @{ id = $_.id; displayName = $_.displayName } }
            )

            $total    = @($filterResp).Count
            $count    = $script:availableFilters.Count
            $skipped  = $total - $count
            $skipNote = if ($skipped -gt 0) { " ($skipped non-Windows/managed-apps filters hidden)" } else { '' }
            Write-Log "Loaded $count Managed Device / Windows 10+ filter$(if ($count -ne 1) {'s'})$skipNote" 'OK'
        } catch {
            if ($_ -match '403|Forbidden') {
                Write-Log "Assignment filters not loaded — missing DeviceManagementConfiguration.Read.All permission." 'Warn'
            } else {
                Write-Log "Could not load filters: $_" 'Warn'
            }
        }
    }

    #endregion

    #region Connect / Disconnect

    $btnConnect.Add_Click({
        $txtFooter.Text = 'Connecting...'
        try {
            Import-Module IntuneWin32App -Force -ErrorAction Stop
            Import-Module MSAL.PS        -Force -ErrorAction Stop

            # Determine effective ClientID
            $effectiveClientID = if ($Config.AuthMethod -eq 'MicrosoftGraphCLI' -or -not $Config.ClientID) {
                '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
            } else {
                $Config.ClientID
            }

            # Use the module's own Connect-MSIntuneGraph so that $Global:AuthenticationHeader
            # is set in exactly the format Add-IntuneWin32App expects (ExpiresOn as UTCDateTime,
            # Authorization via CreateAuthorizationHeader()).  Bypassing it via Get-MsalToken
            # produces a subtly wrong ExpiresOn type that makes the module's token-lifetime
            # check always treat the token as expired, silently aborting the upload.
            Connect-MSIntuneGraph -TenantID $Config.TenantID -ClientID $effectiveClientID `
                                  -Interactive -ErrorAction Stop | Out-Null

            # Store for Invoke-TenantGraphRequest fallback methods
            $global:IntuneUploaderClientID  = $effectiveClientID
            $global:IntuneUploaderTenantID  = $Config.TenantID
            $global:IntuneUploaderLoginHint = ''

            # Signed-in user display name
            $userLabel = ''
            try {
                $me = Invoke-TenantGraphRequest -Url 'https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName'
                $userLabel = "$($me.displayName) ($($me.userPrincipalName))"
                $global:IntuneUploaderLoginHint = $me.userPrincipalName
            } catch {}

            Set-Connected -UserDisplay $userLabel
            Write-Log "Connected to Intune tenant: $($Config.TenantID)" 'OK'

            Invoke-RefreshTenantData
        }
        catch {
            Set-Disconnected
            Write-Log "Connection failed: $_" 'Fail'
            [System.Windows.MessageBox]::Show(
                "Could not connect to Intune:`n`n$_`n`nCheck your Tenant ID and permissions.`n`nRun Setup-Win32Forge.ps1 if you haven't already.",
                'Connection Failed', 'OK', 'Error')
        }
    })

    $btnRefresh.Add_Click({
        $txtFooter.Text = 'Refreshing...'
        Write-Log 'Refreshing categories and filters...' 'Info'
        Invoke-RefreshTenantData
        $txtFooter.Text = 'Connected to Intune'
    })

    $btnSignOut.Add_Click({
        $Global:AuthenticationHeader    = $null
        $global:IntuneUploaderClientID  = $null
        $global:IntuneUploaderTenantID  = $null
        $global:IntuneUploaderLoginHint = $null
        try { Disconnect-MSIntuneGraph -ErrorAction SilentlyContinue } catch {}
        $script:availableCategories = @()
        $script:availableFilters    = @()
        Set-Disconnected
        Write-Log 'Signed out.' 'Info'
    })

    #endregion

    #region Upload Single App

    $btnUploadSingle.Add_Click({
        if (-not $script:connected) {
            [System.Windows.MessageBox]::Show('Please connect to Intune first.', 'Not Connected', 'OK', 'Warning')
            return
        }

        $appConfig = Show-AppUploadForm `
            -TemplateFolder      $TemplateFolder `
            -DefaultOutput       $Config.DefaultOutputPath `
            -DefaultTemplate     $Config.DefaultTemplate `
            -Config              $Config `
            -AvailableCategories $script:availableCategories `
            -AvailableFilters    $script:availableFilters

        if (-not $appConfig) {
            Write-Log 'Upload cancelled by user.' 'Info'
            return
        }

        # Log everything the user selected so they can see what's about to happen
        $asgDesc = switch ($appConfig.Assignment.Type) {
            'AllDevices' { 'All Devices' }
            'AllUsers'   { 'All Users'   }
            'Group'      { "Group: $($appConfig.Assignment.GroupName)" }
            default      { 'None (manual)' }
        }
        $catDesc = if ($appConfig.Categories -and $appConfig.Categories.Count) {
            $appConfig.Categories -join ', '
        } else { 'None' }
        $filterDesc = if ($appConfig.Assignment.FilterID) { $appConfig.Assignment.FilterIntent } else { 'None' }

        Write-Log "─── Upload: $($appConfig.DisplayName) $($appConfig.Version) ───" 'Info'
        Write-Log "  Source:     $($appConfig.SourceFolder)" 'Info'
        Write-Log "  Detection:  $($appConfig.Detection.Type)" 'Info'
        Write-Log "  Assignment: $asgDesc" 'Info'
        Write-Log "  Filter:     $filterDesc" 'Info'
        Write-Log "  Categories: $catDesc" 'Info'
        Write-Log "  Template:   $($appConfig.Template)" 'Info'

        $txtFooter.Text = "Packaging: $($appConfig.DisplayName)..."
        Write-Log "Packaging .intunewin..." 'Info'

        try {
            $result = Invoke-ProcessApp -AppConfig $appConfig -Config $Config -TemplateFolder $TemplateFolder

            if ($result.Success) {
                Write-Log "Uploading to Intune... done." 'OK'
                Write-Log "$($appConfig.DisplayName) — complete  (ID: $($result.App.id))" 'OK'
                Write-Log "  Documentation: $($result.DocPath)" 'Info'
                $txtFooter.Text = "Done: $($appConfig.DisplayName)"
                [System.Windows.MessageBox]::Show(
                    "Successfully uploaded:`n$($appConfig.DisplayName) $($appConfig.Version)`n`nApp ID: $($result.App.id)`nDoc: $($result.DocPath)",
                    'Upload Complete', 'OK', 'Information')
            }
            else {
                Write-Log "FAILED: $($result.Error)" 'Fail'
                if ($result.StackTrace) { Write-Log "  Stack: $($result.StackTrace)" 'Fail' }
                $txtFooter.Text = "Failed: $($appConfig.DisplayName)"
                [System.Windows.MessageBox]::Show(
                    "Upload failed for: $($appConfig.DisplayName)`n`n$($result.Error)`n`n$($result.StackTrace)",
                    'Upload Failed', 'OK', 'Error')
            }
        }
        catch {
            Write-Log "Unexpected error: $_" 'Fail'
            $txtFooter.Text = 'Error'
        }
    })

    #endregion

    #region Bulk Upload

    $btnBulkUpload.Add_Click({
        if (-not $script:connected) {
            [System.Windows.MessageBox]::Show('Please connect to Intune first.', 'Not Connected', 'OK', 'Warning')
            return
        }

        Write-Log 'Opening Bulk Upload Manager...' 'Info'
        $txtFooter.Text = 'Bulk Upload Manager open'

        # Capture UI controls so the logger closure can write to them after
        # Show-BulkManager takes over the message loop.
        $capturedLog      = $txtLog
        $capturedScroller = $logScroller
        $bulkLogger = {
            param([string]$Text, [string]$Level)
            $prefix = switch ($Level) {
                'OK'   { '[OK]   ' }
                'Warn' { '[WARN] ' }
                'Fail' { '[FAIL] ' }
                default{ '[INFO] ' }
            }
            $line = "$(Get-Date -Format 'HH:mm:ss')  $prefix $Text`n"
            $capturedLog.AppendText($line)
            $capturedScroller.ScrollToEnd()
        }.GetNewClosure()

        Show-BulkManager `
            -Config              $Config `
            -TemplateFolder      $TemplateFolder `
            -ToolRoot            $ToolRoot `
            -AvailableCategories $script:availableCategories `
            -AvailableFilters    $script:availableFilters `
            -Logger              $bulkLogger

        $txtFooter.Text = 'Ready'
        Write-Log 'Bulk Upload Manager closed.' 'Info'
    })

    #endregion

    #region Templates

    $btnTemplates.Add_Click({
        Write-Log 'Opening Template Editor...' 'Info'
        Show-TemplateEditor `
            -TemplateFolder   $TemplateFolder `
            -AvailableFilters $script:availableFilters
        Write-Log 'Template Editor closed.' 'Info'
    })

    #endregion

    #region Settings

    $btnSettings.Add_Click({
        $configPath = Join-Path $ToolRoot 'Config\config.json'
        Write-Log 'Opening Settings...' 'Info'
        Show-SettingsEditor `
            -Config         $Config `
            -ConfigPath     $configPath `
            -TemplateFolder $TemplateFolder
        Write-Log 'Settings closed.' 'Info'
    })

    #endregion

    $btnClearLog.Add_Click({ $txtLog.Clear() })

    # Initial log entry
    Write-Log 'Win32Forge started. Click Connect to sign in.' 'Info'
    Write-Log "Tool root: $ToolRoot" 'Info'

    $window.ShowDialog() | Out-Null
}
