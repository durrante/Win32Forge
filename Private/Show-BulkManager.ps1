<#
.SYNOPSIS
    WPF Bulk Upload Manager — inline spreadsheet-style queue.

.DESCRIPTION
    Each row is a Win32 app config.  Users edit cells directly:
      - Type/paste a source folder path → PSADT metadata is auto-scanned
      - Pick a template from the dropdown → defaults are applied
      - Edit Name / Version / Publisher inline
      - "Browse Source..." opens a folder picker for the selected row
      - "Full Setup..." opens Show-AppUploadForm for detection / assignment
      - Double-click a row for the same Full Setup experience
      - Import / Export JSON, Upload Selected / All
#>

function Show-BulkManager {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$TemplateFolder,
        [string]$ToolRoot,

        [string[]]$AvailableCategories = @(),
        [object[]]$AvailableFilters    = @(),

        # Optional scriptblock called during bulk upload for main-window logging.
        # Signature: param([string]$Text, [string]$Level)  Level = Info|OK|Warn|Fail
        [scriptblock]$Logger = $null
    )

    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms | Out-Null

    # ── Template list ─────────────────────────────────────────────────────────
    $script:templateNames = @('PSADT-Default','Generic-Default')
    if ($TemplateFolder -and (Test-Path $TemplateFolder)) {
        $loaded = @(
            Get-ChildItem -Path $TemplateFolder -Filter '*.json' |
            ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) } |
            Sort-Object
        )
        if ($loaded.Count) { $script:templateNames = $loaded }
    }

    # ── XAML ──────────────────────────────────────────────────────────────────
    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Win32Forge — Bulk Upload Manager"
    Width="1500" Height="700"
    WindowStartupLocation="CenterScreen"
    MinWidth="820" MinHeight="480">

  <Window.Resources>

    <Style x:Key="ToolBtn" TargetType="Button">
      <Setter Property="Padding"         Value="10,4"/>
      <Setter Property="Margin"          Value="0,0,5,0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush"     Value="#CCC"/>
      <Setter Property="Background"      Value="White"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="3" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.82"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.65"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="PrimaryBtn" TargetType="Button" BasedOn="{StaticResource ToolBtn}">
      <Setter Property="Foreground"      Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>

    <!-- Hint text style for empty cells -->
    <Style x:Key="HintCell" TargetType="TextBlock">
      <Setter Property="Foreground"   Value="#BBB"/>
      <Setter Property="FontStyle"    Value="Italic"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Padding"      Value="4,0"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="56"/>   <!-- Header gradient -->
      <RowDefinition Height="Auto"/> <!-- Toolbar -->
      <RowDefinition Height="*"/>    <!-- DataGrid -->
      <RowDefinition Height="Auto"/> <!-- Status bar -->
    </Grid.RowDefinitions>

    <!-- ═══ HEADER ═══ -->
    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0.5" EndPoint="1,0.5">
          <GradientStop Color="#0693E3" Offset="0"/>
          <GradientStop Color="#9B51E0" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <Grid Margin="18,0">
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Bulk Upload Manager" FontSize="20" FontWeight="Light" Foreground="White"/>
          <TextBlock Text="Add rows, edit inline — Source Folder auto-scans PSADT metadata. Full Setup for detection/assignment."
                     FontSize="11" Foreground="#D4C5F9" Margin="0,1,0,0"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ═══ TOOLBAR ═══ -->
    <Border Grid.Row="1" Background="#F8F8F8" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1" Padding="10,7">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">

        <Button x:Name="BtnAddRow"    Content="+ Add Row"         Style="{StaticResource PrimaryBtn}" Background="#4A2B8F"/>

        <TextBlock Text="Default Template:" VerticalAlignment="Center" FontSize="11"
                   Foreground="#555" Margin="10,0,4,0"/>
        <ComboBox x:Name="CmbDefaultTemplate" Width="130" VerticalAlignment="Center"
                  Padding="4,3" FontSize="11"
                  ToolTip="Template applied to new rows added with '+ Add Row'"/>

        <Separator Width="1" Background="#DDD" Margin="8,2,8,2"/>

        <Button x:Name="BtnBrowse"    Content="Browse Source..."   Style="{StaticResource ToolBtn}"/>
        <Button x:Name="BtnBrowseLogo" Content="Browse Logo..."    Style="{StaticResource ToolBtn}"/>
        <Button x:Name="BtnAssignment" Content="Set Assignment ▾"  Style="{StaticResource ToolBtn}"/>
        <Button x:Name="BtnFullSetup" Content="Detection / Config..." Style="{StaticResource ToolBtn}"/>

        <Separator Width="1" Background="#DDD" Margin="4,2,10,2"/>

        <Button x:Name="BtnRemove"   Content="Remove Selected"  Style="{StaticResource ToolBtn}"/>
        <Button x:Name="BtnClear"    Content="Clear All"        Style="{StaticResource ToolBtn}"/>

        <Separator Width="1" Background="#DDD" Margin="4,2,10,2"/>

        <Button x:Name="BtnImport"   Content="Import JSON"      Style="{StaticResource ToolBtn}"/>
        <Button x:Name="BtnExport"   Content="Export JSON"      Style="{StaticResource ToolBtn}" IsEnabled="False"/>

        <Separator Width="1" Background="#DDD" Margin="4,2,10,2"/>

        <Button x:Name="BtnUploadSel" Content="Upload Selected"
                Style="{StaticResource PrimaryBtn}" Background="#5BA3E8" IsEnabled="False"/>
        <Button x:Name="BtnUploadAll" Content="Upload All"
                Style="{StaticResource PrimaryBtn}" IsEnabled="False">
          <Button.Background>
            <LinearGradientBrush StartPoint="0,0.5" EndPoint="1,0.5">
              <GradientStop Color="#0693E3" Offset="0"/>
              <GradientStop Color="#9B51E0" Offset="1"/>
            </LinearGradientBrush>
          </Button.Background>
        </Button>

      </StackPanel>
    </Border>

    <!-- ═══ DATA GRID ═══ -->
    <DataGrid x:Name="BulkGrid" Grid.Row="2"
              AutoGenerateColumns="False"
              CanUserAddRows="False"
              CanUserDeleteRows="False"
              SelectionMode="Extended"
              SelectionUnit="FullRow"
              IsReadOnly="False"
              GridLinesVisibility="Horizontal"
              HeadersVisibility="Column"
              AlternatingRowBackground="#FAFAFA"
              RowBackground="White"
              BorderThickness="0"
              ColumnHeaderHeight="30"
              RowHeight="28"
              AllowDrop="True">

      <DataGrid.ColumnHeaderStyle>
        <Style TargetType="DataGridColumnHeader">
          <Setter Property="Background"      Value="#F0EBF9"/>
          <Setter Property="Foreground"      Value="#4A2B8F"/>
          <Setter Property="FontWeight"      Value="SemiBold"/>
          <Setter Property="FontSize"        Value="12"/>
          <Setter Property="Padding"         Value="8,0"/>
          <Setter Property="BorderBrush"     Value="#DDD"/>
          <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
      </DataGrid.ColumnHeaderStyle>

      <DataGrid.CellStyle>
        <Style TargetType="DataGridCell">
          <Setter Property="Padding"         Value="0"/>
          <Setter Property="BorderThickness" Value="0"/>
          <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
              <Setter Property="Background" Value="#E8DEFF"/>
              <Setter Property="Foreground" Value="#2D1B69"/>
            </Trigger>
          </Style.Triggers>
        </Style>
      </DataGrid.CellStyle>

      <DataGrid.Columns>

        <!-- 0 — Source Folder (full path, editable) -->
        <DataGridTemplateColumn Header="Source Folder" Width="170" MinWidth="100" SortMemberPath="SourceFolder">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock VerticalAlignment="Center" Padding="6,0" TextTrimming="CharacterEllipsis">
                <TextBlock.Style>
                  <Style TargetType="TextBlock">
                    <Setter Property="Text"       Value="{Binding [SourceFolder]}"/>
                    <Setter Property="ToolTip"    Value="{Binding [SourceFolder]}"/>
                    <Setter Property="Foreground" Value="Black"/>
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding [SourceFolder]}" Value="">
                        <Setter Property="Text"       Value="e.g. D:\Apps\7-Zip\22.01  •  one app per folder"/>
                        <Setter Property="Foreground" Value="#BBB"/>
                        <Setter Property="FontStyle"  Value="Italic"/>
                        <Setter Property="ToolTip"    Value="Enter the app's own folder — not a parent folder that contains multiple apps. All files in the folder are packaged together."/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </TextBlock.Style>
              </TextBlock>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
          <DataGridTemplateColumn.CellEditingTemplate>
            <DataTemplate>
              <TextBox Text="{Binding [SourceFolder], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                       VerticalAlignment="Center" Padding="5,0" BorderThickness="0" Background="Transparent"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellEditingTemplate>
        </DataGridTemplateColumn>

        <!-- 1 — Template (ComboBox) — second column so the most important workflow choice is immediately visible -->
        <DataGridTemplateColumn Header="Template" Width="120" MinWidth="80" SortMemberPath="Template">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding [Template]}" VerticalAlignment="Center" Padding="6,0"
                         TextTrimming="CharacterEllipsis"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
          <DataGridTemplateColumn.CellEditingTemplate>
            <DataTemplate>
              <ComboBox x:Name="CmbTemplate"
                        SelectedItem="{Binding [Template], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                        VerticalAlignment="Center" BorderThickness="0" Padding="4,0"
                        Background="Transparent"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellEditingTemplate>
        </DataGridTemplateColumn>

        <!-- 2 — Display Name -->
        <DataGridTextColumn Header="Display Name" Binding="{Binding [DisplayName]}"
                            Width="150" MinWidth="80" SortMemberPath="DisplayName">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 2 — Version -->
        <DataGridTextColumn Header="Version" Binding="{Binding [Version]}"
                            Width="80" SortMemberPath="Version">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 3 — Publisher -->
        <DataGridTextColumn Header="Publisher" Binding="{Binding [Publisher]}"
                            Width="110" SortMemberPath="Publisher">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 4 — Setup File (ComboBox populated from source folder) -->
        <DataGridTemplateColumn Header="Setup File" Width="120" MinWidth="60" SortMemberPath="SetupFile">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding [SetupFile]}" VerticalAlignment="Center" Padding="6,0"
                         TextTrimming="CharacterEllipsis" ToolTip="{Binding [SetupFile]}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
          <DataGridTemplateColumn.CellEditingTemplate>
            <DataTemplate>
              <ComboBox x:Name="CmbSetupFile"
                        Text="{Binding [SetupFile], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                        IsEditable="True"
                        VerticalAlignment="Center" BorderThickness="0" Padding="4,0"
                        Background="Transparent"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellEditingTemplate>
        </DataGridTemplateColumn>

        <!-- 5 — Install Command -->
        <DataGridTextColumn Header="Install Cmd" Binding="{Binding [InstallCmd]}"
                            Width="140" MinWidth="80" SortMemberPath="InstallCmd">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [InstallCmd]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 6 — Uninstall Command -->
        <DataGridTextColumn Header="Uninstall Cmd" Binding="{Binding [UninstallCmd]}"
                            Width="140" MinWidth="80" SortMemberPath="UninstallCmd">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [UninstallCmd]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 8 — Category (ComboBox) -->
        <DataGridTemplateColumn Header="Category" Width="110" MinWidth="70" SortMemberPath="Category">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding [Category]}" VerticalAlignment="Center" Padding="6,0"
                         TextTrimming="CharacterEllipsis"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
          <DataGridTemplateColumn.CellEditingTemplate>
            <DataTemplate>
              <ComboBox x:Name="CmbCategory"
                        SelectedItem="{Binding [Category], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                        VerticalAlignment="Center" BorderThickness="0" Padding="4,0"
                        Background="Transparent"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellEditingTemplate>
        </DataGridTemplateColumn>

        <!-- 9 — Detection (read-only summary) -->
        <DataGridTemplateColumn Header="Detection" Width="100" IsReadOnly="True" SortMemberPath="Detection">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock VerticalAlignment="Center" Padding="6,0" TextTrimming="CharacterEllipsis">
                <TextBlock.Style>
                  <Style TargetType="TextBlock">
                    <Setter Property="Text"       Value="{Binding [Detection]}"/>
                    <Setter Property="Foreground"  Value="#444"/>
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding [Detection]}" Value="—">
                        <Setter Property="Foreground" Value="#F59E0B"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding [Detection]}" Value="">
                        <Setter Property="Foreground" Value="#F59E0B"/>
                        <Setter Property="Text"       Value="Not set"/>
                        <Setter Property="FontStyle"  Value="Italic"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </TextBlock.Style>
              </TextBlock>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>

        <!-- 10 — Assignment (read-only summary) -->
        <DataGridTemplateColumn Header="Assignment" Width="140" MinWidth="100" IsReadOnly="True" SortMemberPath="Assignment">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding [Assignment]}" VerticalAlignment="Center" Padding="6,0"
                         TextTrimming="CharacterEllipsis" Foreground="#444"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>

        <!-- 11 — Description -->
        <DataGridTextColumn Header="Description" Binding="{Binding [Description]}"
                            Width="110" MinWidth="60" SortMemberPath="Description">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [Description]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 12 — Info URL -->
        <DataGridTextColumn Header="Info URL" Binding="{Binding [InformationURL]}"
                            Width="100" MinWidth="60" SortMemberPath="InformationURL">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [InformationURL]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 13 — Privacy URL -->
        <DataGridTextColumn Header="Privacy URL" Binding="{Binding [PrivacyURL]}"
                            Width="100" MinWidth="60" SortMemberPath="PrivacyURL">
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [PrivacyURL]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 14 — Logo Path (drag an image file onto a row to auto-fill) -->
        <DataGridTextColumn Width="100" MinWidth="60" SortMemberPath="LogoPath"
                            Binding="{Binding [LogoPath]}">
          <DataGridTextColumn.Header>
            <TextBlock Text="Logo Path"
                       ToolTip="PNG, JPG, or JPEG only — drag and drop an image file onto the row, or browse using the toolbar button."/>
          </DataGridTextColumn.Header>
          <DataGridTextColumn.ElementStyle>
            <Style TargetType="TextBlock">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="6,0"/>
              <Setter Property="TextTrimming"      Value="CharacterEllipsis"/>
              <Setter Property="ToolTip"           Value="{Binding [LogoPath]}"/>
            </Style>
          </DataGridTextColumn.ElementStyle>
          <DataGridTextColumn.EditingElementStyle>
            <Style TargetType="TextBox">
              <Setter Property="VerticalAlignment" Value="Center"/>
              <Setter Property="Padding"           Value="5,0"/>
              <Setter Property="BorderThickness"   Value="0"/>
              <Setter Property="Background"        Value="Transparent"/>
            </Style>
          </DataGridTextColumn.EditingElementStyle>
        </DataGridTextColumn>

        <!-- 15 — Status (read-only, colour-coded) -->
        <DataGridTemplateColumn Header="Status" Width="*" MinWidth="80" IsReadOnly="True" SortMemberPath="Status">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <TextBlock Text="{Binding [Status]}" FontWeight="SemiBold"
                         VerticalAlignment="Center" Padding="6,0">
                <TextBlock.Style>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#888"/>
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding [Status]}" Value="Done">
                        <Setter Property="Foreground" Value="#2E7D32"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding [Status]}" Value="Failed">
                        <Setter Property="Foreground" Value="#C62828"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding [Status]}" Value="Uploading...">
                        <Setter Property="Foreground" Value="#1565C0"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </TextBlock.Style>
              </TextBlock>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>

      </DataGrid.Columns>
    </DataGrid>

    <!-- ═══ STATUS BAR ═══ -->
    <Border Grid.Row="3" Background="#F5F5F5" BorderBrush="#DDD" BorderThickness="0,1,0,0" Padding="14,5">
      <Grid>
        <TextBlock x:Name="TxtStatus" FontSize="11" Foreground="#555" VerticalAlignment="Center"
                   Text="0 apps in queue  •  Click + Add Row to begin  •  Double-click a row for full config  •  Drag a logo image onto a row to set its logo"/>
        <TextBlock x:Name="TxtUploadResult" HorizontalAlignment="Right"
                   FontSize="11" FontWeight="SemiBold" Foreground="#4A2B8F" VerticalAlignment="Center"/>
      </Grid>
    </Border>

  </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    function Find { param($n) $window.FindName($n) }

    $bulkGrid            = Find 'BulkGrid'
    $global:IntuneUploaderGrid = $bulkGrid  # global ref so Dispatcher [System.Action] lambdas can reach it
    $btnAddRow           = Find 'BtnAddRow'
    $cmbDefaultTemplate  = Find 'CmbDefaultTemplate'
    $btnBrowse           = Find 'BtnBrowse'
    $btnBrowseLogo   = Find 'BtnBrowseLogo'
    $btnAssignment   = Find 'BtnAssignment'
    $btnFullSetup    = Find 'BtnFullSetup'
    $btnRemove       = Find 'BtnRemove'
    $btnClear        = Find 'BtnClear'
    $btnImport       = Find 'BtnImport'
    $btnExport       = Find 'BtnExport'
    $btnUploadSel    = Find 'BtnUploadSel'
    $btnUploadAll    = Find 'BtnUploadAll'
    $txtStatus       = Find 'TxtStatus'
    $txtUploadResult = Find 'TxtUploadResult'

    # ── Data structures ───────────────────────────────────────────────────────
    # $script:bmRows  — List of hashtables; full AppConfig + _id + _status
    # $script:bmTable — DataTable; parallel display data; never rebuilt (updated in-place)
    $script:bmRows  = [System.Collections.Generic.List[hashtable]]::new()
    $script:bmTable = New-Object System.Data.DataTable

    # _RowIndex (integer, hidden) tracks insertion order for the default sort
    $script:bmTable.Columns.Add('_RowIndex', [int]) | Out-Null
    foreach ($col in @('_Id','SourceFolder','DisplayName','Version',
                        'Publisher','Description','Category','SetupFile',
                        'InstallCmd','UninstallCmd','InformationURL','PrivacyURL',
                        'LogoPath','Template','Detection','Assignment','Status')) {
        $script:bmTable.Columns.Add($col, [string]) | Out-Null
    }
    $bulkGrid.ItemsSource = $script:bmTable.DefaultView
    $script:bmTable.DefaultView.Sort = '_RowIndex ASC'

    # Populate default-template ComboBox; honour the config default if present
    foreach ($t in $script:templateNames) { $cmbDefaultTemplate.Items.Add($t) | Out-Null }
    $cfgDefault = $Config.DefaultTemplate
    $defIdx = if ($cfgDefault) { $cmbDefaultTemplate.Items.IndexOf($cfgDefault) } else { -1 }
    $cmbDefaultTemplate.SelectedIndex = if ($defIdx -ge 0) { $defIdx } else { 0 }

    # ─────────────────────────────────────────────────────────────────────────
    #region Helper functions
    # ─────────────────────────────────────────────────────────────────────────

    function Get-DetectionSummary {
        param($Det)
        if (-not $Det) { return '—' }
        $d = if ($Det -is [hashtable]) { $Det } else {
            $h = @{}; $Det.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h
        }
        switch ($d.Type) {
            'Script'   { 'Script' }
            'Registry' { "Reg: $(Split-Path ($d.KeyPath ?? '') -Leaf)" }
            'MSI'      { "MSI: $($d.ProductCode -replace '^\{|\}$','')" }
            'File'     { "File: $($d.FileOrFolder)" }
            default    { $d.Type ?? '—' }
        }
    }

    function Get-AssignmentSummary {
        param($Asg)
        if (-not $Asg) { return 'Not set' }
        $a = if ($Asg -is [hashtable]) { $Asg }
             elseif ($Asg -is [PSCustomObject]) {
                 $h = @{}; $Asg.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h
             } else { @{} }
        $intent = if ($a.Intent) { " — $($a.Intent)" } else { '' }
        switch ($a.Type) {
            'AllDevices' { "All Devices$intent" }
            'AllUsers'   { "All Users$intent" }
            'Group' {
                # Support new Groups array and old scalar GroupName
                if ($a.Groups -and @($a.Groups).Count -gt 0) {
                    $names = @($a.Groups | ForEach-Object {
                        if ($_ -is [hashtable]) { $_.GroupName ?? $_.GroupID ?? '?' }
                        else { $_.GroupName ?? $_.GroupID ?? '?' }
                    })
                    $n = $names.Count
                    if ($n -eq 1) { "Group: $($names[0])$intent" }
                    else           { "$n Groups$intent" }
                } else {
                    "Group: $($a.GroupName ?? $a.GroupID ?? '?')$intent"
                }
            }
            'None'  { 'None' }
            default { $a.Type ?? 'Not set' }
        }
    }

    function Find-RowById {
        param([string]$Id)
        for ($i = 0; $i -lt $script:bmRows.Count; $i++) {
            if ($script:bmRows[$i]._id -eq $Id) { return $i }
        }
        return -1
    }

    function Find-DataRow {
        param([string]$Id)
        foreach ($dr in $script:bmTable.Rows) {
            if ($dr['_Id'] -eq $Id) { return $dr }
        }
        return $null
    }

    function Refresh-StatusBar {
        $total  = $script:bmRows.Count
        $done   = @($script:bmRows | Where-Object { $_._status -eq 'Done'   }).Count
        $failed = @($script:bmRows | Where-Object { $_._status -eq 'Failed' }).Count
        $plural = if ($total -ne 1) { 's' } else { '' }

        $parts = @("$total app$plural in queue")
        if ($done)   { $parts += "$done done" }
        if ($failed) { $parts += "$failed failed" }
        if ($total -eq 0) {
            $parts += 'Click + Add Row to begin'
        }
        $parts += 'Double-click a row for full config'
        $parts += 'Drag a logo image onto a row to set it'
        $txtStatus.Text = $parts -join '  •  '

        $hasRows = $total -gt 0
        $btnExport.IsEnabled   = $hasRows
        $btnUploadAll.IsEnabled = $hasRows
        $btnUploadSel.IsEnabled = $hasRows
    }

    function Update-RowStatus {
        param([string]$Id, [string]$Status)
        $idx = Find-RowById -Id $Id
        if ($idx -ge 0) { $script:bmRows[$idx]._status = $Status }
        $dr = Find-DataRow -Id $Id
        if ($dr) { $dr['Status'] = $Status }
    }

    # Walk the WPF visual tree to find a child of a given type
    # Needed for DataGridTemplateColumn where $e.EditingElement is a ContentPresenter wrapper
    function Find-VisualChild {
        param($Parent, [Type]$TargetType)
        if (-not $Parent) { return $null }
        $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
        for ($i = 0; $i -lt $count; $i++) {
            $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
            if ($child -is $TargetType) { return $child }
            $found = Find-VisualChild -Parent $child -TargetType $TargetType
            if ($found) { return $found }
        }
        return $null
    }

    # Called after a source folder is set — scans PSADT metadata and updates the row in-place
    function Invoke-SourceScan {
        param([string]$Id, [string]$Path)

        $idx = Find-RowById -Id $Id
        if ($idx -lt 0) { return }

        $dr = Find-DataRow -Id $Id
        $row = $script:bmRows[$idx]

        # Always update the source folder display
        if ($dr) { $dr['SourceFolder'] = $Path }
        $row.SourceFolder = $Path

        if (-not $Path -or -not (Test-Path $Path -PathType Container)) {
            Refresh-StatusBar
            return
        }

        # Try PSADT scan
        $meta = $null
        try { $meta = Get-PSADTMetadata -SourceFolder $Path } catch {}

        if ($meta) {
            $row.IsPSADT   = $true
            $row.SetupFile = $meta.SetupFile

            if (-not $row.DisplayName) { $row.DisplayName = $meta.AppName }
            if (-not $row.Version)     { $row.Version     = $meta.AppVersion }
            if (-not $row.Publisher)   { $row.Publisher   = $meta.AppVendor }
            if (-not $row.Owner)       { $row.Owner       = $meta.AppOwner }
            if (-not $row.Notes)       { $row.Notes       = "PSADT v4 package ($($meta.AppName))" }

            # Switch to PSADT template if still on a non-PSADT one; commands come from template, not metadata
            $currentTpl = $row.Template ?? ($cmbDefaultTemplate.SelectedItem -as [string])
            $tplIsPSADT = $false
            if ($currentTpl -and $TemplateFolder) {
                $tp = Join-Path $TemplateFolder "$currentTpl.json"
                if (Test-Path $tp) { try { $tplIsPSADT = [bool](Get-Content $tp -Raw | ConvertFrom-Json).IsPSADT } catch {} }
            }
            if (-not $tplIsPSADT) {
                $psadtTpl = if ($Config.DefaultTemplate -and $Config.DefaultTemplate -like 'PSADT-*') {
                    $Config.DefaultTemplate
                } else { 'PSADT-Default' }
                $row.Template = $psadtTpl
            }
            # Apply template to get correct commands (Force so they come from template, not stale metadata)
            Apply-TemplateToRow -Id $Id -TemplateName $row.Template -Force
        }
        else {
            $row.IsPSADT = $false

            # If current template is PSADT, offer to switch to a non-PSADT one
            $currentTpl = $row.Template ?? ($cmbDefaultTemplate.SelectedItem -as [string])
            $tplIsPSADT = $false
            if ($currentTpl -and $TemplateFolder) {
                $tp = Join-Path $TemplateFolder "$currentTpl.json"
                if (Test-Path $tp) { try { $tplIsPSADT = [bool](Get-Content $tp -Raw | ConvertFrom-Json).IsPSADT } catch {} }
            }
            if ($tplIsPSADT) {
                $nonPsadtTpls = @(Get-ChildItem $TemplateFolder -Filter '*.json' -ErrorAction SilentlyContinue | ForEach-Object {
                    try { $d = Get-Content $_.FullName -Raw | ConvertFrom-Json; if (-not $d.IsPSADT) { $_.BaseName } } catch {}
                } | Where-Object { $_ })
                if ($nonPsadtTpls.Count -gt 0) {
                    $suggestTpl = $nonPsadtTpls[0]
                    $ans = [System.Windows.MessageBox]::Show(
                        "This folder does not contain a PSADT package.`n`n" +
                        "The selected template ('$currentTpl') is configured for PSADT.`n`n" +
                        "Switch to '$suggestTpl'?",
                        'Non-PSADT Folder Detected', 'YesNo', 'Question')
                    if ($ans -eq 'Yes') {
                        $row.Template = $suggestTpl
                        Apply-TemplateToRow -Id $Id -TemplateName $suggestTpl -Force
                    }
                }
            }

            # Auto-set setup file from first installer found in root of source folder
            if (-not $row.SetupFile) {
                $firstInstaller = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -match '^\.(exe|msi)$' } |
                    Sort-Object { $_.Extension -eq '.msi' } -Descending |
                    Select-Object -First 1
                if ($firstInstaller) {
                    $row.SetupFile = $firstInstaller.Name
                    Set-RowCommandSuggestion -Id $Id -FileName $firstInstaller.Name
                }
            }
        }

        # Collect auto-detections first — show prompts after both are resolved
        $detScript = $null
        $logoFound  = $null

        if (-not $row.Detection) {
            $detScript = Get-ChildItem -Path $Path -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -match 'detection' } |
                         Select-Object -First 1
        }

        # Logo: root of source folder only (PNG/JPG/JPEG only — Intune does not accept other formats)
        if (-not $row.LogoPath) {
            $logoFiles = @(Get-ChildItem -Path $Path -Filter '*.png'  -ErrorAction SilentlyContinue) +
                         @(Get-ChildItem -Path $Path -Filter '*.jpg'  -ErrorAction SilentlyContinue) +
                         @(Get-ChildItem -Path $Path -Filter '*.jpeg' -ErrorAction SilentlyContinue)
            $logoFiles = @($logoFiles | Where-Object { $_ })
            if ($logoFiles.Count -gt 0) { $logoFound = $logoFiles[0].FullName }
        }

        # Apply both to the row
        if ($detScript) {
            $row.Detection = @{
                Type                  = 'Script'
                ScriptPath            = $detScript.FullName
                EnforceSignatureCheck = $false
                RunAs32Bit            = $false
            }
            if ($dr) { $dr['Detection'] = Get-DetectionSummary -Det $row.Detection }
        }
        if ($logoFound) { $row.LogoPath = $logoFound }

        # Show auto-set prompts — one combined message if both found, individual otherwise
        if ($detScript -and $logoFound) {
            [System.Windows.MessageBox]::Show(
                "Two settings were auto-detected and applied:`n`n" +
                "  Detection script : $($detScript.Name)`n" +
                "  Logo             : $(Split-Path $logoFound -Leaf)`n`n" +
                "Please confirm or adjust these via the row config before uploading.",
                'Auto-Detected Settings', 'OK', 'Information')
        } elseif ($detScript) {
            [System.Windows.MessageBox]::Show(
                "Detection script auto-detected:`n  $($detScript.Name)`n`n" +
                "This has been set as the detection method for this app.`n" +
                "Please confirm or change it via 'Detection / Config...' before uploading.",
                'Detection Auto-Set', 'OK', 'Information')
        } elseif ($logoFound) {
            [System.Windows.MessageBox]::Show(
                "Logo auto-detected:`n  $(Split-Path $logoFound -Leaf)`n`n" +
                "This has been set as the logo for this app.`n" +
                "Please confirm or change it via 'Config...' before uploading.",
                'Logo Auto-Set', 'OK', 'Information')
        }

        if ($dr) {
            $dr['DisplayName']   = $row.DisplayName         ?? ''
            $dr['Version']       = $row.Version             ?? ''
            $dr['Publisher']     = $row.Publisher           ?? ''
            $dr['Description']   = $row.Description         ?? ''
            $dr['SetupFile']     = $row.SetupFile           ?? ''
            $dr['InstallCmd']    = $row.InstallCommandLine   ?? ''
            $dr['UninstallCmd']  = $row.UninstallCommandLine ?? ''
            $dr['LogoPath']      = $row.LogoPath            ?? ''
            $dr['Template']      = $row.Template            ?? ''
            $dr['Detection']     = Get-DetectionSummary -Det $row.Detection
        }

        Refresh-StatusBar
    }

    # Apply template defaults to a row's bmRows entry.
    # $Force = $true when the user explicitly picks a new template — overwrites assignment and arch.
    # $Force = $false (default) on initial load — only fills fields that are blank.
    function Apply-TemplateToRow {
        param([string]$Id, [string]$TemplateName, [switch]$Force)

        $idx = Find-RowById -Id $Id
        if ($idx -lt 0 -or -not $TemplateName) { return }

        $tplPath = Join-Path $TemplateFolder "$TemplateName.json"
        if (-not (Test-Path $tplPath)) { return }

        try {
            $tpl = Get-Content $tplPath -Raw | ConvertFrom-Json
            # Guard: if the file was accidentally saved as a JSON array (e.g. [true, {...}]),
            # unwrap it to the first PSCustomObject element so property access works correctly.
            if ($tpl -is [array]) {
                $tpl = $tpl | Where-Object { $_ -is [PSCustomObject] } | Select-Object -First 1
            }
            $row = $script:bmRows[$idx]

            # Architecture / MinOS
            if ($tpl.Architecture -and ($Force -or -not $row.Architecture)) {
                $row.Architecture = $tpl.Architecture
            }
            if ($tpl.MinimumSupportedWindowsRelease -and ($Force -or -not $row.MinimumSupportedWindowsRelease)) {
                $row.MinimumSupportedWindowsRelease = $tpl.MinimumSupportedWindowsRelease
            }

            # Install experience / restart behaviour
            if ($tpl.InstallExperience -and ($Force -or -not $row.InstallExperience)) {
                $row.InstallExperience = $tpl.InstallExperience
            }
            if ($tpl.RestartBehavior -and ($Force -or -not $row.RestartBehavior)) {
                $row.RestartBehavior = $tpl.RestartBehavior
            }
            if ($tpl.MaximumInstallationTimeInMinutes -and ($Force -or -not $row.MaximumInstallationTimeInMinutes)) {
                $row.MaximumInstallationTimeInMinutes = $tpl.MaximumInstallationTimeInMinutes
            }
            if ($null -ne $tpl.AllowAvailableUninstall -and ($Force -or $null -eq $row.AllowAvailableUninstall)) {
                $row.AllowAvailableUninstall = $tpl.AllowAvailableUninstall
            }
            if ($tpl.ReturnCodes -and ($Force -or -not $row.ReturnCodes)) {
                $row.ReturnCodes = $tpl.ReturnCodes
            }

            # Command lines:
            #   Non-PSADT — only fill blanks; admin should set these explicitly, never auto-overwrite.
            #   PSADT — template commands are PSADT-specific args (e.g. -DeployMode Auto vs Silent),
            #           so apply when forced (user explicitly switched template) or when blank.
            if ($tpl.InstallCommandLine) {
                if ($Force -or -not $row.InstallCommandLine) { $row.InstallCommandLine = $tpl.InstallCommandLine }
            }
            if ($tpl.UninstallCommandLine) {
                if ($Force -or -not $row.UninstallCommandLine) { $row.UninstallCommandLine = $tpl.UninstallCommandLine }
            }

            # Assignment: always apply when forced (user picked a template for a reason);
            # on initial load only fill if not yet set
            if ($tpl.Assignment -and ($Force -or -not $row.Assignment)) {
                $h = if ($tpl.Assignment -is [PSCustomObject]) {
                    $ht = @{}
                    $tpl.Assignment.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
                    $ht
                } else { $tpl.Assignment }
                $row.Assignment = $h
            }

            if ($tpl.Notes -and -not $row.Notes) { $row.Notes = $tpl.Notes }

            # Sync DataRow columns that may have been updated by the template
            $dr = Find-DataRow -Id $Id
            if ($dr) {
                $dr['InstallCmd']   = $row.InstallCommandLine   ?? ''
                $dr['UninstallCmd'] = $row.UninstallCommandLine ?? ''
                $dr['Assignment']   = Get-AssignmentSummary -Asg $row.Assignment
            }
        }
        catch { }
    }

    # Applies MSI/EXE command suggestions to a row based on the setup file extension.
    # Always overwrites commands (mirrors -Force behaviour from the ad-hoc form).
    function Set-RowCommandSuggestion {
        param([string]$Id, [string]$FileName)
        $idx = Find-RowById -Id $Id
        if ($idx -lt 0 -or -not $FileName) { return }
        $row = $script:bmRows[$idx]
        $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
        $dr  = Find-DataRow -Id $Id
        $warnExe = $false

        if ($ext -eq '.msi') {
            $row.InstallCommandLine   = "msiexec /i `"$FileName`" /qn /norestart"
            $row.UninstallCommandLine = "msiexec /x `"$FileName`" /qn /norestart"
        } elseif ($ext -eq '.exe') {
            $row.InstallCommandLine   = "`"$FileName`""
            $row.UninstallCommandLine = "`"$FileName`""
            $warnExe = $true
        } else { return }

        if ($dr) {
            $dr['InstallCmd']   = $row.InstallCommandLine
            $dr['UninstallCmd'] = $row.UninstallCommandLine
        }
        if ($warnExe) {
            [System.Windows.MessageBox]::Show(
                "EXE installer selected — Install and Uninstall commands have been pre-filled with the filename only.`n`n" +
                "Please add the appropriate silent switches (e.g. /S, /quiet, /silent) before uploading.`n" +
                "The uninstall command will also need to be updated with the correct path and switch.",
                'Update Commands Required', 'OK', 'Warning') | Out-Null
        }
    }

    # Add a new row to both DataTable and bmRows
    function Add-BmRow {
        param([hashtable]$Config = @{})

        $id      = [guid]::NewGuid().ToString()
        $defTpl  = $Config.DefaultTemplate ?? ($cmbDefaultTemplate.SelectedItem -as [string]) ?? 'PSADT-Default'

        $newRow = $Config.Clone()
        if (-not $newRow._id)     { $newRow._id     = $id }
        if (-not $newRow._status) { $newRow._status = 'Pending' }
        if (-not $newRow.Template){ $newRow.Template = $defTpl }
        # Do NOT hardcode a default assignment here — let Apply-TemplateToRow
        # pull assignment (and arch/commands) from the selected template below.
        # Only fall back to AllDevices if there is no template to load from.
        $script:bmRows.Add($newRow)

        $dr = $script:bmTable.NewRow()
        $dr['_RowIndex']      = $script:bmTable.Rows.Count
        $dr['_Id']            = $newRow._id
        $dr['SourceFolder']   = $newRow.SourceFolder        ?? ''
        $dr['DisplayName']    = $newRow.DisplayName         ?? ''
        $dr['Version']        = $newRow.Version             ?? ''
        $dr['Publisher']      = $newRow.Publisher           ?? ''
        $dr['Description']    = $newRow.Description         ?? ''
        $dr['Category']       = $newRow.Category            ?? ''
        $dr['SetupFile']      = $newRow.SetupFile           ?? ''
        $dr['InstallCmd']     = $newRow.InstallCommandLine  ?? ''
        $dr['UninstallCmd']   = $newRow.UninstallCommandLine ?? ''
        $dr['InformationURL'] = $newRow.InformationURL      ?? ''
        $dr['PrivacyURL']     = $newRow.PrivacyURL          ?? ''
        $dr['LogoPath']       = $newRow.LogoPath            ?? ''
        $dr['Template']       = $newRow.Template            ?? ''
        $dr['Detection']      = Get-DetectionSummary  -Det $newRow.Detection
        $dr['Assignment']     = Get-AssignmentSummary -Asg $newRow.Assignment
        $dr['Status']         = $newRow._status
        $script:bmTable.Rows.Add($dr)

        # Apply template defaults (assignment, arch, commands) now that the row exists.
        # $Force=$false so pre-populated fields from $Config are not overwritten.
        Apply-TemplateToRow -Id $id -TemplateName $newRow.Template

        # If the template had no assignment definition, set a safe default
        if (-not $script:bmRows[(Find-RowById -Id $id)].Assignment) {
            $script:bmRows[(Find-RowById -Id $id)].Assignment = @{ Type = 'AllDevices'; Intent = 'required'; Notification = 'showAll' }
            $dr2 = Find-DataRow -Id $id
            if ($dr2) { $dr2['Assignment'] = Get-AssignmentSummary -Asg $script:bmRows[(Find-RowById -Id $id)].Assignment }
        }

        Refresh-StatusBar
        return $newRow._id
    }

    #endregion

    # ─────────────────────────────────────────────────────────────────────────
    #region Grid events
    # ─────────────────────────────────────────────────────────────────────────

    # Set ComboBox ItemsSource when Template column enters edit mode.
    # $e.EditingElement for a DataGridTemplateColumn is a ContentPresenter wrapper,
    # so we walk the visual tree to reach the actual ComboBox.
    $bulkGrid.Add_PreparingCellForEdit({
        param($s, $e)
        $hdr = $e.Column.Header -as [string]
        if ($hdr -eq 'Template') {
            $cmb = Find-VisualChild -Parent $e.EditingElement `
                                    -TargetType ([System.Windows.Controls.ComboBox])
            if ($cmb) {
                $cmb.ItemsSource = $script:templateNames
                $cmb.Add_DropDownClosed({
                    $global:IntuneUploaderGrid.Dispatcher.BeginInvoke(
                        [System.Action]{ $global:IntuneUploaderGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) },
                        [System.Windows.Threading.DispatcherPriority]::Background)
                })
            }
        }
        elseif ($hdr -eq 'Category') {
            $cmb = Find-VisualChild -Parent $e.EditingElement `
                                    -TargetType ([System.Windows.Controls.ComboBox])
            if ($cmb) {
                $catList = [System.Collections.Generic.List[string]]::new()
                $catList.Add('')
                foreach ($c in ($AvailableCategories | Sort-Object)) { $catList.Add($c) }
                $cmb.ItemsSource = $catList
                $cmb.Add_DropDownClosed({
                    $global:IntuneUploaderGrid.Dispatcher.BeginInvoke(
                        [System.Action]{ $global:IntuneUploaderGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) },
                        [System.Windows.Threading.DispatcherPriority]::Background)
                })
            }
        }
        elseif ($hdr -eq 'Setup File') {
            $cmb = Find-VisualChild -Parent $e.EditingElement `
                                    -TargetType ([System.Windows.Controls.ComboBox])
            if ($cmb) {
                $rowView    = $e.Row.Item -as [System.Data.DataRowView]
                $capturedId = if ($rowView) { $rowView['_Id'] -as [string] } else { $null }
                $rowIdx     = if ($capturedId) { Find-RowById -Id $capturedId } else { -1 }
                $srcFolder  = if ($rowIdx -ge 0) { $script:bmRows[$rowIdx].SourceFolder } else { $null }

                $fileList = [System.Collections.Generic.List[string]]::new()
                if ($srcFolder -and (Test-Path $srcFolder)) {
                    $installers = @(Get-ChildItem -Path $srcFolder -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Extension -match '^\.(exe|msi|cmd|bat|ps1)$' } | Sort-Object Name)
                    $others = @(Get-ChildItem -Path $srcFolder -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Extension -notmatch '^\.(exe|msi|cmd|bat|ps1)$' } | Sort-Object Name)
                    foreach ($f in ($installers + $others)) { $fileList.Add($f.Name) }
                }
                $cmb.ItemsSource = $fileList

                # Pre-stash the row ID now (PreparingCellForEdit scope) so the handler
                # below can read it without needing GetNewClosure.
                $script:pendingSetupFileId = $capturedId

                # No GetNewClosure — $global:IntuneUploaderGrid is reachable from any scope;
                # SelectedItem is read from $sender so $cmb capture is not needed.
                # CommitEdit(Cell) fires CellEditEnding which updates commands in the DataTable,
                # then CommitEdit(Row) forces the DataGrid to redraw all cells from the DataTable
                # so the updated Install/Uninstall commands appear immediately.
                $cmb.Add_DropDownClosed({
                    param($sndr, $eargs)
                    $sf = ($sndr -as [System.Windows.Controls.ComboBox]).SelectedItem -as [string]
                    if ($sf) {
                        $script:pendingSetupFile = $sf
                    }
                    $global:IntuneUploaderGrid.Dispatcher.BeginInvoke(
                        [System.Action]{
                            $global:IntuneUploaderGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
                            $global:IntuneUploaderGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,  $true)
                        },
                        [System.Windows.Threading.DispatcherPriority]::Background)
                })
            }
        }
    })

    # 3-state column sort: A→Z  →  Z→A  →  (third click) reset to insertion order
    $bulkGrid.Add_Sorting({
        param($s, $e)
        if ($e.Column.SortDirection -eq [System.ComponentModel.ListSortDirection]::Descending) {
            # Third click: cancel WPF sort and revert to _RowIndex order
            $e.Handled = $true
            foreach ($c in $bulkGrid.Columns) { $c.SortDirection = $null }
            $script:bmTable.DefaultView.Sort = '_RowIndex ASC'
        }
        # else: let WPF handle Ascending → Descending automatically
    })

    # ── Logo drag-and-drop ────────────────────────────────────────────────────
    # Dragging an image file directly onto a row auto-fills the Logo Path column.
    $bulkGrid.Add_DragOver({
        param($s, $e)
        $e.Effects = [System.Windows.DragDropEffects]::None
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $files = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop))
            if (@($files | Where-Object { $_ -match '\.(png|jpg|jpeg)$' }).Count -gt 0) {
                $e.Effects = [System.Windows.DragDropEffects]::Copy
            }
        }
        $e.Handled = $true
    })

    $bulkGrid.Add_Drop({
        param($s, $e)
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
        $imgFiles = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) |
                      Where-Object { $_ -match '\.(png|jpg|jpeg)$' })
        if ($imgFiles.Count -eq 0) { return }
        $logoFile = $imgFiles[0]

        # Walk the hit-test result up to the DataGridRow that was dropped on
        $dep = $e.OriginalSource -as [System.Windows.DependencyObject]
        while ($dep -and $dep -isnot [System.Windows.Controls.DataGridRow]) {
            $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
        }

        $targetId = $null
        if ($dep -is [System.Windows.Controls.DataGridRow]) {
            $item = $dep.Item -as [System.Data.DataRowView]
            if ($item) { $targetId = $item['_Id'] -as [string] }
        }

        # Fall back to the currently selected row if the drop landed on empty space
        if (-not $targetId) {
            $sel = $bulkGrid.SelectedItem -as [System.Data.DataRowView]
            if ($sel) { $targetId = $sel['_Id'] -as [string] }
        }

        if ($targetId) {
            $idx = Find-RowById -Id $targetId
            if ($idx -ge 0) {
                $script:bmRows[$idx].LogoPath = $logoFile
                $dr = Find-DataRow -Id $targetId
                if ($dr) { $dr['LogoPath'] = $logoFile }
            }
        }
    })

    # Sync bmRows when a cell is committed
    $bulkGrid.Add_CellEditEnding({
        param($s, $e)
        if ($e.EditAction.ToString() -ne 'Commit') { return }

        $header  = $e.Column.Header -as [string]
        $rowView = $e.Row.Item
        if (-not $rowView) { return }

        $id  = $rowView['_Id'] -as [string]
        $idx = Find-RowById -Id $id
        if ($idx -lt 0) { return }

        switch ($header) {
            'Source Folder' {
                # Source Folder is a DataGridTemplateColumn — EditingElement is a ContentPresenter
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if (-not $tb) {
                    $tb = Find-VisualChild -Parent $e.EditingElement `
                                          -TargetType ([System.Windows.Controls.TextBox])
                }
                if ($tb) {
                    $rawPath = $tb.Text.Trim()
                    $newPath = $rawPath.Trim('"')   # strip surrounding quotes (e.g. pasted from Explorer)
                    if ($newPath -ne $rawPath) {
                        # Correct the TextBox text and DataRow so the binding shows the clean path
                        $tb.Text = $newPath
                        $dr = Find-DataRow -Id $id
                        if ($dr) { $dr['SourceFolder'] = $newPath }
                    }
                    if ($newPath -ne ($script:bmRows[$idx].SourceFolder ?? '')) {
                        Invoke-SourceScan -Id $id -Path $newPath
                    }
                }
            }
            'Display Name' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) { $script:bmRows[$idx].DisplayName = $tb.Text.Trim() }
            }
            'Version' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) { $script:bmRows[$idx].Version = $tb.Text.Trim() }
            }
            'Publisher' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) { $script:bmRows[$idx].Publisher = $tb.Text.Trim() }
            }
            'Description' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) { $script:bmRows[$idx].Description = $tb.Text.Trim() }
            }
            'Category' {
                $cmb = Find-VisualChild -Parent $e.EditingElement `
                                        -TargetType ([System.Windows.Controls.ComboBox])
                if ($cmb) {
                    $newCat = $cmb.SelectedItem -as [string]
                    $script:bmRows[$idx].Category = $newCat
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['Category'] = $newCat ?? '' }
                }
            }
            'Setup File' {
                $cmb = Find-VisualChild -Parent $e.EditingElement `
                                        -TargetType ([System.Windows.Controls.ComboBox])
                # Prefer stashed value from DropDownClosed (dropdown pick); fall back to typed text.
                $newSetup = if ($script:pendingSetupFileId -eq $id -and $script:pendingSetupFile) {
                    $pf = $script:pendingSetupFile
                    $script:pendingSetupFile   = $null
                    $script:pendingSetupFileId = $null
                    $pf
                } elseif ($cmb) { $cmb.Text.Trim() } else { '' }
                if ($newSetup -and $newSetup -ne ($script:bmRows[$idx].SetupFile ?? '')) {
                    $script:bmRows[$idx].SetupFile = $newSetup
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['SetupFile'] = $newSetup }
                    if (-not $script:bmRows[$idx].IsPSADT) {
                        Set-RowCommandSuggestion -Id $id -FileName $newSetup
                        # Write the new commands into the DataRowView shadow copy.
                        # While the row is in edit mode WPF displays the shadow copy, not the
                        # DataRow, so this makes the change visible immediately.  CommitEdit(Row)
                        # (queued by DropDownClosed) then commits the shadow copy to the DataRow.
                        $drv = $rowView -as [System.Data.DataRowView]
                        if ($drv) {
                            $drv['InstallCmd']   = $script:bmRows[$idx].InstallCommandLine
                            $drv['UninstallCmd'] = $script:bmRows[$idx].UninstallCommandLine
                        }
                    }
                }
            }
            'Install Cmd' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) {
                    $script:bmRows[$idx].InstallCommandLine = $tb.Text.Trim()
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['InstallCmd'] = $tb.Text.Trim() }
                }
            }
            'Uninstall Cmd' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) {
                    $script:bmRows[$idx].UninstallCommandLine = $tb.Text.Trim()
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['UninstallCmd'] = $tb.Text.Trim() }
                }
            }
            'Info URL' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) {
                    $script:bmRows[$idx].InformationURL = $tb.Text.Trim()
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['InformationURL'] = $tb.Text.Trim() }
                }
            }
            'Privacy URL' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) {
                    $script:bmRows[$idx].PrivacyURL = $tb.Text.Trim()
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['PrivacyURL'] = $tb.Text.Trim() }
                }
            }
            'Logo Path' {
                $tb = $e.EditingElement -as [System.Windows.Controls.TextBox]
                if ($tb) {
                    # Strip surrounding double-quotes (e.g. pasted from Explorer with quotes)
                    $logo = $tb.Text.Trim().Trim('"')
                    if ($logo) {
                        $ext = [System.IO.Path]::GetExtension($logo).ToLower()
                        if ($ext -notin @('.png', '.jpg', '.jpeg')) {
                            [System.Windows.MessageBox]::Show(
                                "Unsupported logo format: $ext`n`nOnly PNG, JPG, and JPEG files are accepted by Intune.",
                                'Invalid Logo', 'OK', 'Warning')
                            $logo    = ''
                            $tb.Text = ''
                        }
                    }
                    $script:bmRows[$idx].LogoPath = $logo
                    $dr = Find-DataRow -Id $id
                    if ($dr) { $dr['LogoPath'] = $logo }
                }
            }
            'Template' {
                # EditingElement is a ContentPresenter for TemplateColumn — walk to ComboBox
                $cmb = Find-VisualChild -Parent $e.EditingElement `
                                        -TargetType ([System.Windows.Controls.ComboBox])
                if ($cmb -and $cmb.SelectedItem) {
                    $newTpl = $cmb.SelectedItem -as [string]
                    $script:bmRows[$idx].Template = $newTpl
                    # -Force so the template's assignment/arch overwrite the current values
                    Apply-TemplateToRow -Id $id -TemplateName $newTpl -Force
                }
            }
        }
    })

    #endregion

    # ─────────────────────────────────────────────────────────────────────────
    #region Button handlers
    # ─────────────────────────────────────────────────────────────────────────

    # ── + Add Row ─────────────────────────────────────────────────────────────
    $btnAddRow.Add_Click({
        # Template defaults to the CmbDefaultTemplate selection (handled inside Add-BmRow)
        Add-BmRow -Config @{} | Out-Null

        # Force layout pass so Items reflects the new row, then select it
        $bulkGrid.UpdateLayout()
        $newIdx = $script:bmTable.Rows.Count - 1
        if ($newIdx -ge 0 -and $newIdx -lt $bulkGrid.Items.Count) {
            $bulkGrid.SelectedIndex = $newIdx
            $bulkGrid.ScrollIntoView($bulkGrid.Items[$newIdx])
            # Put the Source Folder cell into edit mode (user can also press F2)
            try {
                $bulkGrid.CurrentCell = New-Object System.Windows.Controls.DataGridCellInfo(
                    $bulkGrid.Items[$newIdx], $bulkGrid.Columns[0])
                $bulkGrid.BeginEdit()
            } catch { }
        }
    })

    # ── Browse Source (folder picker for selected row) ────────────────────────
    $btnBrowse.Add_Click({
        $sel = $bulkGrid.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show(
                'Select a row first, then click Browse Source.',
                'Browse', 'OK', 'Information')
            return
        }

        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select the application source folder'
        $currentPath = $sel['SourceFolder'] -as [string]
        if ($currentPath -and (Test-Path $currentPath)) { $dlg.SelectedPath = $currentPath }

        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $newPath = $dlg.SelectedPath
        $id      = $sel['_Id'] -as [string]

        Invoke-SourceScan -Id $id -Path $newPath
    })

    # ── Browse Logo (image file picker for selected row) ─────────────────────
    $btnBrowseLogo.Add_Click({
        $sel = $bulkGrid.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show(
                'Select a row first, then click Browse Logo.',
                'Browse Logo', 'OK', 'Information')
            return
        }

        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title  = 'Select logo image'
        $dlg.Filter = 'Image files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*'
        $currentLogo = $sel['LogoPath'] -as [string]
        if ($currentLogo -and (Test-Path $currentLogo)) {
            $dlg.InitialDirectory = [System.IO.Path]::GetDirectoryName($currentLogo)
        }
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $newLogo = $dlg.FileName
        $ext     = [System.IO.Path]::GetExtension($newLogo).ToLower()
        if ($ext -notin @('.png', '.jpg', '.jpeg')) {
            [System.Windows.MessageBox]::Show(
                "Unsupported logo format: $ext`n`nOnly PNG, JPG, and JPEG files are accepted by Intune.",
                'Invalid Logo', 'OK', 'Warning')
            return
        }
        $rowId   = $sel['_Id'] -as [string]
        $rowIdx  = Find-RowById -Id $rowId
        if ($rowIdx -ge 0) {
            $script:bmRows[$rowIdx].LogoPath = $newLogo
            $dr = Find-DataRow -Id $rowId
            if ($dr) { $dr['LogoPath'] = $newLogo }
        }
    })

    # ── Set Assignment dropdown ───────────────────────────────────────────────
    # Applies a predefined assignment to all selected rows (or all rows if none selected).
    # Uses MenuItem.Tag to avoid PowerShell closure/variable-capture issues.
    function Apply-AssignmentPreset {
        param([hashtable]$Asg)

        $targets = @($bulkGrid.SelectedItems)
        if (-not $targets) { $targets = @($bulkGrid.Items) }

        foreach ($rowView in $targets) {
            if ($rowView -isnot [System.Data.DataRowView]) { continue }
            $id  = $rowView['_Id'] -as [string]
            $idx = Find-RowById -Id $id
            if ($idx -lt 0) { continue }

            $script:bmRows[$idx].Assignment = $Asg.Clone()

            $dr = Find-DataRow -Id $id
            if ($dr) { $dr['Assignment'] = Get-AssignmentSummary -Asg $Asg }
        }

        Refresh-StatusBar
    }

    # Build ContextMenu programmatically so MenuItem.Tag can carry the payload,
    # avoiding closure capture of loop variables.
    $asgMenu = New-Object System.Windows.Controls.ContextMenu

    $asgPresets = [ordered]@{
        'All Devices — Required'  = @{ Type='AllDevices'; Intent='required';  Notification='showAll' }
        'All Devices — Available' = @{ Type='AllDevices'; Intent='available'; Notification='showAll' }
        'All Users — Required'    = @{ Type='AllUsers';   Intent='required';  Notification='showAll' }
        'All Users — Available'   = @{ Type='AllUsers';   Intent='available'; Notification='showAll' }
        '---'                     = $null
        'None'                    = @{ Type='None' }
    }

    foreach ($label in $asgPresets.Keys) {
        if ($label -eq '---') {
            $asgMenu.Items.Add([System.Windows.Controls.Separator]::new()) | Out-Null
            continue
        }
        $mi        = [System.Windows.Controls.MenuItem]::new()
        $mi.Header = $label
        $mi.Tag    = $asgPresets[$label]   # payload stored on the item — no closure needed
        $mi.Add_Click({
            param($sender, $e)
            Apply-AssignmentPreset -Asg $sender.Tag
        })
        $asgMenu.Items.Add($mi) | Out-Null
    }

    $btnAssignment.ContextMenu = $asgMenu
    $btnAssignment.Add_Click({
        $btnAssignment.ContextMenu.PlacementTarget = $btnAssignment
        $btnAssignment.ContextMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
        $btnAssignment.ContextMenu.IsOpen    = $true
    })

    # ── Full Setup (opens Show-AppUploadForm pre-populated) ───────────────────
    function Open-FullSetup {
        param([string]$Id)
        $idx = Find-RowById -Id $Id
        if ($idx -lt 0) { return }

        $existing = $script:bmRows[$idx]

        $updated = Show-AppUploadForm `
            -TemplateFolder      $TemplateFolder `
            -DefaultOutput       $Config.DefaultOutputPath `
            -DefaultTemplate     ($Config.DefaultTemplate ?? 'PSADT-Default') `
            -Config              $Config `
            -AvailableCategories $AvailableCategories `
            -AvailableFilters    $AvailableFilters `
            -PrePopulate         $existing `
            -SubmitLabel         'Save to Queue'

        if ($updated) {
            $updated._id     = $existing._id
            $updated._status = $existing._status
            $script:bmRows[$idx] = $updated

            $dr = Find-DataRow -Id $Id
            if ($dr) {
                $dr['SourceFolder'] = $updated.SourceFolder ?? ''
                $dr['DisplayName']  = $updated.DisplayName  ?? ''
                $dr['Version']      = $updated.Version      ?? ''
                $dr['Publisher']    = $updated.Publisher    ?? ''
                $dr['Description']    = $updated.Description         ?? ''
                $dr['SetupFile']     = $updated.SetupFile            ?? ''
                $dr['InstallCmd']    = $updated.InstallCommandLine   ?? ''
                $dr['UninstallCmd']  = $updated.UninstallCommandLine ?? ''
                $dr['InformationURL'] = $updated.InformationURL      ?? ''
                $dr['PrivacyURL']    = $updated.PrivacyURL           ?? ''
                $dr['LogoPath']      = $updated.LogoPath             ?? ''
                $dr['Template']     = $updated.Template     ?? ''
                $dr['Detection']    = Get-DetectionSummary  -Det $updated.Detection
                $dr['Assignment']   = Get-AssignmentSummary -Asg $updated.Assignment

                # Category: store first selected (grid shows single; full list in bmRows)
                $catSummary = if ($updated.Categories -and $updated.Categories.Count -gt 0) {
                    $updated.Categories -join ', '
                } else { '' }
                $dr['Category'] = $catSummary
            }

            Refresh-StatusBar
        }
    }

    $btnFullSetup.Add_Click({
        $sel = $bulkGrid.SelectedItem
        if (-not $sel) {
            [System.Windows.MessageBox]::Show(
                'Select a row to configure.',
                'Full Setup', 'OK', 'Information')
            return
        }
        Open-FullSetup -Id ($sel['_Id'] -as [string])
    })

    # Double-click opens Full Setup — but only when the row was already selected before
    # the click (i.e. not on the first click that just selects the row).
    $script:preClickRowId = $null

    $bulkGrid.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        # Walk the hit-test result up to the containing DataGridRow to find which row was clicked
        $dep = $e.OriginalSource -as [System.Windows.DependencyObject]
        while ($dep -and $dep -isnot [System.Windows.Controls.DataGridRow]) {
            $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
        }
        if ($dep -is [System.Windows.Controls.DataGridRow]) {
            $item = $dep.Item -as [System.Data.DataRowView]
            $script:preClickRowId = if ($item) { $item['_Id'] -as [string] } else { $null }
        } else {
            $script:preClickRowId = $null
        }
    })

    $bulkGrid.Add_MouseDoubleClick({
        $sel = $bulkGrid.SelectedItem
        if ($sel) {
            $currentId = $sel['_Id'] -as [string]
            # Only open Full Setup when the row was already selected before this click
            if ($currentId -and $currentId -eq $script:preClickRowId) {
                $bulkGrid.CommitEdit()
                Open-FullSetup -Id $currentId
            }
        }
    })

    # ── Remove Selected ───────────────────────────────────────────────────────
    $btnRemove.Add_Click({
        $selIds = @($bulkGrid.SelectedItems | ForEach-Object { $_['_Id'] -as [string] })
        if (-not $selIds) {
            [System.Windows.MessageBox]::Show(
                'Select one or more rows to remove.', 'Remove', 'OK', 'Information')
            return
        }

        $script:bmRows.RemoveAll([Predicate[hashtable]]{ param($r) $selIds -contains $r._id }) | Out-Null

        $toRemove = @($script:bmTable.Rows | Where-Object { $selIds -contains $_['_Id'] })
        foreach ($dr in $toRemove) { $script:bmTable.Rows.Remove($dr) }

        Refresh-StatusBar
    })

    # ── Clear All ─────────────────────────────────────────────────────────────
    $btnClear.Add_Click({
        if ($script:bmRows.Count -eq 0) { return }
        $confirm = [System.Windows.MessageBox]::Show(
            "Remove all $($script:bmRows.Count) app(s) from the queue?",
            'Clear Queue', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }

        $script:bmRows.Clear()
        $script:bmTable.Rows.Clear()
        Refresh-StatusBar
    })

    # ── Import JSON ───────────────────────────────────────────────────────────
    $btnImport.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title  = 'Import bulk upload JSON'
        $dlg.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $apps = Get-Content $dlg.FileName -Raw | ConvertFrom-Json
            if ($apps -isnot [array]) { $apps = @($apps) }

            $added = 0
            foreach ($appJson in $apps) {
                $cfg = ConvertFrom-AppJson -AppJson $appJson
                if ($cfg) {
                    Add-BmRow -Config $cfg | Out-Null
                    $added++
                }
            }

            [System.Windows.MessageBox]::Show(
                "Imported $added of $($apps.Count) app(s) from:`n$($dlg.FileName)",
                'Import Complete', 'OK', 'Information')
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Import failed:`n$_", 'Import Error', 'OK', 'Error')
        }
    })

    # ── Export JSON ───────────────────────────────────────────────────────────
    $btnExport.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title    = 'Export queue to JSON'
        $dlg.Filter   = 'JSON files (*.json)|*.json'
        $dlg.FileName = 'BulkUpload.json'
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $export = $script:bmRows | ForEach-Object {
            $h = @{}
            foreach ($key in $_.Keys) {
                if ($key -notmatch '^_') { $h[$key] = $_[$key] }
            }
            $h
        }

        $export | ConvertTo-Json -Depth 10 | Set-Content $dlg.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show(
            "Exported $($script:bmRows.Count) app(s) to:`n$($dlg.FileName)",
            'Export Complete', 'OK', 'Information')
    })

    # ── Upload ────────────────────────────────────────────────────────────────
    function Start-BulkUpload {
        param([bool]$SelectedOnly = $false)

        if ($SelectedOnly) {
            $selIds = @($bulkGrid.SelectedItems | ForEach-Object { $_['_Id'] -as [string] } | Where-Object { $_ })
            if ($selIds.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    'No rows selected. Click a row to select it, then try again.',
                    'Upload Selected', 'OK', 'Warning')
                return
            }
            $toProcess = @($script:bmRows | Where-Object { $selIds -contains $_._id })
        } else {
            $toProcess = @($script:bmRows)
        }

        if ($toProcess.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'No apps to upload.', 'Upload', 'OK', 'Warning')
            return
        }

        $plural  = if ($toProcess.Count -ne 1) { 's' } else { '' }

        # Build assignment summary for the confirmation prompt
        $asgLines = $toProcess | ForEach-Object {
            $asgText = Get-AssignmentSummary -Asg $_.Assignment
            "  • $($_.DisplayName ?? $_.SourceFolder ?? '(unnamed)'):  $asgText"
        }
        $confirmMsg = "Upload $($toProcess.Count) app$plural to Intune?`n`n" +
                      "Assignment summary:`n" + ($asgLines -join "`n") +
                      "`n`nPlease confirm the assignments above are correct before proceeding.`n" +
                      "This may take several minutes."

        $confirm = [System.Windows.MessageBox]::Show(
            $confirmMsg, 'Confirm Upload & Assignments', 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }

        # Disable toolbar during upload
        foreach ($btn in @($btnAddRow,$btnBrowse,$btnBrowseLogo,$btnAssignment,$btnFullSetup,$btnRemove,$btnClear,
                           $btnImport,$btnExport,$btnUploadSel,$btnUploadAll)) {
            $btn.IsEnabled = $false
        }
        $txtUploadResult.Text = ''

        $ok = 0; $fail = 0; $i = 0

        if ($Logger) { & $Logger "─── Bulk upload started: $($toProcess.Count) app(s) ───" 'Info' }

        foreach ($row in $toProcess) {
            $i++
            $appLabel = $row.DisplayName ?? $row.SourceFolder ?? "(row $i)"
            Update-RowStatus -Id $row._id -Status 'Uploading...'
            $txtStatus.Text = "Uploading $i of $($toProcess.Count): $appLabel…"
            $window.Dispatcher.Invoke([action]{}, 'Background')

            if ($Logger) { & $Logger "[$i/$($toProcess.Count)] $appLabel" 'Info' }

            $appConfig = @{}
            foreach ($key in $row.Keys) {
                if ($key -notmatch '^_') { $appConfig[$key] = $row[$key] }
            }

            # Publisher is required by Intune
            if (-not $appConfig.Publisher) {
                $errMsg = 'Publisher is required'
                Update-RowStatus -Id $row._id -Status "FAILED — $errMsg"
                if ($Logger) { & $Logger "  $appLabel — FAILED: $errMsg" 'Fail' }
                $fail++
                continue
            }

            # Normalise Category (single string from grid) into Categories array
            # unless Categories was already populated (e.g. via Full Setup multi-select)
            if (-not $appConfig.Categories -or $appConfig.Categories.Count -eq 0) {
                if ($appConfig.Category -and $appConfig.Category -ne '') {
                    $appConfig.Categories = @($appConfig.Category)
                } else {
                    $appConfig.Categories = @()
                }
            }

            try {
                $result = Invoke-ProcessApp `
                    -AppConfig      $appConfig `
                    -Config         $Config `
                    -TemplateFolder $TemplateFolder

                if ($result.Success) {
                    Update-RowStatus -Id $row._id -Status 'Done'; $ok++
                    $appId = if ($result.App) { $result.App.id } else { '?' }
                    if ($Logger) { & $Logger "  $appLabel — uploaded (ID: $appId)" 'OK' }
                } else {
                    Update-RowStatus -Id $row._id -Status 'Failed'; $fail++
                    if ($Logger) { & $Logger "  $appLabel — FAILED: $($result.Error)" 'Fail' }
                }
            }
            catch {
                Update-RowStatus -Id $row._id -Status 'Failed'; $fail++
                if ($Logger) { & $Logger "  $appLabel — ERROR: $_" 'Fail' }
            }
        }

        if ($Logger) {
            $lvl = if ($fail -gt 0) { 'Warn' } else { 'OK' }
            & $Logger "─── Bulk upload complete: $ok succeeded, $fail failed ───" $lvl
        }

        Refresh-StatusBar
        $summary = "$ok succeeded"
        if ($fail) { $summary += ", $fail failed" }
        $txtUploadResult.Text = $summary

        # Re-enable toolbar
        foreach ($btn in @($btnAddRow,$btnBrowse,$btnBrowseLogo,$btnAssignment,$btnFullSetup,$btnRemove,$btnClear,$btnImport)) {
            $btn.IsEnabled = $true
        }
        Refresh-StatusBar  # re-enables export/upload buttons based on row count

        [System.Windows.MessageBox]::Show(
            "Bulk upload complete:`n  Succeeded: $ok`n  Failed:    $fail",
            'Upload Complete', 'OK',
            $(if ($fail -gt 0) { 'Warning' } else { 'Information' }))
    }

    $btnUploadAll.Add_Click({ Start-BulkUpload -SelectedOnly $false })
    $btnUploadSel.Add_Click({ Start-BulkUpload -SelectedOnly $true  })

    #endregion

    # ── Commit on focus-away ──────────────────────────────────────────────────
    # When any toolbar button receives focus (user clicked it), commit any in-progress
    # cell edit so text-column and ComboBox changes don't need an explicit Enter press.
    $commitEdit = { if ($bulkGrid.IsEditing) { $bulkGrid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } }
    foreach ($btn in @($btnAddRow,$btnBrowse,$btnBrowseLogo,$btnAssignment,$btnFullSetup,
                        $btnRemove,$btnClear,$btnImport,$btnExport,$btnUploadSel,$btnUploadAll)) {
        $btn.Add_GotFocus($commitEdit)
    }

    Refresh-StatusBar
    $window.ShowDialog() | Out-Null
    $global:IntuneUploaderGrid = $null
}
