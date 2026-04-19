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
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Image x:Name="ImgLogoBulk" Width="38" Height="38" Margin="0,0,12,0" VerticalAlignment="Center"
                     RenderOptions.BitmapScalingMode="HighQuality"/>
              <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Bulk Upload Manager" FontSize="20" FontWeight="Light" Foreground="White"/>
                <TextBlock Text="Add rows, edit inline ÔÇö Source Folder auto-scans PSADT metadata. Full Setup for detection/assignment."
                           FontSize="11" Foreground="#D4C5F9" Margin="0,1,0,0"/>
              </StackPanel>
            </StackPanel>
          </Grid>
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

    $imgLogoBulk = Find 'ImgLogoBulk'
    try {
        $toolLogoBase64 = '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'
        $logoBytes  = [Convert]::FromBase64String($toolLogoBase64)
        $logoStream = [System.IO.MemoryStream]::new($logoBytes)
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.StreamSource = $logoStream
        $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $logoStream.Dispose()
        $imgLogoBulk.Source = $bmp
    } catch {}

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
            $detScript = Get-ChildItem -Path $Path -Filter '*.ps1' -ErrorAction SilentlyContinue |
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
