<#
.SYNOPSIS
    GUI app (WPF) for discovering what depends on a folder before you move/delete it.

.DESCRIPTION
    Wraps the same discovery checks as the console version (scheduled tasks, services,
    registry, IIS, shortcuts, script references, open handles, env vars) in a WPF window
    with a path picker, toggleable checks, a live progress log, a sortable/filterable
    results grid color-coded by risk, and CSV/HTML export.

    The scan runs in a background runspace so the UI stays responsive.

.NOTES
    Author:  Tejjo Gadiraju
    Version: 1.0
    Date:    November 2026

    - Run as Administrator for full visibility (services, all scheduled tasks, IIS, handles).
    - Read-only: never moves, deletes, or modifies anything.
    - Requires: Windows PowerShell 5.1 or PowerShell 7+ with .NET Desktop Runtime (WPF).
    - To run: right-click > Run with PowerShell, or:  powershell.exe -ExecutionPolicy Bypass -File .\FolderDependencyScanner.ps1

.EXAMPLE
    .\FolderDependencyScanner.ps1
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Xaml

# ---------------------------------------------------------------------------
# XAML - window layout
# ---------------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Folder Dependency Scanner" Height="860" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#F4F6F9" FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="AccentBrush" Color="#2D6CDF"/>
        <SolidColorBrush x:Key="AccentDarkBrush" Color="#1F4FB0"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#D64545"/>
        <SolidColorBrush x:Key="DangerDarkBrush" Color="#B23333"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="BorderBrush2" Color="#DCE1E8"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#1F2430"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#6B7280"/>

        <!-- Card-style GroupBox: colored header strip + white rounded body -->
        <Style TargetType="GroupBox">
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Margin="0,0,0,10">
                            <StackPanel>
                                <Border Background="{StaticResource AccentBrush}" CornerRadius="8,8,0,0" Padding="12,7">
                                    <ContentPresenter ContentSource="Header" TextElement.Foreground="White"
                                                      TextElement.FontWeight="SemiBold" TextElement.FontSize="13"/>
                                </Border>
                                <ContentPresenter Margin="14,12,14,14"/>
                            </StackPanel>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Plain card (no header) for the run bar and export bar -->
        <Style x:Key="PlainCard" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="12,10"/>
        </Style>

        <!-- Base button: flat, rounded, hover/pressed/disabled states -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#EDEFF3"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#E2E5EB"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#D3D7DE"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#F2F2F2"/>
                                <Setter Property="Foreground" Value="#AEB2BB"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary (accent blue) button, for Run Scan -->
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentDarkBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentDarkBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#B7C6E8"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Danger (red) button, for Cancel -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource DangerBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderBrush" Value="{StaticResource DangerDarkBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource DangerDarkBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#F2F2F2"/>
                                <Setter Property="Foreground" Value="#AEB2BB"/>
                                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Rounded text boxes -->
        <Style TargetType="TextBox">
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,4,18,4"/>
        </Style>

        <!-- Rounded, accent-colored progress bar -->
        <Style TargetType="ProgressBar">
            <Setter Property="Height" Value="20"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Grid>
                            <Border x:Name="PART_Track" CornerRadius="10" Background="#E4E7EC"/>
                            <Border x:Name="PART_Indicator" CornerRadius="10" Background="{StaticResource AccentBrush}"
                                    HorizontalAlignment="Left"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#EEF0F3"/>
            <Setter Property="AlternatingRowBackground" Value="#FAFBFC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="RowHeight" Value="26"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="BorderBrush" Value="#4A80E4"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>

        <Style x:Key="RiskRowStyle" TargetType="DataGridRow">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Risk}" Value="High">
                    <Setter Property="Background" Value="#FFCFCF"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Risk}" Value="Medium">
                    <Setter Property="Background" Value="#FFF0B3"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Risk}" Value="Low">
                    <Setter Property="Background" Value="#DFFFDD"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="NEW">
                    <Setter Property="BorderBrush" Value="#CC0000"/>
                    <Setter Property="BorderThickness" Value="0,0,0,2"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="120"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,14">
            <TextBlock Text="Folder Dependency Scanner" FontSize="24" FontWeight="Bold" Foreground="{StaticResource TextPrimaryBrush}"/>
            <TextBlock Text="Check what actually depends on a folder before you move, delete, or archive it."
                       FontSize="12" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,3,0,0"/>
            <TextBlock Text="Created by Tejjo Gadiraju"
                       FontSize="11" FontStyle="Italic" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,2,0,0"/>
            <Border Height="3" Background="{StaticResource AccentBrush}" Margin="0,10,0,0" CornerRadius="2"/>
        </StackPanel>

        <GroupBox Grid.Row="1" Header="Target" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Folder path:" VerticalAlignment="Center" Width="90" Foreground="{StaticResource TextPrimaryBrush}"/>
                <TextBox x:Name="txtPath" Width="600" Margin="4,0" VerticalContentAlignment="Center"/>
                <Button x:Name="btnBrowse" Content="Browse..." Width="90" Margin="4,0"/>
            </StackPanel>
        </GroupBox>

        <GroupBox Grid.Row="2" Header="Options" Margin="0,0,0,8">
            <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                    <TextBlock Text="Script search roots:" VerticalAlignment="Center" Width="120" Foreground="{StaticResource TextPrimaryBrush}"/>
                    <TextBox x:Name="txtSearchRoots" Width="570" Margin="4,0" VerticalContentAlignment="Center"
                             Text="C:\Scripts"/>
                    <Button x:Name="btnAddRoot" Content="Add folder..." Width="90" Margin="4,0"/>
                </StackPanel>
                <TextBlock Text="Tip: keep this list to specific automation folders (e.g. C:\Scripts, an app's own directory) - whole drives or C:\ProgramData are usually what makes a scan slow."
                           TextWrapping="Wrap" FontStyle="Italic" Foreground="{StaticResource TextSecondaryBrush}" Margin="124,0,0,6"/>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                    <TextBlock Text="handle.exe path:" VerticalAlignment="Center" Width="120" Foreground="{StaticResource TextPrimaryBrush}"/>
                    <TextBox x:Name="txtHandleExe" Width="570" Margin="4,0" VerticalContentAlignment="Center"/>
                    <Button x:Name="btnBrowseHandle" Content="Browse..." Width="90" Margin="4,0"/>
                </StackPanel>
                <WrapPanel Orientation="Horizontal">
                    <CheckBox x:Name="chkFileTimestamps" Content="File timestamps" IsChecked="True"/>
                    <CheckBox x:Name="chkScheduledTasks" Content="Scheduled tasks" IsChecked="True"/>
                    <CheckBox x:Name="chkServices" Content="Services" IsChecked="True"/>
                    <CheckBox x:Name="chkRegistry" Content="Registry" IsChecked="True"/>
                    <CheckBox x:Name="chkIIS" Content="IIS" IsChecked="True"/>
                    <CheckBox x:Name="chkShortcuts" Content="Shortcuts" IsChecked="True"/>
                    <CheckBox x:Name="chkScriptRefs" Content="Script references" IsChecked="True"/>
                    <CheckBox x:Name="chkHandles" Content="Open handles" IsChecked="True"/>
                    <CheckBox x:Name="chkEnvVars" Content="Env variables" IsChecked="True"/>
                    <CheckBox x:Name="chkProcesses" Content="Running processes" IsChecked="True"/>
                    <CheckBox x:Name="chkJunctions" Content="Junctions/symlinks" IsChecked="True"/>
                    <CheckBox x:Name="chkOdbc" Content="ODBC DSNs" IsChecked="True"/>
                    <CheckBox x:Name="chkOracle" Content="Oracle registry/config" IsChecked="True"/>
                </WrapPanel>
            </StackPanel>
        </GroupBox>

        <Border Grid.Row="3" Style="{StaticResource PlainCard}" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal">
                <Button x:Name="btnRun" Content="Run Scan" Width="120" Height="32" Style="{StaticResource PrimaryButton}"/>
                <Button x:Name="btnCancel" Content="Cancel" Width="90" Height="32" Margin="8,0,0,0" IsEnabled="False" Style="{StaticResource DangerButton}"/>
                <ProgressBar x:Name="progressBar" Width="420" Margin="14,0" Minimum="0" Maximum="100" VerticalAlignment="Center"/>
                <TextBlock x:Name="txtStatus" Text="Idle" VerticalAlignment="Center" FontWeight="SemiBold" Margin="8,0" Foreground="{StaticResource TextPrimaryBrush}"/>
            </StackPanel>
        </Border>

        <DataGrid x:Name="dataGridResults" Grid.Row="4" AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserSortColumns="True" RowStyle="{StaticResource RiskRowStyle}" Margin="0,0,0,10">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="60"/>
                <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="140"/>
                <DataGridTextColumn Header="Risk" Binding="{Binding Risk}" Width="80"/>
                <DataGridTextColumn Header="Item" Binding="{Binding Item}" Width="300"/>
                <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>

        <GroupBox Grid.Row="5" Header="Log" Margin="0,0,0,10">
            <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     TextWrapping="Wrap" FontFamily="Consolas" FontSize="11" Background="#1E1E1E" Foreground="#D4D4D4"/>
        </GroupBox>

        <Border Grid.Row="6" Style="{StaticResource PlainCard}">
            <StackPanel Orientation="Horizontal">
                <Button x:Name="btnExportCsv" Content="Export CSV" Width="110" Margin="0,0,8,0"/>
                <Button x:Name="btnExportHtml" Content="Export HTML" Width="110" Margin="0,0,8,0"/>
                <Button x:Name="btnClear" Content="Clear Results" Width="110" Margin="0,0,8,0"/>
                <Button x:Name="btnSaveSnapshot" Content="Save Snapshot" Width="110" Margin="0,0,8,0"/>
                <Button x:Name="btnLoadBaseline" Content="Compare to Baseline..." Width="160"/>
                <TextBlock x:Name="txtBaselineInfo" VerticalAlignment="Center" Margin="10,0,0,0" FontStyle="Italic" Foreground="{StaticResource TextSecondaryBrush}"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Grab named controls
$controls = @{}
foreach ($name in @('txtPath','btnBrowse','txtSearchRoots','btnAddRoot','txtHandleExe','btnBrowseHandle',
                     'chkFileTimestamps','chkScheduledTasks','chkServices','chkRegistry','chkIIS','chkShortcuts',
                     'chkScriptRefs','chkHandles','chkEnvVars','chkProcesses','chkJunctions','chkOdbc','chkOracle',
                     'btnRun','btnCancel','progressBar','txtStatus',
                     'dataGridResults','txtLog','btnExportCsv','btnExportHtml','btnClear',
                     'btnSaveSnapshot','btnLoadBaseline','txtBaselineInfo')) {
    $controls[$name] = $window.FindName($name)
}

# ObservableCollection bound to the grid (updated on the UI thread only)
$resultsCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$controls.dataGridResults.ItemsSource = $resultsCollection

# ---------------------------------------------------------------------------
# Browse buttons
# ---------------------------------------------------------------------------
$controls.btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $controls.txtPath.Text = $dlg.SelectedPath
    }
})

$controls.btnAddRoot.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($controls.txtSearchRoots.Text.Trim().Length -gt 0) {
            $controls.txtSearchRoots.Text += ";" + $dlg.SelectedPath
        } else {
            $controls.txtSearchRoots.Text = $dlg.SelectedPath
        }
    }
})

$controls.btnBrowseHandle.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "handle.exe|handle.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $controls.txtHandleExe.Text = $dlg.FileName
    }
})

$controls.btnClear.Add_Click({
    $resultsCollection.Clear()
    $controls.txtLog.Text = ""
    $controls.txtStatus.Text = "Idle"
    $controls.progressBar.Value = 0
})

$controls.btnSaveSnapshot.Add_Click({
    if ($resultsCollection.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to save yet - run a scan first.", "Nothing to save", "OK", "Information") | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Snapshot JSON (*.json)|*.json"
    $dlg.FileName = "FolderDependencySnapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $snapshot = [PSCustomObject]@{
            TargetPath = $controls.txtPath.Text
            ScannedAt  = (Get-Date).ToString("o")
            Findings   = @($resultsCollection | ForEach-Object {
                [PSCustomObject]@{ Category = $_.Category; Item = $_.Item; Detail = $_.Detail; Risk = $_.Risk }
            })
        }
        $snapshot | ConvertTo-Json -Depth 6 | Out-File -FilePath $dlg.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("Snapshot saved to $($dlg.FileName)", "Done", "OK", "Information") | Out-Null
    }
})

$controls.btnLoadBaseline.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Snapshot JSON (*.json)|*.json"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $loaded = Get-Content -LiteralPath $dlg.FileName -Raw | ConvertFrom-Json
            $script:baselineKeys = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($f in $loaded.Findings) {
                [void]$script:baselineKeys.Add("$($f.Category)|$($f.Item)|$($f.Detail)")
            }
            $controls.txtBaselineInfo.Text = "Baseline loaded: $($loaded.TargetPath) - $($loaded.ScannedAt) ($($script:baselineKeys.Count) findings). Next completed scan will flag new items."
        } catch {
            [System.Windows.MessageBox]::Show("Could not load snapshot: $($_.Exception.Message)", "Load failed", "OK", "Error") | Out-Null
        }
    }
})

# ---------------------------------------------------------------------------
# The actual scan logic - runs inside a background runspace.
# Defined as plain text so it can be injected into a separate runspace
# (which does not share functions/closures with the main runspace).
# ---------------------------------------------------------------------------
$scanScriptText = @'
param($ScanPath, $SearchRootsCsv, $HandleExePath,
      $DoFileTimestamps, $DoScheduledTasks, $DoServices, $DoRegistry, $DoIIS,
      $DoShortcuts, $DoScriptRefs, $DoHandles, $DoEnvVars,
      $DoProcesses, $DoJunctions, $DoOdbc, $DoOracle, $SyncHash)

function Write-Log { param($msg) $SyncHash.Log.Add("$(Get-Date -Format 'HH:mm:ss')  $msg") | Out-Null }
function Add-Finding {
    param($Category, $Item, $Detail, $Risk = "Info")
    $SyncHash.Results.Add([PSCustomObject]@{
        Category = $Category; Item = $Item; Detail = $Detail; Risk = $Risk; Status = ""
    }) | Out-Null
}
function Set-Progress { param($pct) $SyncHash.Progress = $pct }

# Manual recursive file walker: works identically on Windows PowerShell 5.1 and
# PowerShell 7+ (avoids System.IO.EnumerationOptions, which only exists on .NET
# Core/PS7 and silently breaks this whole step on 5.1), and tolerates individual
# folders it can't read instead of aborting the entire walk on one ACL-denied path.
function Get-FilesRobust {
    param(
        [string]$RootPath,
        [string[]]$ExcludeDirNames = @(),
        [int]$MaxFiles = [int]::MaxValue
    )
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($RootPath)
    $count = 0
    while ($stack.Count -gt 0 -and $count -lt $MaxFiles) {
        $dir = $stack.Pop()
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($dir)
        } catch { $subDirs = @() }
        foreach ($sd in $subDirs) {
            $leaf = Split-Path $sd -Leaf
            if ($ExcludeDirNames -notcontains $leaf) { $stack.Push($sd) }
        }
        try {
            $filesHere = [System.IO.Directory]::GetFiles($dir)
        } catch { $filesHere = @() }
        foreach ($f in $filesHere) {
            if ($count -ge $MaxFiles) { break }
            $count++
            Write-Output $f
        }
    }
}

$normalizedPath = $ScanPath.TrimEnd('\').ToLower()
$pathLeaf = Split-Path $ScanPath -Leaf
$searchRoots = $SearchRootsCsv -split ';' | Where-Object { $_ -and $_.Trim().Length -gt 0 }

$totalSteps = 13
$step = 0
$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
function Complete-Step {
    param($name)
    Write-Log "  -> '$name' finished at +$([math]::Round($overallStopwatch.Elapsed.TotalSeconds,1))s total"
}

try {

if (-not (Test-Path -LiteralPath $ScanPath)) {
    Write-Log "WARNING: Path '$ScanPath' does not exist. It may already have been moved. Continuing with checks that do not require it to exist."
} else {
    try {
        $acl = Get-Acl -LiteralPath $ScanPath -ErrorAction Stop
        Add-Finding "FolderOwner" $ScanPath "Owner: $($acl.Owner)" "Info"
    } catch { Write-Log "Could not read folder owner: $($_.Exception.Message)" }
}

# 1. File timestamps
$step++
if ($DoFileTimestamps -and (Test-Path -LiteralPath $ScanPath)) {
    Write-Log "Checking file-level access/write times..."
    try {
        $filePaths = Get-FilesRobust -RootPath $ScanPath -MaxFiles 50000
        $files = foreach ($fp in $filePaths) { try { [System.IO.FileInfo]::new($fp) } catch { } }
        if ($files) {
            $files | Sort-Object LastAccessTime -Descending | Select-Object -First 5 | ForEach-Object {
                $ageDays = (New-TimeSpan -Start $_.LastAccessTime -End (Get-Date)).Days
                $risk = if ($ageDays -lt 30) { "High" } elseif ($ageDays -lt 180) { "Medium" } else { "Low" }
                Add-Finding "FileTimestamp" $_.FullName "LastAccessTime: $($_.LastAccessTime) ($ageDays days ago)" $risk
            }
            $newest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Add-Finding "FileTimestamp" $newest.FullName "Most recent LastWriteTime: $($newest.LastWriteTime)" "Info"
        } else {
            Add-Finding "FileTimestamp" $ScanPath "Folder is empty or unreadable" "Info"
        }
    } catch { Write-Log "File timestamp check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped file timestamp check." }
Complete-Step "File timestamps"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 2. Scheduled tasks
$step++
if ($DoScheduledTasks) {
    Write-Log "Checking scheduled tasks..."
    try {
        Get-ScheduledTask -ErrorAction SilentlyContinue | ForEach-Object {
            $t = $_
            foreach ($a in $t.Actions) {
                $combined = "$($a.Execute) $($a.Arguments) $($a.WorkingDirectory)"
                if ($combined.ToLower().Contains($normalizedPath) -or ($pathLeaf -and $combined -match [regex]::Escape($pathLeaf))) {
                    Add-Finding "ScheduledTask" $t.TaskName "Path: $($t.TaskPath) | Action: $combined" "High"
                }
            }
        }
    } catch { Write-Log "Scheduled task check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped scheduled task check." }
Complete-Step "Scheduled tasks"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 3. Services
$step++
if ($DoServices) {
    Write-Log "Checking Windows services..."
    try {
        Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.PathName -and $_.PathName.ToLower().Contains($normalizedPath)) {
                Add-Finding "Service" $_.Name "DisplayName: $($_.DisplayName) | PathName: $($_.PathName) | State: $($_.State)" "High"
            }
        }
    } catch { Write-Log "Service check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped service check." }
Complete-Step "Services"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 4. Registry
$step++
if ($DoRegistry) {
    Write-Log "Checking registry (Run keys, App Paths, Uninstall)..."
    $regRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($root in $regRoots) {
        if (Test-Path $root) {
            try {
                $rootProps = Get-ItemProperty -LiteralPath $root -ErrorAction SilentlyContinue
                foreach ($p in $rootProps.PSObject.Properties) {
                    if ($p.Value -is [string] -and $p.Value.ToLower().Contains($normalizedPath)) {
                        Add-Finding "Registry" "$root\$($p.Name)" $p.Value "Medium"
                    }
                }
                Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                    $props = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                    $propString = ($props.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
                    if ($propString.ToLower().Contains($normalizedPath)) {
                        Add-Finding "Registry" $_.PSPath $propString "Medium"
                    }
                }
            } catch { }
        }
    }
} else { Write-Log "Skipped registry check." }
Complete-Step "Registry"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 5. IIS
$step++
if ($DoIIS) {
    Write-Log "Checking IIS configuration (if present)..."
    if (Get-Module -ListAvailable -Name WebAdministration) {
        try {
            Import-Module WebAdministration -ErrorAction Stop
            Get-ChildItem IIS:\Sites -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_.physicalPath -and $_.physicalPath.ToLower().Contains($normalizedPath)) {
                    Add-Finding "IIS-Site" $_.Name "PhysicalPath: $($_.physicalPath)" "High"
                }
            }
        } catch { Write-Log "IIS check error: $($_.Exception.Message)" }
    } else {
        Add-Finding "IIS" "N/A" "IIS WebAdministration module not present - skipped" "Info"
    }
} else { Write-Log "Skipped IIS check." }
Complete-Step "IIS"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 6. Shortcuts
$step++
if ($DoShortcuts) {
    Write-Log "Checking shortcuts (.lnk)..."
    $shortcutRoots = @(
        "$env:PUBLIC\Desktop", "$env:USERPROFILE\Desktop",
        "$env:ProgramData\Microsoft\Windows\Start Menu",
        "$env:APPDATA\Microsoft\Windows\Start Menu"
    ) | Where-Object { Test-Path $_ }
    try {
        $shell = New-Object -ComObject WScript.Shell
        foreach ($root in $shortcutRoots) {
            Get-ChildItem -LiteralPath $root -Recurse -Filter *.lnk -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $lnk = $shell.CreateShortcut($_.FullName)
                    if ($lnk.TargetPath -and $lnk.TargetPath.ToLower().Contains($normalizedPath)) {
                        Add-Finding "Shortcut" $_.FullName "Target: $($lnk.TargetPath)" "Medium"
                    }
                } catch { }
            }
        }
    } catch { Write-Log "Shortcut check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped shortcut check." }
Complete-Step "Shortcuts"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 7. Script references
$step++
if ($DoScriptRefs) {
    Write-Log "Searching script files for hardcoded references..."
    $scriptExtSet = @(".ps1",".bat",".cmd",".vbs",".xml",".config",".ini",".sql",".psm1")
    # Folders that are almost never useful to grep and can be huge/slow (Windows internals,
    # package caches, VCS metadata, build output). Skip these to avoid multi-hour scans.
    $excludeDirNames = @("WinSxS","Installer",'$Recycle.Bin',"node_modules",".git",".hg",".svn",
                          "obj","bin","Temp","Temporary Internet Files","Package Cache")
    $maxFileSizeBytes = 5MB   # script/config files bigger than this are almost always logs, not references
    $maxFilesPerRoot   = 25000 # safety cap so one enormous root can't hang the whole scan

    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) {
            Add-Finding "ScriptReference" $root "Search root does not exist - skipped" "Info"
            continue
        }
        Write-Log "  Scanning $root ..."
        $rootStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $scannedInRoot = 0
        try {
            foreach ($filePath in (Get-FilesRobust -RootPath $root -ExcludeDirNames $excludeDirNames -MaxFiles $maxFilesPerRoot)) {
                $ext = [System.IO.Path]::GetExtension($filePath)
                if ($scriptExtSet -notcontains $ext.ToLower()) { continue }
                try {
                    $fi = [System.IO.FileInfo]::new($filePath)
                    if ($fi.Length -gt $maxFileSizeBytes) { continue }
                    $content = [System.IO.File]::ReadAllText($filePath)
                    if ($content -and $content.ToLower().Contains($normalizedPath)) {
                        Add-Finding "ScriptReference" $filePath "Contains reference to path" "High"
                    }
                } catch { }
                $scannedInRoot++
            }
        } catch { Write-Log "  Error scanning $root : $($_.Exception.Message)" }
        $rootStopwatch.Stop()
        Write-Log "  $root : $scannedInRoot file(s) checked in $([math]::Round($rootStopwatch.Elapsed.TotalSeconds,1))s"
    }
} else { Write-Log "Skipped script reference search." }
Complete-Step "Script references"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 8. Open handles
$step++
if ($DoHandles) {
    Write-Log "Checking open file handles..."
    if ($HandleExePath -and (Test-Path $HandleExePath)) {
        try {
            $handleOutput = & $HandleExePath -accepteula $ScanPath 2>$null
            if ($handleOutput -match "pid:") {
                Add-Finding "OpenHandle" $ScanPath "Live open handle(s) detected:`n$handleOutput" "High"
            } else {
                Add-Finding "OpenHandle" $ScanPath "No open handles detected at scan time" "Info"
            }
        } catch { Write-Log "handle.exe check error: $($_.Exception.Message)" }
    } else {
        Add-Finding "OpenHandle" "N/A" "handle.exe not provided/found - skipped" "Info"
    }
} else { Write-Log "Skipped open handle check." }
Complete-Step "Open handles"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 9. Env vars
$step++
if ($DoEnvVars) {
    Write-Log "Checking environment variables..."
    foreach ($scope in @("Machine","User")) {
        $vars = [System.Environment]::GetEnvironmentVariables($scope)
        foreach ($key in $vars.Keys) {
            $val = [string]$vars[$key]
            if ($val.ToLower().Contains($normalizedPath)) {
                Add-Finding "EnvironmentVariable" "$scope`:$key" $val "Medium"
            }
        }
    }
} else { Write-Log "Skipped environment variable check." }
Complete-Step "Environment variables"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 10. Running processes (path + loaded modules)
$step++
if ($DoProcesses) {
    Write-Log "Checking running processes and loaded modules..."
    try {
        Get-Process -ErrorAction SilentlyContinue | ForEach-Object {
            $proc = $_
            try {
                if ($proc.Path -and $proc.Path.ToLower().Contains($normalizedPath)) {
                    Add-Finding "RunningProcess" "$($proc.ProcessName) (PID $($proc.Id))" "Path: $($proc.Path)" "High"
                }
            } catch { }
            try {
                foreach ($m in $proc.Modules) {
                    if ($m.FileName -and $m.FileName.ToLower().Contains($normalizedPath)) {
                        Add-Finding "RunningProcess" "$($proc.ProcessName) (PID $($proc.Id))" "Loaded module: $($m.FileName)" "High"
                        break
                    }
                }
            } catch { } # access denied on modules (bitness mismatch / non-admin) - expected, not fatal
        }
    } catch { Write-Log "Process check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped running process check." }
Complete-Step "Running processes"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 11. Junctions / symlinks pointing at the target
$step++
if ($DoJunctions) {
    Write-Log "Checking for junctions/symlinks (best-effort, limited to search roots + target's parent)..."
    try {
        if (Test-Path -LiteralPath $ScanPath) {
            $item = Get-Item -LiteralPath $ScanPath -Force -ErrorAction SilentlyContinue
            if ($item -and ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                Add-Finding "Junction" $ScanPath "Target itself is a reparse point (junction/symlink) - the real data may live elsewhere" "High"
            }
        }
        $junctionCheckRoots = @($searchRoots) + @((Split-Path $ScanPath -Parent))
        foreach ($jroot in ($junctionCheckRoots | Select-Object -Unique)) {
            if (-not (Test-Path $jroot)) { continue }
            try {
                $dirOutput = & cmd.exe /c "dir `"$jroot`" /AL /S 2>NUL"
                foreach ($line in $dirOutput) {
                    if ($line -match '<(JUNCTION|SYMLINKD|SYMLINK)>\s+(.+?)\s+\[(.+)\]') {
                        $linkName = $Matches[2].Trim()
                        $linkTarget = $Matches[3].Trim()
                        if ($linkTarget.ToLower().Contains($normalizedPath)) {
                            Add-Finding "Junction" $linkName "Reparse point targeting: $linkTarget" "High"
                        }
                    }
                }
            } catch { }
        }
    } catch { Write-Log "Junction check error: $($_.Exception.Message)" }
} else { Write-Log "Skipped junction/symlink check." }
Complete-Step "Junctions/symlinks"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 12. ODBC DSNs (system + user, 32-bit + 64-bit registry views)
$step++
if ($DoOdbc) {
    Write-Log "Checking ODBC Data Sources..."
    $odbcRoots = @(
        "HKLM:\SOFTWARE\ODBC\ODBC.INI",
        "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI",
        "HKCU:\SOFTWARE\ODBC\ODBC.INI"
    )
    foreach ($odbcRoot in $odbcRoots) {
        if (-not (Test-Path $odbcRoot)) { continue }
        try {
            Get-ChildItem $odbcRoot -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -ne "ODBC Data Sources" } | ForEach-Object {
                try {
                    $props = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                    $propString = ($props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join "; "
                    if ($propString.ToLower().Contains($normalizedPath)) {
                        Add-Finding "OdbcDsn" $_.PSChildName "$propString  [$odbcRoot]" "High"
                    }
                } catch { }
            }
        } catch { Write-Log "ODBC check error at $odbcRoot : $($_.Exception.Message)" }
    }
} else { Write-Log "Skipped ODBC DSN check." }
Complete-Step "ODBC DSNs"
Set-Progress ([math]::Round(($step / $totalSteps) * 100))

# 13. Oracle-specific: registry ORACLE_HOME entries + tnsnames.ora/sqlnet.ora content
$step++
if ($DoOracle) {
    Write-Log "Checking Oracle registry entries and network config files..."
    $oracleRegRoots = @("HKLM:\SOFTWARE\ORACLE", "HKLM:\SOFTWARE\WOW6432Node\ORACLE")
    $oracleHomes = New-Object System.Collections.Generic.List[string]
    foreach ($oraRoot in $oracleRegRoots) {
        if (-not (Test-Path $oraRoot)) { continue }
        try {
            Get-ChildItem $oraRoot -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $props = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
                    foreach ($p in $props.PSObject.Properties) {
                        if ($p.Value -is [string]) {
                            if ($p.Name -match 'ORACLE_HOME' -and $p.Value.Trim().Length -gt 0) {
                                $oracleHomes.Add($p.Value) | Out-Null
                            }
                            if ($p.Value.ToLower().Contains($normalizedPath)) {
                                Add-Finding "OracleRegistry" $_.PSChildName "$($p.Name)=$($p.Value)  [$oraRoot]" "High"
                            }
                        }
                    }
                } catch { }
            }
        } catch { Write-Log "Oracle registry check error at $oraRoot : $($_.Exception.Message)" }
    }
    # Also consider TNS_ADMIN / ORACLE_HOME env vars as extra places to look for config files
    if ($env:TNS_ADMIN) { $oracleHomes.Add($env:TNS_ADMIN) | Out-Null }
    if ($env:ORACLE_HOME) { $oracleHomes.Add((Join-Path $env:ORACLE_HOME "network\admin")) | Out-Null }
    foreach ($home in ($oracleHomes | Select-Object -Unique)) {
        $networkAdmin = if ($home -match 'network\\admin$') { $home } else { Join-Path $home "network\admin" }
        foreach ($cfgName in @("tnsnames.ora","sqlnet.ora")) {
            $cfgPath = Join-Path $networkAdmin $cfgName
            if (Test-Path -LiteralPath $cfgPath) {
                try {
                    $content = Get-Content -LiteralPath $cfgPath -Raw -ErrorAction SilentlyContinue
                    if ($content -and $content.ToLower().Contains($normalizedPath)) {
                        Add-Finding "OracleConfig" $cfgPath "Contains reference to path" "High"
                    }
                } catch { }
            }
        }
    }
} else { Write-Log "Skipped Oracle-specific check." }
Complete-Step "Oracle registry/config"
Set-Progress 100

Write-Log "Scan complete."

} catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)"
    Write-Log $_.ScriptStackTrace
} finally {
    $SyncHash.Completed = $true
}
'@

# ---------------------------------------------------------------------------
# Run button - kicks off the scan on a background runspace
# ---------------------------------------------------------------------------
$controls.btnRun.Add_Click({
    $path = $controls.txtPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.MessageBox]::Show("Enter or browse to a folder path first.", "Missing path", "OK", "Warning") | Out-Null
        return
    }

    $resultsCollection.Clear()
    $controls.txtLog.Text = ""
    $controls.progressBar.Value = 0
    $controls.txtStatus.Text = "Scanning..."
    $controls.txtStatus.Foreground = "Black"
    $controls.btnRun.IsEnabled = $false
    $controls.btnCancel.IsEnabled = $true

    try {
        $script:syncHash = [hashtable]::Synchronized(@{
            Results   = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
            Log       = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
            Progress  = 0
            Completed = $false
        })

        $script:runspace = [runspacefactory]::CreateRunspace()
        $script:runspace.ApartmentState = "STA"
        $script:runspace.ThreadOptions = "ReuseThread"
        $script:runspace.Open()

        $script:ps = [powershell]::Create()
        $script:ps.Runspace = $script:runspace
        [void]$script:ps.AddScript($scanScriptText).AddParameters(@{
            ScanPath         = $path
            SearchRootsCsv   = $controls.txtSearchRoots.Text
            HandleExePath    = $controls.txtHandleExe.Text
            DoFileTimestamps = [bool]$controls.chkFileTimestamps.IsChecked
            DoScheduledTasks = [bool]$controls.chkScheduledTasks.IsChecked
            DoServices       = [bool]$controls.chkServices.IsChecked
            DoRegistry       = [bool]$controls.chkRegistry.IsChecked
            DoIIS            = [bool]$controls.chkIIS.IsChecked
            DoShortcuts      = [bool]$controls.chkShortcuts.IsChecked
            DoScriptRefs     = [bool]$controls.chkScriptRefs.IsChecked
            DoHandles        = [bool]$controls.chkHandles.IsChecked
            DoEnvVars        = [bool]$controls.chkEnvVars.IsChecked
            DoProcesses      = [bool]$controls.chkProcesses.IsChecked
            DoJunctions      = [bool]$controls.chkJunctions.IsChecked
            DoOdbc           = [bool]$controls.chkOdbc.IsChecked
            DoOracle         = [bool]$controls.chkOracle.IsChecked
            SyncHash         = $script:syncHash
        })

        $script:asyncResult = $script:ps.BeginInvoke()

        # Track how much of Results/Log we've already flushed to the UI collections
        $script:lastResultCount = 0
        $script:lastLogCount = 0

        $script:timer = New-Object System.Windows.Threading.DispatcherTimer
        $script:timer.Interval = [TimeSpan]::FromMilliseconds(300)

        $script:timer.Add_Tick({
            try {
                # Flush new results
                if ($script:syncHash.Results.Count -gt $script:lastResultCount) {
                    for ($i = $script:lastResultCount; $i -lt $script:syncHash.Results.Count; $i++) {
                        $resultsCollection.Add($script:syncHash.Results[$i])
                    }
                    $script:lastResultCount = $script:syncHash.Results.Count
                }
                # Flush new log lines
                if ($script:syncHash.Log.Count -gt $script:lastLogCount) {
                    for ($i = $script:lastLogCount; $i -lt $script:syncHash.Log.Count; $i++) {
                        $controls.txtLog.AppendText($script:syncHash.Log[$i] + "`r`n")
                    }
                    $controls.txtLog.ScrollToEnd()
                    $script:lastLogCount = $script:syncHash.Log.Count
                }
                $controls.progressBar.Value = $script:syncHash.Progress

                if ($script:syncHash.Completed) {
                    $script:timer.Stop()
                    $script:ps.EndInvoke($script:asyncResult) | Out-Null
                    $script:ps.Dispose()
                    $script:runspace.Close()
                    $script:runspace.Dispose()

                    $highCount = ($resultsCollection | Where-Object Risk -eq "High").Count
                    $medCount  = ($resultsCollection | Where-Object Risk -eq "Medium").Count

                    if ($highCount -gt 0) {
                        $controls.txtStatus.Text = "$highCount HIGH-risk finding(s) - do not move without investigating"
                        $controls.txtStatus.Foreground = "Red"
                    } elseif ($medCount -gt 0) {
                        $controls.txtStatus.Text = "$medCount medium-risk finding(s) - worth a quick check"
                        $controls.txtStatus.Foreground = "#B8860B"
                    } else {
                        $controls.txtStatus.Text = "No strong dependency signals found"
                        $controls.txtStatus.Foreground = "Green"
                    }
                    $controls.btnRun.IsEnabled = $true
                    $controls.btnCancel.IsEnabled = $false

                    if ($script:baselineKeys) {
                        $newCount = 0
                        foreach ($item in $resultsCollection) {
                            $key = "$($item.Category)|$($item.Item)|$($item.Detail)"
                            if (-not $script:baselineKeys.Contains($key)) {
                                $item.Status = "NEW"
                                $newCount++
                            }
                        }
                        if ($newCount -gt 0) {
                            $controls.txtStatus.Text += "  ($newCount new since baseline)"
                        }
                    }
                }
            } catch {
                $script:timer.Stop()
                $controls.txtStatus.Text = "ERROR during scan - see log"
                $controls.txtStatus.Foreground = "Red"
                $controls.txtLog.AppendText("UI TIMER ERROR: $($_.Exception.Message)`r`n$($_.ScriptStackTrace)`r`n")
                $controls.btnRun.IsEnabled = $true
                $controls.btnCancel.IsEnabled = $false
            }
        })

        $script:timer.Start()
    } catch {
        $controls.txtStatus.Text = "ERROR starting scan - see log"
        $controls.txtStatus.Foreground = "Red"
        $controls.txtLog.AppendText("STARTUP ERROR: $($_.Exception.Message)`r`n$($_.ScriptStackTrace)`r`n")
        $controls.btnRun.IsEnabled = $true
        $controls.btnCancel.IsEnabled = $false
    }
})

$controls.btnCancel.Add_Click({
    if ($script:timer) { $script:timer.Stop() }
    try {
        if ($script:ps) {
            $script:ps.Stop()
            $script:ps.Dispose()
        }
        if ($script:runspace) {
            $script:runspace.Close()
            $script:runspace.Dispose()
        }
    } catch { }
    $controls.txtStatus.Text = "Cancelled - partial results shown below"
    $controls.txtStatus.Foreground = "#B8860B"
    $controls.btnRun.IsEnabled = $true
    $controls.btnCancel.IsEnabled = $false
})

# ---------------------------------------------------------------------------
# Export buttons
# ---------------------------------------------------------------------------
$controls.btnExportCsv.Add_Click({
    if ($resultsCollection.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to export yet - run a scan first.", "Nothing to export", "OK", "Information") | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV file (*.csv)|*.csv"
    $dlg.FileName = "FolderDependencyReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $resultsCollection | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.MessageBox]::Show("Exported to $($dlg.FileName)", "Done", "OK", "Information") | Out-Null
    }
})

$controls.btnExportHtml.Add_Click({
    if ($resultsCollection.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to export yet - run a scan first.", "Nothing to export", "OK", "Information") | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "HTML file (*.html)|*.html"
    $dlg.FileName = "FolderDependencyReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-Type -AssemblyName System.Web
        $rows = foreach ($f in ($resultsCollection | Sort-Object @{Expression = {
                switch ($_.Risk) { "High" {0} "Medium" {1} "Low" {2} default {3} }
            }})) {
            "<tr class='$($f.Risk)'><td>$($f.Status)</td><td>$($f.Category)</td><td>$($f.Risk)</td><td>$([System.Web.HttpUtility]::HtmlEncode($f.Item))</td><td>$([System.Web.HttpUtility]::HtmlEncode($f.Detail))</td></tr>"
        }
        $html = @"
<html><head><title>Folder Dependency Report</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; vertical-align: top; font-size: 13px; }
th { background: #333; color: #fff; }
tr.High { background: #ffdddd; } tr.Medium { background: #fff6d5; } tr.Low { background: #eaffea; }
</style></head><body>
<h2>Folder Dependency Report</h2>
<p><b>Target:</b> $($controls.txtPath.Text)</p>
<p><b>Generated:</b> $(Get-Date)</p>
<table><tr><th>Status</th><th>Category</th><th>Risk</th><th>Item</th><th>Detail</th></tr>
$($rows -join "`n")
</table></body></html>
"@
        $html | Out-File -FilePath $dlg.FileName -Encoding UTF8
        [System.Windows.MessageBox]::Show("Exported to $($dlg.FileName)", "Done", "OK", "Information") | Out-Null
    }
})

# ---------------------------------------------------------------------------
# Show window
# ---------------------------------------------------------------------------
$window.ShowDialog() | Out-Null
