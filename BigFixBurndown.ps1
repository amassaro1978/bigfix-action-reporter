#Requires -Version 5.1
<#
.SYNOPSIS
    BigFix M365 Burndown Chart - Visualize legacy Office version decline over time
.DESCRIPTION
    Enter multiple Action IDs from M365 deployment actions. The tool combines all successful
    installs, deduplicates by computer, and plots a burndown chart showing legacy Office
    versions declining as M365 rolls out.
.NOTES
    Author: Anthony Massaro / thomasShartalter
    Date: 2026-03-06
    Based on BigFix Action Reporter v6
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName WindowsFormsIntegration

# --- Configuration -----------------------------------------------------------
$script:Config = @{
    ServerUrl = ""; Username = ""; Password = ""; ApiBase = "/api"
}

# --- CMTrace Logging ---------------------------------------------------------
$script:LogFile = "C:\temp\BigFixBurndown.log"
$script:LogComponent = "BigFixBurndown"

function Write-CMLog {
    param(
        [string]$Message,
        [ValidateSet("Info","Warning","Error")]
        [string]$Severity = "Info"
    )
    $sevInt = switch ($Severity) { "Info" { 1 } "Warning" { 2 } "Error" { 3 } }
    $time = Get-Date -Format "HH:mm:ss.fff"
    $date = Get-Date -Format "MM-dd-yyyy"
    $tzOffset = [System.TimeZone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
    $timeStr = "$time+$($tzOffset.ToString('000'))"
    $logLine = "<![LOG[$Message]LOG]!><time=`"$timeStr`" date=`"$date`" component=`"$script:LogComponent`" context=`"`" type=`"$sevInt`" thread=`"$([System.Threading.Thread]::CurrentThread.ManagedThreadId)`" file=`"BigFixBurndown.ps1`">"
    try {
        $dir = Split-Path $script:LogFile -Parent
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        Add-Content -Path $script:LogFile -Value $logLine -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

# --- API Functions -----------------------------------------------------------
function Get-BigFixCredential {
    $secPass = $txtPass.SecurePassword
    return New-Object System.Management.Automation.PSCredential($txtUser.Text, $secPass)
}

function Invoke-BigFixAPI {
    param([string]$Endpoint)
    $uri = "$($txtServer.Text.TrimEnd('/'))$($script:Config.ApiBase)$Endpoint"
    $cred = Get-BigFixCredential
    Write-CMLog "API GET $uri"
    try {
        if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
            Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
"@
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $result = Invoke-RestMethod -Uri $uri -Credential $cred -Method Get -ContentType "application/xml"
        Write-CMLog "API GET $Endpoint - OK"
        return $result
    } catch {
        Write-CMLog "API GET $Endpoint - FAILED: $($_.Exception.Message)" -Severity Error
        throw "API Error: $($_.Exception.Message)"
    }
}

# --- Parse Status Data -------------------------------------------------------
function Parse-StatusData {
    param($StatusData)
    $endpoints = @()
    $computers = $StatusData.BESAPI.ActionResults.Computer
    if (-not $computers) { $computers = $StatusData.SelectNodes("//Computer") }
    foreach ($c in $computers) {
        $name = $c.Name; if (-not $name) { $name = $c.GetAttribute("Name") }
        $status = $c.Status; if (-not $status) { $status = ($c.SelectSingleNode("Status")).'#text' }
        $isFixed = $status -match "Fixed|executed successfully|completed|succeeded"
        $endpoints += [PSCustomObject]@{
            ComputerName = $name
            IsFixed      = $isFixed
            EndTime      = $c.EndTime
            RawStatus    = $status
        }
    }
    return $endpoints
}

# --- Build Burndown Data ----------------------------------------------------
function Build-BurndownData {
    param($AllEndpoints, [int]$TotalMachines)
    
    # Keep only successful installs with valid EndTime
    $fixed = $AllEndpoints | Where-Object { $_.IsFixed -and $_.EndTime } | Sort-Object { [datetime]$_.EndTime }
    
    # Deduplicate by computer name (keep earliest install)
    $seen = @{}
    $unique = @()
    foreach ($ep in $fixed) {
        if (-not $seen[$ep.ComputerName]) {
            $seen[$ep.ComputerName] = $true
            $unique += $ep
        }
    }
    
    # Group by date
    $byDate = $unique | Group-Object { ([datetime]$_.EndTime).ToString("yyyy-MM-dd") } | Sort-Object Name
    
    $cumulative = 0
    $burndown = @()
    foreach ($group in $byDate) {
        $cumulative += $group.Count
        $remaining = $TotalMachines - $cumulative
        $burndown += [PSCustomObject]@{
            Date             = $group.Name
            InstalledThatDay = $group.Count
            TotalInstalled   = $cumulative
            LegacyRemaining  = [Math]::Max(0, $remaining)
        }
    }
    
    return @{
        BurndownData    = $burndown
        UniqueInstalls  = $unique.Count
        TotalActions    = $AllEndpoints.Count
        TotalFixed      = $fixed.Count
        DuplicatesRemoved = $fixed.Count - $unique.Count
    }
}

# --- WPF Window --------------------------------------------------------------
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BigFix M365 Burndown Chart" Width="1100" Height="750"
    Background="#1e1e2e" WindowStartupLocation="CenterScreen">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0" Text="M365 Migration Burndown" FontSize="24" FontWeight="Light"
                   Foreground="#cdd6f4" FontFamily="Segoe UI" Margin="0,0,0,16"/>

        <!-- Connection -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
            <TextBlock Text="Server:" Foreground="#a6adc8" VerticalAlignment="Center" Margin="0,0,8,0" FontFamily="Segoe UI"/>
            <TextBox x:Name="txtServer" Width="280" Height="28" Background="#313244" Foreground="#cdd6f4"
                     BorderBrush="#45475a" FontFamily="Segoe UI" Padding="6,3" VerticalContentAlignment="Center"/>
            <TextBlock Text="User:" Foreground="#a6adc8" VerticalAlignment="Center" Margin="16,0,8,0" FontFamily="Segoe UI"/>
            <TextBox x:Name="txtUser" Width="120" Height="28" Background="#313244" Foreground="#cdd6f4"
                     BorderBrush="#45475a" FontFamily="Segoe UI" Padding="6,3" VerticalContentAlignment="Center"/>
            <TextBlock Text="Pass:" Foreground="#a6adc8" VerticalAlignment="Center" Margin="16,0,8,0" FontFamily="Segoe UI"/>
            <PasswordBox x:Name="txtPass" Width="120" Height="28" Background="#313244" Foreground="#cdd6f4"
                         BorderBrush="#45475a" FontFamily="Segoe UI" Padding="6,3" VerticalContentAlignment="Center"/>
            <Button x:Name="btnConnect" Content="Connect" Width="90" Height="28" Margin="16,0,0,0"
                    Background="#45475a" Foreground="#cdd6f4" BorderBrush="#585b70" FontFamily="Segoe UI" Cursor="Hand"/>
        </StackPanel>

        <!-- Action IDs + Total + Generate -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,12">
            <TextBlock Text="Action IDs (comma-separated):" Foreground="#a6adc8" VerticalAlignment="Center" Margin="0,0,8,0" FontFamily="Segoe UI"/>
            <TextBox x:Name="txtActionIds" Width="340" Height="28" Background="#313244" Foreground="#cdd6f4"
                     BorderBrush="#45475a" FontFamily="Segoe UI" Padding="6,3" VerticalContentAlignment="Center"
                     ToolTip="Enter Action IDs separated by commas, e.g.: 12345, 12346, 12347, 12348"/>
            <TextBlock Text="Total machines:" Foreground="#a6adc8" VerticalAlignment="Center" Margin="16,0,8,0" FontFamily="Segoe UI"/>
            <TextBox x:Name="txtTotal" Width="80" Height="28" Background="#313244" Foreground="#cdd6f4"
                     BorderBrush="#45475a" FontFamily="Segoe UI" Padding="6,3" VerticalContentAlignment="Center"
                     ToolTip="Total number of machines being migrated (for burndown calculation)"/>
            <Button x:Name="btnGenerate" Content="Generate Burndown" Width="150" Height="28" Margin="16,0,0,0"
                    Background="#89b4fa" Foreground="#1e1e2e" BorderBrush="#89b4fa" FontFamily="Segoe UI" FontWeight="SemiBold" Cursor="Hand"/>
        </StackPanel>

        <!-- Chart Area -->
        <Border Grid.Row="3" Background="#181825" CornerRadius="8" Padding="4" Margin="0,0,0,12">
            <WindowsFormsHost x:Name="chartHost"/>
        </Border>

        <!-- Stats Bar -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,12">
            <Border Background="#313244" CornerRadius="6" Padding="16,10" Margin="0,0,12,0">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Legacy Remaining:" Foreground="#a6adc8" FontFamily="Segoe UI" Margin="0,0,8,0"/>
                    <TextBlock x:Name="lblRemaining" Text="--" Foreground="#f38ba8" FontFamily="Segoe UI" FontWeight="Bold" FontSize="16"/>
                </StackPanel>
            </Border>
            <Border Background="#313244" CornerRadius="6" Padding="16,10" Margin="0,0,12,0">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="M365 Installed:" Foreground="#a6adc8" FontFamily="Segoe UI" Margin="0,0,8,0"/>
                    <TextBlock x:Name="lblInstalled" Text="--" Foreground="#a6e3a1" FontFamily="Segoe UI" FontWeight="Bold" FontSize="16"/>
                </StackPanel>
            </Border>
            <Border Background="#313244" CornerRadius="6" Padding="16,10" Margin="0,0,12,0">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Completion:" Foreground="#a6adc8" FontFamily="Segoe UI" Margin="0,0,8,0"/>
                    <TextBlock x:Name="lblPercent" Text="--" Foreground="#f9e2af" FontFamily="Segoe UI" FontWeight="Bold" FontSize="16"/>
                </StackPanel>
            </Border>
            <Border Background="#313244" CornerRadius="6" Padding="16,10" Margin="0,0,12,0">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Duplicates Removed:" Foreground="#a6adc8" FontFamily="Segoe UI" Margin="0,0,8,0"/>
                    <TextBlock x:Name="lblDupes" Text="--" Foreground="#6c7086" FontFamily="Segoe UI" FontWeight="Bold"/>
                </StackPanel>
            </Border>
        </StackPanel>

        <!-- Status + Export -->
        <DockPanel Grid.Row="5">
            <Button x:Name="btnExport" Content="Export CSV" DockPanel.Dock="Right" Width="100" Height="28"
                    Background="#45475a" Foreground="#cdd6f4" BorderBrush="#585b70" FontFamily="Segoe UI" Cursor="Hand" Margin="12,0,0,0"/>
            <TextBlock x:Name="lblStatus" Text="Enter BigFix server URL and credentials, then connect." 
                       Foreground="#6c7086" FontFamily="Segoe UI" VerticalAlignment="Center"/>
        </DockPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtServer    = $window.FindName("txtServer")
$txtUser      = $window.FindName("txtUser")
$txtPass      = $window.FindName("txtPass")
$txtActionIds = $window.FindName("txtActionIds")
$txtTotal     = $window.FindName("txtTotal")
$btnConnect   = $window.FindName("btnConnect")
$btnGenerate  = $window.FindName("btnGenerate")
$btnExport    = $window.FindName("btnExport")
$chartHost    = $window.FindName("chartHost")
$lblStatus    = $window.FindName("lblStatus")
$lblRemaining = $window.FindName("lblRemaining")
$lblInstalled = $window.FindName("lblInstalled")
$lblPercent   = $window.FindName("lblPercent")
$lblDupes     = $window.FindName("lblDupes")

# --- Create Chart ------------------------------------------------------------
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#181825")
$chart.Dock = [System.Windows.Forms.DockStyle]::Fill
$chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All

$area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "BurndownArea"
$area.BackColor = [System.Drawing.Color]::Transparent
$area.AxisX.LabelStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
$area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$area.AxisX.LabelStyle.Angle = -45
$area.AxisX.MajorGrid.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#232346")
$area.AxisX.MajorGrid.LineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Dot
$area.AxisX.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
$area.AxisX.MajorTickMark.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
$area.AxisY.LabelStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
$area.AxisY.LabelStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$area.AxisY.MajorGrid.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#232346")
$area.AxisY.MajorGrid.LineDashStyle = [System.Windows.Forms.DataVisualization.Charting.ChartDashStyle]::Dot
$area.AxisY.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
$area.AxisY.MajorTickMark.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
$area.AxisY.Title = "Legacy Office Remaining"
$area.AxisY.TitleForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
$area.AxisY.TitleFont = New-Object System.Drawing.Font("Segoe UI", 9)
$area.AxisY.Minimum = 0

# Secondary Y axis for M365 installed
$area.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::True
$area.AxisY2.LabelStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
$area.AxisY2.LabelStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$area.AxisY2.MajorGrid.Enabled = $false
$area.AxisY2.LineColor = [System.Drawing.ColorTranslator]::FromHtml("#313244")
$area.AxisY2.Title = "M365 Installed"
$area.AxisY2.TitleForeColor = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
$area.AxisY2.TitleFont = New-Object System.Drawing.Font("Segoe UI", 9)
$area.AxisY2.Minimum = 0

$chart.ChartAreas.Add($area)

# Burndown line (legacy remaining) -- red, declining
$burnSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series "LegacyRemaining"
$burnSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Spline
$burnSeries.Color = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
$burnSeries.BorderWidth = 3
$burnSeries.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
$burnSeries.MarkerSize = 6
$burnSeries.MarkerColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
$burnSeries.ToolTip = "#VALX\n#VAL legacy remaining"
$chart.Series.Add($burnSeries)

# Burndown area fill
$burnArea = New-Object System.Windows.Forms.DataVisualization.Charting.Series "BurndownFill"
$burnArea.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::SplineArea
$burnArea.Color = [System.Drawing.Color]::FromArgb(40, 243, 139, 168)
$burnArea.BorderColor = [System.Drawing.Color]::Transparent
$burnArea.BorderWidth = 0
$chart.Series.Add($burnArea)

# M365 installed line -- green, rising
$installSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series "M365Installed"
$installSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Spline
$installSeries.Color = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
$installSeries.BorderWidth = 3
$installSeries.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
$installSeries.MarkerSize = 6
$installSeries.MarkerColor = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
$installSeries.YAxisType = [System.Windows.Forms.DataVisualization.Charting.AxisType]::Secondary
$installSeries.ToolTip = "#VALX\n#VAL M365 installed"
$chart.Series.Add($installSeries)

# M365 area fill
$installArea = New-Object System.Windows.Forms.DataVisualization.Charting.Series "InstallFill"
$installArea.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::SplineArea
$installArea.Color = [System.Drawing.Color]::FromArgb(40, 166, 227, 161)
$installArea.BorderColor = [System.Drawing.Color]::Transparent
$installArea.BorderWidth = 0
$installArea.YAxisType = [System.Windows.Forms.DataVisualization.Charting.AxisType]::Secondary
$chart.Series.Add($installArea)

# Legend
$legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
$legend.BackColor = [System.Drawing.Color]::Transparent
$legend.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#a6adc8")
$legend.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Top
$legend.Alignment = [System.Drawing.StringAlignment]::Center
$chart.Legends.Add($legend)
# Hide fill series from legend
$burnArea.IsVisibleInLegend = $false
$installArea.IsVisibleInLegend = $false

$chartHost.Child = $chart

# --- Store results for export ------------------------------------------------
$script:BurndownResult = $null

# --- Connect Button ----------------------------------------------------------
$btnConnect.Add_Click({
    try {
        $lblStatus.Text = "Connecting to BigFix server..."
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        Invoke-BigFixAPI "/login" | Out-Null
        $lblStatus.Text = "Connected to $($txtServer.Text)"
        Write-CMLog "Connected successfully"
    } catch {
        $lblStatus.Text = "Connection failed: $($_.Exception.Message)"
        Write-CMLog "Connection failed: $($_.Exception.Message)" -Severity Error
    }
})

# --- Generate Burndown ------------------------------------------------------
$btnGenerate.Add_Click({
    try {
        $ids = $txtActionIds.Text -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($ids.Count -eq 0) {
            $lblStatus.Text = "Enter at least one Action ID."
            return
        }
        
        $totalMachines = 0
        if ($txtTotal.Text -match '^\d+$') {
            $totalMachines = [int]$txtTotal.Text
        } else {
            $lblStatus.Text = "Enter a valid total machine count."
            return
        }
        
        # Fetch all actions
        $allEndpoints = @()
        $actionNames = @()
        foreach ($id in $ids) {
            $lblStatus.Text = "Fetching action $id ($($ids.IndexOf($id) + 1) of $($ids.Count))..."
            $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
            
            $title = "(unknown)"
            try {
                $info = Invoke-BigFixAPI "/action/$id"
                $title = $info.BES.SingleAction.Title
            } catch {}
            $actionNames += $title
            
            $statusData = Invoke-BigFixAPI "/action/$id/status"
            $parsed = Parse-StatusData $statusData
            $allEndpoints += $parsed
            Write-CMLog "Action $id ($title): $($parsed.Count) endpoints"
        }
        
        $lblStatus.Text = "Building burndown data..."
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        
        $result = Build-BurndownData -AllEndpoints $allEndpoints -TotalMachines $totalMachines
        $script:BurndownResult = $result
        
        # Update chart
        $chart.Series["LegacyRemaining"].Points.Clear()
        $chart.Series["BurndownFill"].Points.Clear()
        $chart.Series["M365Installed"].Points.Clear()
        $chart.Series["InstallFill"].Points.Clear()
        
        # Add starting point (day before first install = total machines)
        if ($result.BurndownData.Count -gt 0) {
            $firstDate = [datetime]$result.BurndownData[0].Date
            $startLabel = $firstDate.AddDays(-1).ToString("MM/dd")
            $chart.Series["LegacyRemaining"].Points.AddXY($startLabel, $totalMachines) | Out-Null
            $chart.Series["BurndownFill"].Points.AddXY($startLabel, $totalMachines) | Out-Null
            $chart.Series["M365Installed"].Points.AddXY($startLabel, 0) | Out-Null
            $chart.Series["InstallFill"].Points.AddXY($startLabel, 0) | Out-Null
        }
        
        foreach ($row in $result.BurndownData) {
            $label = ([datetime]$row.Date).ToString("MM/dd")
            $chart.Series["LegacyRemaining"].Points.AddXY($label, $row.LegacyRemaining) | Out-Null
            $chart.Series["BurndownFill"].Points.AddXY($label, $row.LegacyRemaining) | Out-Null
            $chart.Series["M365Installed"].Points.AddXY($label, $row.TotalInstalled) | Out-Null
            $chart.Series["InstallFill"].Points.AddXY($label, $row.TotalInstalled) | Out-Null
        }
        
        $area.AxisY.Maximum = $totalMachines
        $area.AxisY2.Maximum = $totalMachines
        
        $buckets = $result.BurndownData.Count + 1
        $area.AxisX.Interval = [Math]::Max(1, [Math]::Floor($buckets / 10))
        
        # Update stats
        $lastRow = $result.BurndownData[-1]
        $lblRemaining.Text = "$($lastRow.LegacyRemaining)"
        $lblInstalled.Text = "$($result.UniqueInstalls)"
        $pct = if ($totalMachines -gt 0) { [Math]::Round(($result.UniqueInstalls / $totalMachines) * 100, 1) } else { 0 }
        $lblPercent.Text = "${pct}%"
        $lblDupes.Text = "$($result.DuplicatesRemoved)"
        
        $actionList = ($actionNames | ForEach-Object { $_.Substring(0, [Math]::Min(40, $_.Length)) }) -join " | "
        $lblStatus.Text = "Burndown generated: $($result.UniqueInstalls)/$totalMachines migrated (${pct}%) - Actions: $actionList"
        Write-CMLog "Burndown complete: $($result.UniqueInstalls)/$totalMachines (${pct}%)"
        
    } catch {
        $lblStatus.Text = "Error: $($_.Exception.Message)"
        Write-CMLog "Generate error: $($_.Exception.Message)" -Severity Error
    }
})

# --- Export CSV --------------------------------------------------------------
$btnExport.Add_Click({
    if (-not $script:BurndownResult) {
        $lblStatus.Text = "Generate a burndown first."
        return
    }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "CSV Files|*.csv"
    $dlg.FileName = "M365_Burndown_$(Get-Date -Format 'yyyyMMdd').csv"
    if ($dlg.ShowDialog()) {
        $script:BurndownResult.BurndownData | Export-Csv -Path $dlg.FileName -NoTypeInformation
        $lblStatus.Text = "Exported to $($dlg.FileName)"
        Write-CMLog "Exported CSV to $($dlg.FileName)"
    }
})

# --- Show Window -------------------------------------------------------------
$window.ShowDialog() | Out-Null
