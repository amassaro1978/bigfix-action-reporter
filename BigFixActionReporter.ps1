#Requires -Version 5.1
<#
.SYNOPSIS
    BigFix Action Reporter v5 - Pure .NET charting, no browser dependencies
.DESCRIPTION
    PowerShell/WPF GUI with .NET Charts for action deployment status visualization.
    No IE11, no WebBrowser control, no security warnings.
.NOTES
    Author: Anthony / thomasShartalter
    Date: 2026-02-12
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName WindowsFormsIntegration

# ─── Configuration ───────────────────────────────────────────────────────────
$script:Config = @{
    ServerUrl = ""; Username = ""; Password = ""; ApiBase = "/api"
}

# ─── XAML UI ─────────────────────────────────────────────────────────────────
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:wfi="clr-namespace:System.Windows.Forms.Integration;assembly=WindowsFormsIntegration"
        Title="BigFix Action Reporter" Height="920" Width="1220"
        WindowStartupLocation="CenterScreen"
        Background="#0f0f1a" Foreground="#cdd6f4">

    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#89b4fa"/>
            <Setter Property="Foreground" Value="#1e1e2e"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#b4d0fb"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#1a1a2e"/>
            <Setter Property="Foreground" Value="#cdd6f4"/>
            <Setter Property="BorderBrush" Value="#313244"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#bac2de"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#0f0f1a"/>
            <Setter Property="Foreground" Value="#cdd6f4"/>
            <Setter Property="BorderBrush" Value="#313244"/>
            <Setter Property="RowBackground" Value="#141425"/>
            <Setter Property="AlternatingRowBackground" Value="#1a1a2e"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#232336"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
            <TextBlock Text="BIGFIX ACTION REPORTER" FontSize="28" FontWeight="Bold" 
                       Foreground="#89b4fa" FontFamily="Segoe UI">
                <TextBlock.Effect>
                    <DropShadowEffect Color="#89b4fa" BlurRadius="20" ShadowDepth="0" Opacity="0.5"/>
                </TextBlock.Effect>
            </TextBlock>
            <TextBlock Text="  v5" FontSize="14" Foreground="#6c7086" VerticalAlignment="Bottom" Margin="0,0,0,4"/>
        </StackPanel>

        <!-- Connection -->
        <Border Grid.Row="1" Background="#141425" CornerRadius="10" Padding="14" Margin="0,0,0,10"
                BorderBrush="#232346" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Content="Server:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtServer" Grid.Column="1" Margin="4,0" ToolTip="https://bigfix-server:52311"/>
                <Label Content="User:" Grid.Column="2" VerticalAlignment="Center"/>
                <TextBox x:Name="txtUser" Grid.Column="3" Margin="4,0"/>
                <Label Content="Pass:" Grid.Column="4" VerticalAlignment="Center"/>
                <PasswordBox x:Name="txtPass" Grid.Column="5" Margin="4,0"
                             Background="#1a1a2e" Foreground="#cdd6f4" BorderBrush="#313244" Padding="8,6" FontSize="14"/>
                <Button x:Name="btnConnect" Content="Connect" Grid.Column="7"/>
            </Grid>
        </Border>

        <!-- Action Bar -->
        <Border Grid.Row="2" Background="#141425" CornerRadius="10" Padding="14" Margin="0,0,0,10"
                BorderBrush="#232346" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="200"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Content="Action IDs:" VerticalAlignment="Center" FontSize="15" FontWeight="SemiBold"/>
                <TextBox x:Name="txtActionId" Grid.Column="1" Margin="4,0" FontSize="14"
                         ToolTip="Comma-separated for multi-action compare"/>
                <Button x:Name="btnFetch" Content="Fetch Status" Grid.Column="2" Margin="8,0"/>
                <Button x:Name="btnRefresh" Content="Refresh" Grid.Column="3" Margin="4,0" Background="#a6e3a1" IsEnabled="False"/>
                <!-- Demo button removed for production -->
                <TextBlock x:Name="lblActionName" Grid.Column="5" VerticalAlignment="Center" 
                           Margin="12,0" FontSize="13" Foreground="#a6adc8" TextTrimming="CharacterEllipsis"/>
                <Button x:Name="btnExport" Content="Export CSV" Grid.Column="6" Background="#f9e2af" IsEnabled="False"/>
            </Grid>
        </Border>

        <!-- Main Content: TabControl for multi-action -->
        <TabControl x:Name="tabActions" Grid.Row="3" Background="#0f0f1a" BorderBrush="#232346"
                    Foreground="#cdd6f4" Padding="0" Margin="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Background" Value="#1a1a2e"/>
                    <Setter Property="Foreground" Value="#bac2de"/>
                    <Setter Property="BorderBrush" Value="#232346"/>
                    <Setter Property="Padding" Value="14,6"/>
                    <Setter Property="FontSize" Value="12"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border x:Name="bd" Background="#1a1a2e" CornerRadius="6,6,0,0" 
                                        Padding="{TemplateBinding Padding}" Margin="2,0,0,0"
                                        BorderBrush="#232346" BorderThickness="1,1,1,0">
                                    <ContentPresenter ContentSource="Header"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="bd" Property="Background" Value="#141425"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </TabControl.Resources>
        </TabControl>

        <!-- Status Bar -->
        <Border Grid.Row="4" Background="#141425" CornerRadius="8" Padding="10,6" Margin="0,8,0,0"
                BorderBrush="#232346" BorderThickness="1">
            <Grid>
                <TextBlock x:Name="lblStatus" Text="Ready -- Enter server details and connect" 
                           FontSize="12" Foreground="#a6adc8"/>
                <TextBlock x:Name="lblLastRefresh" Text="" FontSize="12" Foreground="#6c7086"
                           HorizontalAlignment="Right"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ─── Load XAML ────────────────────────────────────────────────────────────────
$reader = New-Object System.Xml.XmlNodeReader $XAML
$window = [Windows.Markup.XamlReader]::Load($reader)

$XAML.SelectNodes("//*[@*[contains(translate(name(),'x','X'),'Name')]]") | ForEach-Object {
    Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) -Scope Script
}

# Store per-tab data
$script:TabData = @{}

# ─── Chart Factory Functions ──────────────────────────────────────────────────

function New-DonutChart {
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#141425")
    $chart.Dock = [System.Windows.Forms.DockStyle]::Fill
    $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All

    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "DonutArea"
    $area.BackColor = [System.Drawing.Color]::Transparent
    $area.Position = New-Object System.Windows.Forms.DataVisualization.Charting.ElementPosition(2, 2, 96, 80)
    $chart.ChartAreas.Add($area)

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Status"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
    $series.SetCustomProperty("DoughnutRadius", "35")
    $series.SetCustomProperty("PieStartAngle", "270")
    $series.SetCustomProperty("PieLabelStyle", "Outside")
    $series.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $series["PieLineColor"] = "#45475a"
    $chart.Series.Add($series)

    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend "StatusLegend"
    $legend.BackColor = [System.Drawing.Color]::Transparent
    $legend.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#bac2de")
    $legend.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Bottom
    $legend.Alignment = [System.Drawing.StringAlignment]::Center
    $chart.Legends.Add($legend)

    return $chart
}

function New-TimelineChart {
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#141425")
    $chart.Dock = [System.Windows.Forms.DockStyle]::Fill
    $chart.AntiAliasing = [System.Windows.Forms.DataVisualization.Charting.AntiAliasingStyles]::All

    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "TimeArea"
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
    $area.AxisY.Title = "Endpoints Completed"
    $area.AxisY.TitleForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
    $area.AxisY.TitleFont = New-Object System.Drawing.Font("Segoe UI", 9)
    $area.AxisY.Minimum = 0
    $chart.ChartAreas.Add($area)

    $areaSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Area"
    $areaSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::SplineArea
    $areaSeries.Color = [System.Drawing.Color]::FromArgb(50, 166, 227, 161)
    $areaSeries.BorderColor = [System.Drawing.Color]::Transparent
    $areaSeries.BorderWidth = 0
    $chart.Series.Add($areaSeries)

    $lineSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Completions"
    $lineSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Spline
    $lineSeries.Color = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
    $lineSeries.BorderWidth = 3
    $lineSeries.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::None
    $lineSeries.ToolTip = "#VALX\n#VAL completed"
    $chart.Series.Add($lineSeries)

    $chart.Legends.Clear()
    return $chart
}

# ─── Tab Builder ──────────────────────────────────────────────────────────────

function New-ActionTab {
    param([string]$TabLabel, [string]$ActionLabel, $Data)
    
    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $TabLabel
    $tabItem.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1a1a2e")
    $tabItem.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#bac2de")
    
    $tabKey = $TabLabel
    $script:TabData[$tabKey] = @{ Data = $Data; Filter = "Fixed" }
    
    # Main grid for tab content
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0f0f1a")
    
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::new(420)
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $mainGrid.RowDefinitions.Add($row1)
    $mainGrid.RowDefinitions.Add($row2)
    
    # Charts row
    $chartsGrid = New-Object System.Windows.Controls.Grid
    $chartsGrid.Margin = [System.Windows.Thickness]::new(0,0,0,10)
    [System.Windows.Controls.Grid]::SetRow($chartsGrid, 0)
    
    $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::new(300)
    $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = [System.Windows.GridLength]::new(280)
    $col3 = New-Object System.Windows.Controls.ColumnDefinition; $col3.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $chartsGrid.ColumnDefinitions.Add($col1)
    $chartsGrid.ColumnDefinitions.Add($col2)
    $chartsGrid.ColumnDefinitions.Add($col3)
    
    $bc = [System.Windows.Media.BrushConverter]::new()
    
    # ── Donut Panel ──
    $donutBorder = New-Object System.Windows.Controls.Border
    $donutBorder.Background = $bc.ConvertFromString("#141425")
    $donutBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $donutBorder.Margin = [System.Windows.Thickness]::new(0,0,10,0)
    $donutBorder.BorderBrush = $bc.ConvertFromString("#232346")
    $donutBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    [System.Windows.Controls.Grid]::SetColumn($donutBorder, 0)
    
    $donutGrid = New-Object System.Windows.Controls.Grid
    $donutTitle = New-Object System.Windows.Controls.TextBlock
    $donutTitle.Text = "STATUS BREAKDOWN"
    $donutTitle.FontSize = 11; $donutTitle.FontWeight = "Bold"
    $donutTitle.Foreground = $bc.ConvertFromString("#89b4fa")
    $donutTitle.Margin = [System.Windows.Thickness]::new(14,10,0,0)
    $donutTitle.VerticalAlignment = "Top"; $donutTitle.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    
    $donutSub = New-Object System.Windows.Controls.TextBlock
    $donutSub.Text = $ActionLabel; $donutSub.FontSize = 9
    $donutSub.Foreground = $bc.ConvertFromString("#6c7086")
    $donutSub.Margin = [System.Windows.Thickness]::new(14,24,14,0)
    $donutSub.VerticalAlignment = "Top"; $donutSub.TextTrimming = "CharacterEllipsis"
    
    $donutHost = New-Object System.Windows.Forms.Integration.WindowsFormsHost
    $donutHost.Margin = [System.Windows.Thickness]::new(8,38,8,8)
    $donutChart = New-DonutChart
    $donutHost.Child = $donutChart
    
    $donutGrid.Children.Add($donutTitle) | Out-Null
    $donutGrid.Children.Add($donutSub) | Out-Null
    $donutGrid.Children.Add($donutHost) | Out-Null
    $donutBorder.Child = $donutGrid
    
    # ── Gauge Panel ──
    $gaugeBorder = New-Object System.Windows.Controls.Border
    $gaugeBorder.Background = $bc.ConvertFromString("#141425")
    $gaugeBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $gaugeBorder.Margin = [System.Windows.Thickness]::new(0,0,10,0)
    $gaugeBorder.BorderBrush = $bc.ConvertFromString("#232346")
    $gaugeBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    [System.Windows.Controls.Grid]::SetColumn($gaugeBorder, 1)
    
    $gaugeGrid = New-Object System.Windows.Controls.Grid
    $gaugeTitle = New-Object System.Windows.Controls.TextBlock
    $gaugeTitle.Text = "COMPLETION RATE"
    $gaugeTitle.FontSize = 11; $gaugeTitle.FontWeight = "Bold"
    $gaugeTitle.Foreground = $bc.ConvertFromString("#cba6f7")
    $gaugeTitle.Margin = [System.Windows.Thickness]::new(14,10,0,0)
    $gaugeTitle.VerticalAlignment = "Top"; $gaugeTitle.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    
    $gaugeSub = New-Object System.Windows.Controls.TextBlock
    $gaugeSub.Text = $ActionLabel; $gaugeSub.FontSize = 9
    $gaugeSub.Foreground = $bc.ConvertFromString("#6c7086")
    $gaugeSub.Margin = [System.Windows.Thickness]::new(14,24,14,0)
    $gaugeSub.VerticalAlignment = "Top"; $gaugeSub.TextAlignment = "Center"; $gaugeSub.TextTrimming = "CharacterEllipsis"
    
    # Completion %
    $relevant = $Data.Total - $Data.StatusCounts.NotRelevant
    $pct = if ($relevant -gt 0) { [math]::Round(($Data.StatusCounts.Fixed / $relevant) * 100, 1) } else { 0 }
    $gaugeColor = if ($pct -ge 80) { "#a6e3a1" } elseif ($pct -ge 50) { "#f9e2af" } else { "#f38ba8" }
    
    $gaugeStack = New-Object System.Windows.Controls.StackPanel
    $gaugeStack.VerticalAlignment = "Center"; $gaugeStack.HorizontalAlignment = "Center"
    $gaugeStack.Margin = [System.Windows.Thickness]::new(0,20,0,0)
    
    $pctText = New-Object System.Windows.Controls.TextBlock
    $pctText.Text = "$pct%"; $pctText.FontSize = 52; $pctText.FontWeight = "ExtraBold"
    $pctText.HorizontalAlignment = "Center"; $pctText.Foreground = $bc.ConvertFromString($gaugeColor)
    $pctText.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    $pctText.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect
    $pctText.Effect.Color = [System.Windows.Media.ColorConverter]::ConvertFromString($gaugeColor)
    $pctText.Effect.BlurRadius = 25; $pctText.Effect.ShadowDepth = 0; $pctText.Effect.Opacity = 0.6
    
    $subLabel = New-Object System.Windows.Controls.TextBlock
    $subLabel.Text = "OF RELEVANT ENDPOINTS"; $subLabel.FontSize = 9
    $subLabel.Foreground = $bc.ConvertFromString("#6c7086"); $subLabel.HorizontalAlignment = "Center"
    $subLabel.Margin = [System.Windows.Thickness]::new(0,2,0,12)
    $subLabel.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI"); $subLabel.FontWeight = "SemiBold"
    
    $progBar = New-Object System.Windows.Controls.ProgressBar
    $progBar.Height = 8; $progBar.Width = 200; $progBar.Minimum = 0; $progBar.Maximum = 100
    $progBar.Value = [Math]::Min($pct, 100)
    $progBar.Background = $bc.ConvertFromString("#232346"); $progBar.Foreground = $bc.ConvertFromString($gaugeColor)
    $progBar.BorderThickness = [System.Windows.Thickness]::new(0)
    
    # Stat pills
    $pillPanel1 = New-Object System.Windows.Controls.WrapPanel
    $pillPanel1.HorizontalAlignment = "Center"; $pillPanel1.Margin = [System.Windows.Thickness]::new(0,14,0,0)
    $pillPanel2 = New-Object System.Windows.Controls.WrapPanel
    $pillPanel2.HorizontalAlignment = "Center"; $pillPanel2.Margin = [System.Windows.Thickness]::new(0,2,0,0)
    
    $pillData = @(
        @{Text="Fixed $($Data.StatusCounts.Fixed)"; Fg="#a6e3a1"; Bg="#1a2e1a"; Panel=1},
        @{Text="Failed $($Data.StatusCounts.Failed)"; Fg="#f38ba8"; Bg="#2e1a1a"; Panel=1},
        @{Text="Running $($Data.StatusCounts.Running)"; Fg="#f9e2af"; Bg="#2e2a1a"; Panel=1},
        @{Text="Pending $($Data.StatusCounts.Pending)"; Fg="#89b4fa"; Bg="#1a1a2e"; Panel=2},
        @{Text="N/R $($Data.StatusCounts.NotRelevant)"; Fg="#6c7086"; Bg="#1a1a1e"; Panel=2},
        @{Text="Expired $($Data.StatusCounts.Expired)"; Fg="#fab387"; Bg="#2e1e1a"; Panel=2}
    )
    foreach ($p in $pillData) {
        $pillBorder = New-Object System.Windows.Controls.Border
        $pillBorder.Background = $bc.ConvertFromString($p.Bg)
        $pillBorder.CornerRadius = [System.Windows.CornerRadius]::new(12)
        $pillBorder.Padding = [System.Windows.Thickness]::new(8,3,8,3)
        $pillBorder.Margin = [System.Windows.Thickness]::new(3)
        $pillText = New-Object System.Windows.Controls.TextBlock
        $pillText.Text = $p.Text; $pillText.FontSize = 11; $pillText.FontWeight = "SemiBold"
        $pillText.Foreground = $bc.ConvertFromString($p.Fg)
        $pillBorder.Child = $pillText
        if ($p.Panel -eq 1) { $pillPanel1.Children.Add($pillBorder) | Out-Null }
        else { $pillPanel2.Children.Add($pillBorder) | Out-Null }
    }
    
    $gaugeStack.Children.Add($pctText) | Out-Null
    $gaugeStack.Children.Add($subLabel) | Out-Null
    $gaugeStack.Children.Add($progBar) | Out-Null
    $gaugeStack.Children.Add($pillPanel1) | Out-Null
    $gaugeStack.Children.Add($pillPanel2) | Out-Null
    
    $gaugeGrid.Children.Add($gaugeTitle) | Out-Null
    $gaugeGrid.Children.Add($gaugeSub) | Out-Null
    $gaugeGrid.Children.Add($gaugeStack) | Out-Null
    $gaugeBorder.Child = $gaugeGrid
    
    # ── Timeline Panel ──
    $timeBorder = New-Object System.Windows.Controls.Border
    $timeBorder.Background = $bc.ConvertFromString("#141425")
    $timeBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $timeBorder.BorderBrush = $bc.ConvertFromString("#232346")
    $timeBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    [System.Windows.Controls.Grid]::SetColumn($timeBorder, 2)
    
    $timeGrid = New-Object System.Windows.Controls.Grid
    
    $timeHeader = New-Object System.Windows.Controls.StackPanel
    $timeHeader.Orientation = "Horizontal"
    $timeHeader.Margin = [System.Windows.Thickness]::new(14,8,0,0)
    $timeHeader.VerticalAlignment = "Top"
    
    $timeLabel = New-Object System.Windows.Controls.TextBlock
    $timeLabel.Text = "TIMELINE:"; $timeLabel.FontSize = 11; $timeLabel.FontWeight = "Bold"
    $timeLabel.Foreground = $bc.ConvertFromString("#a6e3a1")
    $timeLabel.VerticalAlignment = "Center"; $timeLabel.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $timeLabel.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    
    $btnFixed = New-Object System.Windows.Controls.Button
    $btnFixed.Content = "Fixed"; $btnFixed.FontSize = 10; $btnFixed.Padding = [System.Windows.Thickness]::new(10,3,10,3)
    $btnFixed.Background = $bc.ConvertFromString("#a6e3a1"); $btnFixed.Foreground = $bc.ConvertFromString("#1e1e2e")
    $btnFixed.Margin = [System.Windows.Thickness]::new(0,0,4,0); $btnFixed.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnFixed.BorderThickness = [System.Windows.Thickness]::new(0)
    
    $btnFailed = New-Object System.Windows.Controls.Button
    $btnFailed.Content = "Failed"; $btnFailed.FontSize = 10; $btnFailed.Padding = [System.Windows.Thickness]::new(10,3,10,3)
    $btnFailed.Background = $bc.ConvertFromString("#313244"); $btnFailed.Foreground = $bc.ConvertFromString("#f38ba8")
    $btnFailed.Margin = [System.Windows.Thickness]::new(0,0,4,0); $btnFailed.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnFailed.BorderThickness = [System.Windows.Thickness]::new(0)
    
    $timeHeader.Children.Add($timeLabel) | Out-Null
    $timeHeader.Children.Add($btnFixed) | Out-Null
    $timeHeader.Children.Add($btnFailed) | Out-Null
    
    $timeSub = New-Object System.Windows.Controls.TextBlock
    $timeSub.Text = $ActionLabel; $timeSub.FontSize = 9
    $timeSub.Foreground = $bc.ConvertFromString("#6c7086")
    $timeSub.Margin = [System.Windows.Thickness]::new(14,28,14,0)
    $timeSub.VerticalAlignment = "Top"; $timeSub.TextTrimming = "CharacterEllipsis"
    
    $timeHost = New-Object System.Windows.Forms.Integration.WindowsFormsHost
    $timeHost.Margin = [System.Windows.Thickness]::new(8,42,8,8)
    $timelineChart = New-TimelineChart
    $timeHost.Child = $timelineChart
    
    $timeGrid.Children.Add($timeHeader) | Out-Null
    $timeGrid.Children.Add($timeSub) | Out-Null
    $timeGrid.Children.Add($timeHost) | Out-Null
    $timeBorder.Child = $timeGrid
    
    # Wire up Fixed/Failed toggle
    $btnFixed.Add_Click({
        $tk = $tabKey
        $td = $script:TabData[$tk]
        if ($td) {
            $td.Filter = "Fixed"
            $btnFixed.Background = $bc.ConvertFromString("#a6e3a1")
            $btnFixed.Foreground = $bc.ConvertFromString("#1e1e2e")
            $btnFailed.Background = $bc.ConvertFromString("#313244")
            $btnFailed.Foreground = $bc.ConvertFromString("#f38ba8")
            Update-TimelineChartObj -Chart $timelineChart -Endpoints $td.Data.Endpoints -FilterStatus "Fixed"
        }
    }.GetNewClosure())
    
    $btnFailed.Add_Click({
        $tk = $tabKey
        $td = $script:TabData[$tk]
        if ($td) {
            $td.Filter = "Failed"
            $btnFailed.Background = $bc.ConvertFromString("#f38ba8")
            $btnFailed.Foreground = $bc.ConvertFromString("#1e1e2e")
            $btnFixed.Background = $bc.ConvertFromString("#313244")
            $btnFixed.Foreground = $bc.ConvertFromString("#a6e3a1")
            Update-TimelineChartObj -Chart $timelineChart -Endpoints $td.Data.Endpoints -FilterStatus "Failed"
        }
    }.GetNewClosure())
    
    $chartsGrid.Children.Add($donutBorder) | Out-Null
    $chartsGrid.Children.Add($gaugeBorder) | Out-Null
    $chartsGrid.Children.Add($timeBorder) | Out-Null
    
    # ── Data Grid ──
    $gridBorder = New-Object System.Windows.Controls.Border
    $gridBorder.Background = $bc.ConvertFromString("#141425")
    $gridBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $gridBorder.Padding = [System.Windows.Thickness]::new(8)
    $gridBorder.BorderBrush = $bc.ConvertFromString("#232346")
    $gridBorder.BorderThickness = [System.Windows.Thickness]::new(1)
    [System.Windows.Controls.Grid]::SetRow($gridBorder, 1)
    
    $gridInner = New-Object System.Windows.Controls.Grid
    $gridTitle = New-Object System.Windows.Controls.TextBlock
    $gridTitle.Text = "ENDPOINT DETAILS"; $gridTitle.FontSize = 14; $gridTitle.FontWeight = "SemiBold"
    $gridTitle.Foreground = $bc.ConvertFromString("#89b4fa")
    $gridTitle.Margin = [System.Windows.Thickness]::new(4,2,0,4); $gridTitle.VerticalAlignment = "Top"
    $gridTitle.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    
    $dg = New-Object System.Windows.Controls.DataGrid
    $dg.Margin = [System.Windows.Thickness]::new(0,26,0,0)
    $dg.AutoGenerateColumns = $false; $dg.IsReadOnly = $true; $dg.CanUserSortColumns = $true
    $dg.SelectionMode = "Single"; $dg.FontSize = 12
    $dg.Background = $bc.ConvertFromString("#0f0f1a"); $dg.Foreground = $bc.ConvertFromString("#cdd6f4")
    $dg.BorderBrush = $bc.ConvertFromString("#313244")
    $dg.RowBackground = $bc.ConvertFromString("#141425")
    $dg.AlternatingRowBackground = $bc.ConvertFromString("#1a1a2e")
    $dg.GridLinesVisibility = "Horizontal"
    $dg.HorizontalGridLinesBrush = $bc.ConvertFromString("#232336")
    $dg.HeadersVisibility = "Column"
    
    @(
        @{H="Computer Name"; B="ComputerName"; W="*"},
        @{H="Status"; B="Status"; W="100"},
        @{H="Start Time"; B="StartTime"; W="150"},
        @{H="End Time"; B="EndTime"; W="150"},
        @{H="Apply Count"; B="ApplyCount"; W="85"},
        @{H="Retry Count"; B="RetryCount"; W="85"}
    ) | ForEach-Object {
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Header = $_.H
        $col.Binding = New-Object System.Windows.Data.Binding($_.B)
        if ($_.W -eq "*") { $col.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star) }
        else { $col.Width = [int]$_.W }
        $dg.Columns.Add($col)
    }
    $dg.ItemsSource = $Data.Endpoints
    
    $gridInner.Children.Add($gridTitle) | Out-Null
    $gridInner.Children.Add($dg) | Out-Null
    $gridBorder.Child = $gridInner
    
    $mainGrid.Children.Add($chartsGrid) | Out-Null
    $mainGrid.Children.Add($gridBorder) | Out-Null
    
    $tabItem.Content = $mainGrid
    
    # Populate charts
    Update-DonutChartObj -Chart $donutChart -StatusCounts $Data.StatusCounts -Total $Data.Total
    Update-TimelineChartObj -Chart $timelineChart -Endpoints $Data.Endpoints -FilterStatus "Fixed"
    
    return $tabItem
}

# ─── API Functions ────────────────────────────────────────────────────────────

function Get-BigFixCredential {
    $secPass = $txtPass.SecurePassword
    return New-Object System.Management.Automation.PSCredential($txtUser.Text, $secPass)
}

function Invoke-BigFixAPI {
    param([string]$Endpoint)
    $uri = "$($txtServer.Text.TrimEnd('/'))$($script:Config.ApiBase)$Endpoint"
    $cred = Get-BigFixCredential
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
        return Invoke-RestMethod -Uri $uri -Credential $cred -Method Get -ContentType "application/xml"
    } catch { throw "API Error: $($_.Exception.Message)" }
}

function Get-ActionStatus {
    param([string]$ActionId)
    $lblStatus.Text = "Fetching action $ActionId..."
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    $title = "(unknown)"
    try {
        $info = Invoke-BigFixAPI "/action/$ActionId"
        $title = $info.BES.SingleAction.Title
    } catch {}
    $statusData = Invoke-BigFixAPI "/action/$ActionId/status"
    $parsed = Parse-StatusData $statusData
    $parsed.Title = $title
    return $parsed
}

function Parse-StatusData {
    param($StatusData)
    $endpoints = @()
    $sc = @{ Fixed=0; Failed=0; Running=0; Pending=0; NotRelevant=0; Expired=0; Other=0 }
    $computers = $StatusData.BESAPI.ActionResults.Computer
    if (-not $computers) { $computers = $StatusData.SelectNodes("//Computer") }
    foreach ($c in $computers) {
        $name = $c.Name; if (-not $name) { $name = $c.GetAttribute("Name") }
        $status = $c.Status; if (-not $status) { $status = ($c.SelectSingleNode("Status")).'#text' }
        $mapped = switch -Wildcard ($status) {
            "*Fixed*"        { "Fixed"; $sc.Fixed++; break }
            "*Failed*"       { "Failed"; $sc.Failed++; break }
            "*Running*"      { "Running"; $sc.Running++; break }
            "*Evaluating*"   { "Running"; $sc.Running++; break }
            "*Waiting*"      { "Pending"; $sc.Pending++; break }
            "*Pending*"      { "Pending"; $sc.Pending++; break }
            "*Not Relevant*" { "Not Relevant"; $sc.NotRelevant++; break }
            "*Expired*"      { "Expired"; $sc.Expired++; break }
            default          { $status; $sc.Other++; break }
        }
        $endpoints += [PSCustomObject]@{
            ComputerName=$name; Status=$mapped; StartTime=$c.StartTime
            EndTime=$c.EndTime; ApplyCount=$c.ApplyCount; RetryCount=$c.RetryCount
        }
    }
    return @{ Endpoints=$endpoints; StatusCounts=$sc; Total=$endpoints.Count; Title="" }
}

# ─── Demo Data ────────────────────────────────────────────────────────────────

function Get-DemoData {
    param([string]$Label = "Windows 11 24H2 Feature Update", [int]$Count = 150)
    $statuses = @('Fixed','Fixed','Fixed','Fixed','Fixed','Fixed','Fixed',
                  'Failed','Failed','Running','Running','Running',
                  'Pending','Pending','Not Relevant','Not Relevant','Expired')
    $prefixes = @('WKS','SRV','LAP','DTK','VDI')
    $sites = @('NYC','CHI','DAL','SEA','ATL','MIA','DEN','BOS','PHX','SFO')
    $baseTime = (Get-Date).AddDays(-3)
    $endpoints = @()
    for ($i = 1; $i -le $Count; $i++) {
        $status = $statuses[(Get-Random -Maximum $statuses.Count)]
        $name = "$($prefixes[(Get-Random -Maximum $prefixes.Count)])-$($sites[(Get-Random -Maximum $sites.Count)])-$('{0:D4}' -f $i)"
        $start = $baseTime.AddMinutes((Get-Random -Minimum 0 -Maximum 4320))
        $end = $null
        if ($status -eq 'Fixed') { $end = $start.AddMinutes((Get-Random -Minimum 2 -Maximum 180)).ToString("yyyy-MM-dd HH:mm:ss") }
        elseif ($status -eq 'Failed') { $end = $start.AddMinutes((Get-Random -Minimum 1 -Maximum 60)).ToString("yyyy-MM-dd HH:mm:ss") }
        $endpoints += [PSCustomObject]@{
            ComputerName=$name; Status=$status; StartTime=$start.ToString("yyyy-MM-dd HH:mm:ss")
            EndTime=$end
            ApplyCount=$(if($status -eq 'Fixed'){(Get-Random -Minimum 1 -Maximum 4).ToString()}else{"0"})
            RetryCount=$(if($status -eq 'Failed'){(Get-Random -Minimum 1 -Maximum 5).ToString()}else{"0"})
        }
    }
    $sc = @{
        Fixed=($endpoints|Where-Object Status -eq 'Fixed').Count
        Failed=($endpoints|Where-Object Status -eq 'Failed').Count
        Running=($endpoints|Where-Object Status -eq 'Running').Count
        Pending=($endpoints|Where-Object Status -eq 'Pending').Count
        NotRelevant=($endpoints|Where-Object Status -eq 'Not Relevant').Count
        Expired=($endpoints|Where-Object Status -eq 'Expired').Count
        Other=0
    }
    return @{ Endpoints=$endpoints; StatusCounts=$sc; Total=$endpoints.Count; Title=$Label }
}

# ─── Chart Update Functions ───────────────────────────────────────────────────

function Update-DonutChartObj {
    param($Chart, $StatusCounts, $Total)
    
    $Chart.Series["Status"].Points.Clear()
    $items = @(
        @{Name="Fixed"; Value=$StatusCounts.Fixed; Color="#a6e3a1"},
        @{Name="Failed"; Value=$StatusCounts.Failed; Color="#f38ba8"},
        @{Name="Running"; Value=$StatusCounts.Running; Color="#f9e2af"},
        @{Name="Pending"; Value=$StatusCounts.Pending; Color="#89b4fa"},
        @{Name="Not Relevant"; Value=$StatusCounts.NotRelevant; Color="#6c7086"},
        @{Name="Expired"; Value=$StatusCounts.Expired; Color="#fab387"}
    )
    foreach ($item in $items) {
        if ($item.Value -gt 0) {
            $pt = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
            $pt.SetValueY($item.Value)
            $pt.AxisLabel = ""
            $pt.LegendText = "$($item.Name) ($($item.Value))"
            $pt.Color = [System.Drawing.ColorTranslator]::FromHtml($item.Color)
            $pt.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#141425")
            $pt.BorderWidth = 2
            $pct = [math]::Round(($item.Value / $Total) * 100, 1)
            $pt.Label = if ($pct -ge 5) { "$($item.Value)" } else { "" }
            $pt.LabelForeColor = [System.Drawing.Color]::White
            $pt.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $pt.ToolTip = "$($item.Name): $($item.Value) ($pct%)"
            $Chart.Series["Status"].Points.Add($pt)
        }
    }
}

function Update-TimelineChartObj {
    param($Chart, $Endpoints, [string]$FilterStatus = "Fixed")
    
    $Chart.Series["Completions"].Points.Clear()
    $Chart.Series["Area"].Points.Clear()
    
    $lineS = $Chart.Series["Completions"]
    $areaS = $Chart.Series["Area"]
    $tArea = $Chart.ChartAreas["TimeArea"]
    
    if ($FilterStatus -eq "Failed") {
        $lineS.Color = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
        $areaS.Color = [System.Drawing.Color]::FromArgb(50, 243, 139, 168)
        $tArea.AxisY.Title = "Endpoints Failed"
        $tArea.AxisY.TitleForeColor = [System.Drawing.ColorTranslator]::FromHtml("#f38ba8")
    } else {
        $lineS.Color = [System.Drawing.ColorTranslator]::FromHtml("#a6e3a1")
        $areaS.Color = [System.Drawing.Color]::FromArgb(50, 166, 227, 161)
        $tArea.AxisY.Title = "Endpoints Completed"
        $tArea.AxisY.TitleForeColor = [System.Drawing.ColorTranslator]::FromHtml("#6c7086")
    }
    
    $completed = $Endpoints | Where-Object { $_.EndTime -and $_.Status -eq $FilterStatus } | 
                 Sort-Object { [datetime]$_.EndTime }
    if ($completed.Count -lt 2) { return }
    
    $minTime = [datetime]$completed[0].EndTime
    $maxTime = [datetime]$completed[-1].EndTime
    $span = ($maxTime - $minTime).TotalMinutes
    if ($span -eq 0) { $span = 1 }
    $buckets = [Math]::Min(25, $completed.Count)
    $bucketSize = $span / $buckets
    
    for ($i = 0; $i -le $buckets; $i++) {
        $bucketEnd = $minTime.AddMinutes($i * $bucketSize)
        $label = $bucketEnd.ToString("MM/dd HH:mm")
        $matching = @($completed | Where-Object { [datetime]$_.EndTime -le $bucketEnd })
        $count = [int]$matching.Count
        $lineS.Points.AddXY($label, $count) | Out-Null
        $areaS.Points.AddXY($label, $count) | Out-Null
    }
    $tArea.AxisX.Interval = [Math]::Max(1, [Math]::Floor($buckets / 8))
}

# ─── Event Handlers ───────────────────────────────────────────────────────────

# Demo button removed for production

$btnConnect.Add_Click({
    try {
        $lblStatus.Text = "Testing connection to $($txtServer.Text)..."
        $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        Invoke-BigFixAPI "/login" | Out-Null
        $lblStatus.Text = "[OK] Connected to BigFix server"
        $btnFetch.IsEnabled = $true
    } catch {
        $lblStatus.Text = "[ERR] Connection failed: $($_.Exception.Message)"
    }
})

$btnFetch.Add_Click({
    $input = $txtActionId.Text.Trim()
    if (-not $input) { $lblStatus.Text = "[!] Enter Action ID(s), comma-separated"; return }
    
    $ids = $input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    
    $tabActions.Items.Clear()
    $script:TabData = @{}
    
    $loaded = 0
    foreach ($id in $ids) {
        try {
            $data = Get-ActionStatus -ActionId $id
            # Extract group name from "Update: App 1.2.3: GroupName" format, fallback to truncated title
            $label = ""
            if ($data.Title -match ':.*:(.+)$') {
                $label = $Matches[1].Trim()
            }
            if (-not $label -or $label.Length -lt 2) {
                $label = if ($data.Title.Length -gt 30) { $data.Title.Substring(0,30) + "..." } else { $data.Title }
            }
            $actionLabel = "$($data.Title) (Action $id)"
            $tab = New-ActionTab -TabLabel $label -ActionLabel $actionLabel -Data $data
            $tabActions.Items.Add($tab) | Out-Null
            $loaded++
        } catch {
            $lblStatus.Text = "[ERR] Action $id : $($_.Exception.Message)"
        }
    }
    
    if ($loaded -gt 0) {
        $tabActions.SelectedIndex = 0
        $lblActionName.Text = "$loaded action(s) loaded"
        $btnRefresh.IsEnabled = $true
        $btnExport.IsEnabled = $true
        $lblStatus.Text = "[OK] Loaded $loaded action(s)"
        $lblLastRefresh.Text = "Last refresh: $(Get-Date -Format 'HH:mm:ss')"
    }
})

$btnRefresh.Add_Click({
    $btnFetch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})

$btnExport.Add_Click({
    $selectedTab = $tabActions.SelectedItem
    if (-not $selectedTab) { return }
    $tabKey = $selectedTab.Header
    $td = $script:TabData[$tabKey]
    if (-not $td) { return }
    
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    $dlg.FileName = "BigFix_$($tabKey -replace '\s','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq $true) {
        $td.Data.Endpoints | Export-Csv -Path $dlg.FileName -NoTypeInformation
        $lblStatus.Text = "[OK] Exported $tabKey to $($dlg.FileName)"
    }
})

$txtActionId.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Return) {
        $btnFetch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

# ─── Launch ───────────────────────────────────────────────────────────────────
$window.ShowDialog() | Out-Null
