param(
    [switch]$NoUI
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-Size {
    param([Int64]$Bytes)

    if ($Bytes -lt 1KB) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    return ("{0:N2} GB" -f ($Bytes / 1GB))
}

function Get-PreferredPath {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            return $candidate
        }
    }

    return $null
}

function Join-SafePath {
    param(
        [string]$BasePath,
        [string]$ChildPath
    )

    if ([string]::IsNullOrWhiteSpace($BasePath)) {
        return $null
    }

    return Join-Path -Path $BasePath -ChildPath $ChildPath
}

function Get-AppPaths {
    $homePath = $HOME
    $userProfile = $env:USERPROFILE
    if ([string]::IsNullOrWhiteSpace($userProfile)) {
        $userProfile = $homePath
    }

    $localAppData = Get-PreferredPath @(
        [Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData),
        $env:LOCALAPPDATA,
        $(if ($userProfile) { Join-Path $userProfile 'AppData\Local' })
    )

    $windowsDir = Get-PreferredPath @(
        $env:WINDIR,
        $env:SystemRoot,
        'C:\Windows'
    )

    $tempPath = Get-PreferredPath @(
        $env:TEMP,
        $env:TMP,
        $(if ($localAppData) { Join-Path $localAppData 'Temp' })
    )

    $crashDumps = if ($localAppData) { Join-Path $localAppData 'CrashDumps' } else { $null }
    $explorerCache = if ($localAppData) { Join-Path $localAppData 'Microsoft\Windows\Explorer' } else { $null }

    return [PSCustomObject]@{
        LocalAppData = $localAppData
        WindowsDir = $windowsDir
        Temp = $tempPath
        CrashDumps = $crashDumps
        ExplorerCache = $explorerCache
    }
}

function Get-DirectorySize {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return 0
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return 0
    }

    try {
        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer } |
            Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) { return 0 }
        return [Int64]$sum
    } catch {
        return 0
    }
}

function Get-TotalDirectorySize {
    param([string[]]$Paths)

    $total = 0L
    foreach ($path in ($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $total += Get-DirectorySize -Path $path
    }
    return $total
}

function Get-RecycleBinSize {
    try {
        $shell = New-Object -ComObject Shell.Application
        $bin = $shell.Namespace(10)
        if ($null -eq $bin) { return 0 }
        $total = 0L
        foreach ($item in $bin.Items()) {
            $sizeText = $bin.GetDetailsOf($item, 1)
            if ([string]::IsNullOrWhiteSpace($sizeText)) { continue }
            $cleaned = ($sizeText -replace '[^0-9\,\.]', '')
            $value = 0.0
            if ([double]::TryParse($cleaned, [ref]$value)) {
                $unit = $sizeText.ToUpperInvariant()
                if ($unit -match 'KB') { $total += [Int64]($value * 1KB) }
                elseif ($unit -match 'MB') { $total += [Int64]($value * 1MB) }
                elseif ($unit -match 'GB') { $total += [Int64]($value * 1GB) }
                else { $total += [Int64]$value }
            }
        }
        return $total
    } catch {
        return 0
    }
}

function Remove-DirectoryContents {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return '路径不可用，已跳过。'
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return "路径不存在，已跳过：$Path"
    }

    $items = Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    $removed = 0
    $failed = 0

    foreach ($item in $items) {
        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        } catch {
            $failed++
        }
    }

    return "已处理 $Path，成功 $removed 项，失败 $failed 项。"
}

function Remove-MultiDirectoryContents {
    param([string[]]$Paths)

    $results = foreach ($path in ($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        Remove-DirectoryContents -Path $path
    }

    if (-not $results) {
        return '没有可清理的有效路径。'
    }

    return ($results -join [Environment]::NewLine)
}

function Remove-RecycleBinContents {
    try {
        Clear-RecycleBin -Force -ErrorAction Stop
        return '回收站已清空。'
    } catch {
        return "清空回收站失败：$($_.Exception.Message)"
    }
}

function Remove-ThumbnailCache {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return '缩略图缓存目录不可用，已跳过。'
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return '缩略图缓存目录不存在，已跳过。'
    }

    $removed = 0
    $failed = 0
    Get-ChildItem -LiteralPath $Path -Filter 'thumbcache*' -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            $removed++
        } catch {
            $failed++
        }
    }

    return "缩略图缓存处理完成，成功 $removed 项，失败 $failed 项。"
}

function New-CleanupItem {
    param(
        [string]$Id,
        [string]$Title,
        [int]$Severity,
        [string]$SeverityText,
        [string]$Description,
        [string]$Impact,
        [string]$Kind,
        [string[]]$Paths,
        [bool]$DefaultChecked,
        [bool]$AdminRecommended
    )

    $validPaths = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    return [PSCustomObject]@{
        Id = $Id
        Title = $Title
        Severity = $Severity
        SeverityText = $SeverityText
        Description = $Description
        Impact = $Impact
        Kind = $Kind
        Paths = $validPaths
        DefaultChecked = $DefaultChecked
        AdminRecommended = $AdminRecommended
    }
}

function Get-CleanupItems {
    $paths = Get-AppPaths
    $localAppData = $paths.LocalAppData
    $windowsDir = $paths.WindowsDir

    $edgeCachePath = Join-SafePath -BasePath $localAppData -ChildPath 'Microsoft\Edge\User Data\Default\Cache'
    $chromeCachePath = Join-SafePath -BasePath $localAppData -ChildPath 'Google\Chrome\User Data\Default\Cache'
    $windowsUpdatePath = Join-SafePath -BasePath $windowsDir -ChildPath 'SoftwareDistribution\Download'
    $windowsTempPath = Join-SafePath -BasePath $windowsDir -ChildPath 'Temp'
    $deliveryOptimizationPath = Join-SafePath -BasePath $windowsDir -ChildPath 'SoftwareDistribution\DeliveryOptimization'

    return @(
        (New-CleanupItem -Id 'RecycleBin' -Title '回收站' -Severity 5 -SeverityText '高' -Description '永久删除回收站中的文件。' -Impact '会影响：此前删除但可能还想恢复的个人文件。清理后通常无法直接恢复。' -Kind 'RecycleBin' -Paths @() -DefaultChecked $false -AdminRecommended $false),
        (New-CleanupItem -Id 'WindowsUpdate' -Title 'Windows 更新下载缓存' -Severity 4 -SeverityText '较高' -Description '删除 Windows 更新的已下载临时包。' -Impact '会影响：Windows Update。可能需要重新下载更新文件，个别待安装更新会被重置。' -Kind 'MultiPath' -Paths @($windowsUpdatePath) -DefaultChecked $false -AdminRecommended $true),
        (New-CleanupItem -Id 'SystemTemp' -Title '系统临时目录' -Severity 3 -SeverityText '中' -Description '清理 Windows 系统临时目录。' -Impact '会影响：部分安装器、系统维护任务的临时文件。正在使用中的文件会被跳过。' -Kind 'MultiPath' -Paths @($windowsTempPath) -DefaultChecked $true -AdminRecommended $true),
        (New-CleanupItem -Id 'BrowserCache' -Title '浏览器缓存（Edge / Chrome）' -Severity 3 -SeverityText '中' -Description '清理常见 Chromium 浏览器的缓存文件。' -Impact '会影响：Edge、Chrome 的网页缓存。首次打开网站会更慢，少量站点可能需要重新加载资源。' -Kind 'MultiPath' -Paths @($edgeCachePath, $chromeCachePath) -DefaultChecked $true -AdminRecommended $false),
        (New-CleanupItem -Id 'CrashDumps' -Title '应用崩溃转储文件' -Severity 3 -SeverityText '中' -Description '清理用户目录下的程序崩溃转储文件。' -Impact '会影响：某些软件故障排查。清理后相关程序的历史崩溃转储将不可用于分析。' -Kind 'MultiPath' -Paths @($paths.CrashDumps) -DefaultChecked $true -AdminRecommended $false),
        (New-CleanupItem -Id 'DeliveryOptimization' -Title '传递优化缓存' -Severity 2 -SeverityText '较低' -Description '清理 Windows 传递优化缓存。' -Impact '会影响：Windows 更新的局域网共享缓存。后续更新可能重新生成缓存。' -Kind 'MultiPath' -Paths @($deliveryOptimizationPath) -DefaultChecked $true -AdminRecommended $true),
        (New-CleanupItem -Id 'UserTemp' -Title '当前用户临时目录' -Severity 2 -SeverityText '较低' -Description '清理当前登录用户的临时文件。' -Impact '会影响：部分软件的短期缓存或安装残留。正在使用中的文件会保留。' -Kind 'MultiPath' -Paths @($paths.Temp) -DefaultChecked $true -AdminRecommended $false),
        (New-CleanupItem -Id 'ThumbnailCache' -Title '缩略图缓存' -Severity 1 -SeverityText '低' -Description '删除资源管理器的图片/视频缩略图缓存。' -Impact '会影响：文件夹预览图。系统会自动重建，首次打开图片目录会稍慢。' -Kind 'ThumbnailCache' -Paths @($paths.ExplorerCache) -DefaultChecked $true -AdminRecommended $false)
    ) | Sort-Object -Property Severity, Title -Descending
}

function Get-EstimateForItem {
    param($Item)

    switch ($Item.Kind) {
        'RecycleBin' { return Get-RecycleBinSize }
        'ThumbnailCache' { return Get-TotalDirectorySize -Paths $Item.Paths }
        default { return Get-TotalDirectorySize -Paths $Item.Paths }
    }
}

function Invoke-CleanupForItem {
    param($Item)

    switch ($Item.Kind) {
        'RecycleBin' { return Remove-RecycleBinContents }
        'ThumbnailCache' { return Remove-ThumbnailCache -Path ($Item.Paths | Select-Object -First 1) }
        default { return Remove-MultiDirectoryContents -Paths $Item.Paths }
    }
}

function Get-PathSummary {
    param($Item)

    if (-not $Item.Paths -or $Item.Paths.Count -eq 0) {
        return '路径：系统对象'
    }

    return '路径：' + ($Item.Paths -join '  |  ')
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="C盘空间清理工具"
        Width="1120"
        Height="780"
        MinWidth="980"
        MinHeight="680"
        WindowStartupLocation="CenterScreen"
        Background="#F5F7FA">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border Background="White" CornerRadius="12" Padding="18" BorderBrush="#D8DEE9" BorderThickness="1">
            <StackPanel>
                <TextBlock FontSize="26" FontWeight="Bold" Foreground="#1F2937" Text="C盘空间清理工具" />
                <TextBlock Margin="0,8,0,0" FontSize="14" Foreground="#4B5563" TextWrapping="Wrap"
                           Text="按风险从高到低排序。勾选后可一键清理；系统级目录建议以管理员身份运行。高风险项默认不勾选，请先确认影响说明。" />
            </StackPanel>
        </Border>

        <Border Grid.Row="1" Margin="0,14,0,14" Background="#FFF7ED" CornerRadius="12" Padding="14" BorderBrush="#FDBA74" BorderThickness="1">
            <StackPanel>
                <TextBlock x:Name="AdminStateText" FontSize="14" FontWeight="SemiBold" Foreground="#9A3412" />
                <TextBlock Margin="0,6,0,0" FontSize="13" Foreground="#9A3412" TextWrapping="Wrap"
                           Text="说明：无法删除的文件通常是正在使用、权限不足，或属于系统保护范围，工具会自动跳过并显示结果。" />
            </StackPanel>
        </Border>

        <Border Grid.Row="2" Background="White" CornerRadius="12" Padding="16" BorderBrush="#D8DEE9" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                    <RowDefinition Height="110" />
                </Grid.RowDefinitions>

                <DockPanel LastChildFill="False">
                    <TextBlock DockPanel.Dock="Left" FontSize="18" FontWeight="Bold" Foreground="#111827" Text="清理项" />
                    <Button x:Name="RefreshButton" DockPanel.Dock="Right" Width="100" Height="34" Margin="10,0,0,0"
                            Background="#E5E7EB" Foreground="#111827" BorderBrush="#D1D5DB" Content="刷新大小" />
                    <Button x:Name="SelectSafeButton" DockPanel.Dock="Right" Width="120" Height="34"
                            Background="#E0F2FE" Foreground="#075985" BorderBrush="#BAE6FD" Content="勾选推荐项" />
                </DockPanel>

                <ScrollViewer Grid.Row="1" Margin="0,14,0,14" VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="ItemPanel" />
                </ScrollViewer>

                <TextBox x:Name="LogBox" Grid.Row="2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                         Background="#0F172A" Foreground="#E2E8F0" BorderThickness="0" Padding="10"
                         FontFamily="Consolas" FontSize="12" />
            </Grid>
        </Border>

        <DockPanel Grid.Row="3" Margin="0,16,0,0" LastChildFill="False">
            <TextBlock x:Name="SummaryText" DockPanel.Dock="Left" VerticalAlignment="Center" FontSize="14" Foreground="#374151" />
            <Button x:Name="CleanButton" DockPanel.Dock="Right" Width="180" Height="44"
                    Background="#DC2626" Foreground="White" BorderBrush="#DC2626" Content="一键清理已勾选项" />
        </DockPanel>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$itemPanel = $window.FindName('ItemPanel')
$logBox = $window.FindName('LogBox')
$summaryText = $window.FindName('SummaryText')
$cleanButton = $window.FindName('CleanButton')
$refreshButton = $window.FindName('RefreshButton')
$selectSafeButton = $window.FindName('SelectSafeButton')
$adminStateText = $window.FindName('AdminStateText')

$script:isAdmin = Test-IsAdministrator
$script:cleanupItems = Get-CleanupItems
$script:uiRows = @()

if ($script:isAdmin) {
    $adminStateText.Text = '当前状态：已使用管理员权限运行。'
} else {
    $adminStateText.Text = '当前状态：未使用管理员权限运行，系统级目录可能无法完全清理。'
}

function Add-Log {
    param([string]$Text)

    $stamp = Get-Date -Format 'HH:mm:ss'
    $logBox.AppendText("[$stamp] $Text`r`n")
    $logBox.ScrollToEnd()
}

function Get-SeverityBrush {
    param([int]$Severity)

    switch ($Severity) {
        5 { return '#FEE2E2' }
        4 { return '#FFEDD5' }
        3 { return '#FEF3C7' }
        2 { return '#DBEAFE' }
        default { return '#DCFCE7' }
    }
}

function Get-SeverityTextBrush {
    param([int]$Severity)

    switch ($Severity) {
        5 { return '#991B1B' }
        4 { return '#9A3412' }
        3 { return '#92400E' }
        2 { return '#1D4ED8' }
        default { return '#166534' }
    }
}

function Update-Summary {
    $selectedRows = $script:uiRows | Where-Object { $_.CheckBox.IsChecked }
    $total = ($selectedRows | Measure-Object -Property EstimateBytes -Sum).Sum
    if ($null -eq $total) { $total = 0 }
    $count = ($selectedRows | Measure-Object).Count
    $summaryText.Text = "已勾选 $count 项，预计可释放 $(Format-Size ([Int64]$total))。"
}

function Build-ItemUi {
    $itemPanel.Children.Clear()
    $script:uiRows = @()

    foreach ($item in $script:cleanupItems) {
        $estimate = [Int64](Get-EstimateForItem -Item $item)

        $border = [System.Windows.Controls.Border]::new()
        $border.CornerRadius = [System.Windows.CornerRadius]::new(10)
        $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E5E7EB')
        $border.BorderThickness = [System.Windows.Thickness]::new(1)
        $border.Padding = [System.Windows.Thickness]::new(16)
        $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
        $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('White')

        $stack = [System.Windows.Controls.StackPanel]::new()

        $headerGrid = [System.Windows.Controls.Grid]::new()
        $null = $headerGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $null = $headerGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $headerGrid.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $headerGrid.ColumnDefinitions[1].Width = [System.Windows.GridLength]::new(180)

        $mainPanel = [System.Windows.Controls.StackPanel]::new()
        [System.Windows.Controls.Grid]::SetColumn($mainPanel, 0)

        $checkBox = [System.Windows.Controls.CheckBox]::new()
        $checkBox.IsChecked = $item.DefaultChecked
        $checkBox.FontSize = 15
        $checkBox.FontWeight = 'SemiBold'
        $checkBox.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#111827')
        $checkBox.Padding = [System.Windows.Thickness]::new(2)
        $checkBox.Margin = [System.Windows.Thickness]::new(0, 0, 12, 0)
        $checkBox.Content = $item.Title
        $null = $mainPanel.Children.Add($checkBox)

        $descText = [System.Windows.Controls.TextBlock]::new()
        $descText.Text = $item.Description
        $descText.FontSize = 13
        $descText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4B5563')
        $descText.TextWrapping = 'Wrap'
        $descText.Margin = [System.Windows.Thickness]::new(28, 6, 10, 0)
        $null = $mainPanel.Children.Add($descText)

        $null = $headerGrid.Children.Add($mainPanel)

        $severityLabel = [System.Windows.Controls.Border]::new()
        [System.Windows.Controls.Grid]::SetColumn($severityLabel, 1)
        $severityLabel.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString((Get-SeverityBrush -Severity $item.Severity))
        $severityLabel.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $severityLabel.Padding = [System.Windows.Thickness]::new(10, 8, 10, 8)
        $severityLabel.Margin = [System.Windows.Thickness]::new(10, 0, 0, 0)
        $severityLabel.HorizontalAlignment = 'Right'
        $severityLabel.VerticalAlignment = 'Top'

        $severityText = [System.Windows.Controls.TextBlock]::new()
        $severityText.Text = "影响等级：$($item.SeverityText)"
        $severityText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString((Get-SeverityTextBrush -Severity $item.Severity))
        $severityText.FontWeight = 'SemiBold'
        $severityText.TextAlignment = 'Center'
        $severityLabel.Child = $severityText
        $null = $headerGrid.Children.Add($severityLabel)

        $null = $stack.Children.Add($headerGrid)

        $impactText = [System.Windows.Controls.TextBlock]::new()
        $impactText.Text = $item.Impact
        $impactText.Margin = [System.Windows.Thickness]::new(28, 10, 0, 0)
        $impactText.TextWrapping = 'Wrap'
        $impactText.FontSize = 13
        $impactText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#7C2D12')
        $null = $stack.Children.Add($impactText)

        $pathText = [System.Windows.Controls.TextBlock]::new()
        $pathText.Text = Get-PathSummary -Item $item
        $pathText.Margin = [System.Windows.Thickness]::new(28, 8, 0, 0)
        $pathText.TextWrapping = 'Wrap'
        $pathText.FontSize = 12
        $pathText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#6B7280')
        $null = $stack.Children.Add($pathText)

        $metaText = [System.Windows.Controls.TextBlock]::new()
        $metaText.Margin = [System.Windows.Thickness]::new(28, 8, 0, 0)
        $metaText.FontSize = 13
        $metaText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#374151')
        $adminHint = if ($item.AdminRecommended) { ' | 建议管理员运行' } else { '' }
        $metaText.Text = "预计可释放：$(Format-Size $estimate)$adminHint"
        $null = $stack.Children.Add($metaText)

        $border.Child = $stack
        $null = $itemPanel.Children.Add($border)

        $script:uiRows += [PSCustomObject]@{
            Id = $item.Id
            CheckBox = $checkBox
            EstimateBytes = $estimate
            MetaText = $metaText
            Item = $item
        }

        $checkBox.Add_Checked({ Update-Summary })
        $checkBox.Add_Unchecked({ Update-Summary })
    }

    Update-Summary
}

function Refresh-Estimates {
    Add-Log '正在刷新每个清理项的空间估算...'
    foreach ($row in $script:uiRows) {
        $estimate = [Int64](Get-EstimateForItem -Item $row.Item)
        $row.EstimateBytes = $estimate
        $adminHint = if ($row.Item.AdminRecommended) { ' | 建议管理员运行' } else { '' }
        $row.MetaText.Text = "预计可释放：$(Format-Size $estimate)$adminHint"
    }
    Update-Summary
    Add-Log '空间估算已刷新。'
}

$refreshButton.Add_Click({ Refresh-Estimates })

$selectSafeButton.Add_Click({
    foreach ($row in $script:uiRows) {
        $row.CheckBox.IsChecked = [bool]$row.Item.DefaultChecked
    }
    Update-Summary
    Add-Log '已恢复为推荐勾选项。'
})

$cleanButton.Add_Click({
    $selectedRows = $script:uiRows | Where-Object { $_.CheckBox.IsChecked }
    if (-not $selectedRows) {
        [System.Windows.MessageBox]::Show('请先勾选至少一个清理项。', '未选择项目', 'OK', 'Warning') | Out-Null
        return
    }

    $totalBytes = ($selectedRows | Measure-Object -Property EstimateBytes -Sum).Sum
    if ($null -eq $totalBytes) { $totalBytes = 0 }
    $message = "将清理 $($selectedRows.Count) 项，预计释放 $(Format-Size ([Int64]$totalBytes))。`n`n高风险项目可能影响文件恢复、更新缓存或软件缓存。是否继续？"
    $result = [System.Windows.MessageBox]::Show($message, '确认一键清理', 'YesNo', 'Warning')
    if ($result -ne 'Yes') {
        Add-Log '用户取消了清理。'
        return
    }

    Add-Log '开始执行一键清理...'
    $cleanButton.IsEnabled = $false
    $refreshButton.IsEnabled = $false
    $selectSafeButton.IsEnabled = $false

    foreach ($row in $selectedRows) {
        Add-Log "正在处理：$($row.Item.Title)"
        try {
            $resultText = Invoke-CleanupForItem -Item $row.Item
            Add-Log $resultText
        } catch {
            Add-Log "处理失败：$($row.Item.Title) - $($_.Exception.Message)"
        }
    }

    Add-Log '一键清理完成，正在刷新空间估算。'
    Refresh-Estimates
    $cleanButton.IsEnabled = $true
    $refreshButton.IsEnabled = $true
    $selectSafeButton.IsEnabled = $true
    [System.Windows.MessageBox]::Show('清理已完成，详细结果请查看下方日志。', '完成', 'OK', 'Information') | Out-Null
})

if ($NoUI) {
    $script:cleanupItems | ForEach-Object {
        [PSCustomObject]@{
            Title = $_.Title
            SeverityText = $_.SeverityText
            Estimate = Format-Size (Get-EstimateForItem -Item $_)
            Paths = if ($_.Paths.Count -eq 0) { '系统对象' } else { ($_.Paths -join '; ') }
        }
    } | Format-Table -Wrap -AutoSize
    return
}

Build-ItemUi
Add-Log '工具已启动。清理项已按影响等级从高到低排序。'

$null = $window.ShowDialog()

