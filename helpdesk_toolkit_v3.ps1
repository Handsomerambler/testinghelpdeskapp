#requires -Version 5.1
<#
.SYNOPSIS
    Help Desk Toolkit v3

.DESCRIPTION
    Professional PowerShell toolkit for common help desk, desktop support, and junior sysadmin workflows.
    Includes computer diagnostics, printer support, software inventory, service health, Active Directory,
    Microsoft Entra / Graph lookups, and exportable reporting.

.NOTES
    Author: Kishan / Handsomerambler
    Version: 3.0.0
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolkitVersion = '3.0.0'
$script:ExportFolder = 'C:\Temp\HelpDeskToolkit'
$script:LogFile = Join-Path $script:ExportFolder 'HelpDeskToolkit.log'
$script:TranscriptFile = Join-Path $script:ExportFolder 'ToolkitTranscript.txt'
$script:LastResults = @()

if (-not (Test-Path $script:ExportFolder)) {
    New-Item -Path $script:ExportFolder -ItemType Directory -Force | Out-Null
}

function Write-ToolkitLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $script:LogFile -Value "$timestamp [$Level] $Message"
}

function Add-ToolkitResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$InputObject
    )

    $script:LastResults += $InputObject
    return $InputObject
}

function Show-ToolkitHeader {
    Clear-Host
    Write-Host '=================================================='
    Write-Host '                Help Desk Toolkit v3              '
    Write-Host '=================================================='
    Write-Host "Version     : $script:ToolkitVersion"
    Write-Host "Export Path : $script:ExportFolder"
    Write-Host "Log File    : $script:LogFile"
    Write-Host ''
}

function Pause-Toolkit {
    Write-Host ''
    Read-Host 'Press Enter to continue'
}

function Confirm-ToolkitAction {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    $response = Read-Host "$Prompt (Y/N)"
    return $response -match '^(Y|y)$'
}

function Get-TargetComputer {
    do {
        $computerName = Read-Host 'Enter computer name'
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            Write-Host 'Computer name cannot be blank.'
        }
    } until (-not [string]::IsNullOrWhiteSpace($computerName))

    return $computerName.Trim()
}

function Test-ToolkitModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module is not installed or not available in this session."
        Write-ToolkitLog "$ModuleName module not available" 'WARN'
        return $false
    }

    return $true
}

function Test-PCOnline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $online = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop
        Write-ToolkitLog "Connectivity check for $ComputerName: $online"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Ping Computer'
            ComputerName = $ComputerName
            Online       = $online
        })
    }
    catch {
        Write-ToolkitLog "Connectivity check failed for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Ping Computer'
            ComputerName = $ComputerName
            Online       = $false
            Error        = $_.Exception.Message
        })
    }
}

function Get-LoggedInUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        $user = (Get-CimInstance Win32_ComputerSystem -ComputerName $ComputerName).UserName
        Write-ToolkitLog "Retrieved logged-in user for $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get Logged In User'
            ComputerName = $ComputerName
            LoggedInUser = if ($user) { $user } else { 'No user detected' }
        })
    }
    catch {
        Write-ToolkitLog "Failed to retrieve logged-in user for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get Logged In User'
            ComputerName = $ComputerName
            LoggedInUser = 'Unavailable'
            Error        = $_.Exception.Message
        })
    }
}

function Get-PCUptime {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName
        $uptime = (Get-Date) - $os.LastBootUpTime
        Write-ToolkitLog "Retrieved uptime for $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get Uptime'
            ComputerName = $ComputerName
            LastBoot     = $os.LastBootUpTime
            UptimeDays   = [math]::Round($uptime.TotalDays, 2)
        })
    }
    catch {
        Write-ToolkitLog "Failed to retrieve uptime for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get Uptime'
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
        })
    }
}

function Get-DiskSpace {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $ComputerName -Filter "DriveType=3"
        Write-ToolkitLog "Retrieved disk space for $ComputerName"
        foreach ($disk in $disks) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Category     = 'Computer'
                Action       = 'Get Disk Space'
                ComputerName = $ComputerName
                Drive        = $disk.DeviceID
                SizeGB       = [math]::Round($disk.Size / 1GB, 2)
                FreeGB       = [math]::Round($disk.FreeSpace / 1GB, 2)
            })
        }
    }
    catch {
        Write-ToolkitLog "Failed to retrieve disk space for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get Disk Space'
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
        })
    }
}

function Get-PCSystemInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $ComputerName
        $bios = Get-CimInstance Win32_BIOS -ComputerName $ComputerName
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName
        Write-ToolkitLog "Retrieved system info for $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time          = Get-Date
            Category      = 'Computer'
            Action        = 'Get System Info'
            ComputerName  = $ComputerName
            Manufacturer  = $cs.Manufacturer
            Model         = $cs.Model
            SerialNumber  = $bios.SerialNumber
            OS            = $os.Caption
            Version       = $os.Version
            LoggedInUser  = $cs.UserName
            LastBoot      = $os.LastBootUpTime
        })
    }
    catch {
        Write-ToolkitLog "Failed to retrieve system info for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Get System Info'
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
        })
    }
}

function Restart-RemoteComputer {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    if (-not (Confirm-ToolkitAction "Are you sure you want to reboot $ComputerName?")) {
        Write-Host 'Action canceled.'
        Write-ToolkitLog "Reboot canceled for $ComputerName" 'WARN'
        return
    }

    try {
        Restart-Computer -ComputerName $ComputerName -Force
        Write-ToolkitLog "Reboot command sent to $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Reboot Computer'
            ComputerName = $ComputerName
            Status       = 'Reboot command sent'
        })
    }
    catch {
        Write-ToolkitLog "Failed to reboot $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Computer'
            Action       = 'Reboot Computer'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Restart-PrintSpooler {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock { Restart-Service -Name Spooler -Force }
        Write-ToolkitLog "Spooler restarted on $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Restart Spooler'
            ComputerName = $ComputerName
            Status       = 'Success'
        })
    }
    catch {
        Write-ToolkitLog "Failed to restart Spooler on $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Restart Spooler'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Clear-PrintQueue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    if (-not (Confirm-ToolkitAction "Clear all print jobs on $ComputerName?")) { return }

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Stop-Service Spooler -Force
            Remove-Item 'C:\Windows\System32\spool\PRINTERS\*' -Force -ErrorAction SilentlyContinue
            Start-Service Spooler
        }
        Write-ToolkitLog "Print queue cleared on $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Clear Print Queue'
            ComputerName = $ComputerName
            Status       = 'Success'
        })
    }
    catch {
        Write-ToolkitLog "Failed to clear print queue on $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Clear Print Queue'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Get-InstalledPrinters {
    [CmdletBinding()]
    param()

    try {
        Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published | ForEach-Object { Add-ToolkitResult $_ }
        Write-ToolkitLog 'Retrieved installed printers'
    }
    catch {
        Write-ToolkitLog "Failed to retrieve printers. $_" 'ERROR'
    }
}

function Get-RemotePrinters {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published
        } | ForEach-Object {
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Category     = 'Printer'
                Action       = 'Get Remote Printers'
                ComputerName = $ComputerName
                Name         = $_.Name
                DriverName   = $_.DriverName
                PortName     = $_.PortName
                Shared       = $_.Shared
                Published    = $_.Published
            })
        }
        Write-ToolkitLog "Retrieved remote printers from $ComputerName"
    }
    catch {
        Write-ToolkitLog "Failed remote printer list for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Get Remote Printers'
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
        })
    }
}

function Remove-RemotePrinter {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName,[Parameter(Mandatory)][string]$PrinterName)

    if (-not (Confirm-ToolkitAction "Remove printer $PrinterName from $ComputerName?")) { return }

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            param($p)
            Remove-Printer -Name $p -ErrorAction Stop
        } -ArgumentList $PrinterName
        Write-ToolkitLog "Removed printer $PrinterName from $ComputerName"
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Remove Remote Printer'
            ComputerName = $ComputerName
            PrinterName  = $PrinterName
            Status       = 'Removed'
        })
    }
    catch {
        Write-ToolkitLog "Failed to remove printer $PrinterName from $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Category     = 'Printer'
            Action       = 'Remove Remote Printer'
            ComputerName = $ComputerName
            PrinterName  = $PrinterName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Test-CommonPorts {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ComputerName)

    $ports = 3389, 445, 443, 80
    foreach ($port in $ports) {
        try {
            $result = Test-NetConnection -ComputerName $ComputerName -Port $port -WarningAction SilentlyContinue
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Category     = 'Network'
                Action       = 'Test Port'
                ComputerName = $ComputerName
                Port         = $port
                Open         = $result.TcpTestSucceeded
            })
        }
        catch {
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Category     = 'Network'
                Action       = 'Test Port'
                ComputerName = $ComputerName
                Port         = $port
                Open         = $false
                Error        = $_.Exception.Message
            })
        }
    }
    Write-ToolkitLog "Completed common port test on $ComputerName"
}

function Get-NetworkAdapters {
    [CmdletBinding()]
    param()

    try {
        Get-NetAdapter | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress | ForEach-Object { Add-ToolkitResult $_ }
        Write-ToolkitLog 'Retrieved local network adapters'
    }
    catch {
        Write-ToolkitLog "Failed to retrieve local network adapters. $_" 'ERROR'
    }
}

function Find-ProcessByName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ProcessName)

    try {
        $processes = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object Name, Id, CPU, StartTime
        if (-not $processes) {
            Add-ToolkitResult ([PSCustomObject]@{ Time = Get-Date; Category='Software'; Action='Find Process'; Process=$ProcessName; Status='Not found' })
        }
        else {
            $processes | ForEach-Object { Add-ToolkitResult $_ }
        }
        Write-ToolkitLog "Process search completed for $ProcessName"
    }
    catch {
        Write-ToolkitLog "Failed process lookup for $ProcessName. $_" 'ERROR'
    }
}

function Stop-ProcessByName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ProcessName)

    if (-not (Confirm-ToolkitAction "Stop process $ProcessName?")) { return }

    try {
        Stop-Process -Name $ProcessName -Force
        Write-ToolkitLog "Stopped process $ProcessName"
        Add-ToolkitResult ([PSCustomObject]@{ Time=Get-Date; Category='Software'; Action='Stop Process'; Process=$ProcessName; Status='Stopped' })
    }
    catch {
        Write-ToolkitLog "Failed to stop process $ProcessName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{ Time=Get-Date; Category='Software'; Action='Stop Process'; Process=$ProcessName; Status='Failed'; Error=$_.Exception.Message })
    }
}

function Get-TopMemoryProcesses {
    [CmdletBinding()]
    param()

    try {
        Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 10 Name, Id, @{Name='MemoryMB';Expression={[math]::Round($_.WorkingSet / 1MB, 2)}} | ForEach-Object { Add-ToolkitResult $_ }
        Write-ToolkitLog 'Retrieved top memory processes'
    }
    catch {
        Write-ToolkitLog "Failed to retrieve top memory processes. $_" 'ERROR'
    }
}

function Get-ServiceStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ServiceName)

    try {
        Get-Service -Name $ServiceName | Select-Object Name, DisplayName, Status, StartType | ForEach-Object { Add-ToolkitResult $_ }
        Write-ToolkitLog "Retrieved status for service $ServiceName"
    }
    catch {
        Write-ToolkitLog "Failed to retrieve service $ServiceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{ Time=Get-Date; Category='System Health'; Action='Get Service Status'; Service=$ServiceName; Error=$_.Exception.Message })
    }
}

function Restart-LocalServiceByName {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ServiceName)

    if (-not (Confirm-ToolkitAction "Restart service $ServiceName?")) { return }

    try {
        Restart-Service -Name $ServiceName -Force
        Write-ToolkitLog "Restarted service $ServiceName"
        Add-ToolkitResult ([PSCustomObject]@{ Time=Get-Date; Category='System Health'; Action='Restart Service'; Service=$ServiceName; Status='Restarted' })
    }
    catch {
        Write-ToolkitLog "Failed to restart service $ServiceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{ Time=Get-Date; Category='System Health'; Action='Restart Service'; Service=$ServiceName; Status='Failed'; Error=$_.Exception.Message })
    }
}

function Get-FailedServices {
    [CmdletBinding()]
    param()

    try {
        Get-Service | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' } | Select-Object Name, DisplayName, Status, StartType | ForEach-Object {
            Add-ToolkitResult ([PSCustomObject]@{
                Time        = Get-Date
                Category    = 'System Health'
                Action      = 'Get Failed Services'
                Name        = $_.Name
                DisplayName = $_.DisplayName
                Status      = $_.Status
                StartType   = $_.StartType
            })
        }
        Write-ToolkitLog 'Retrieved failed automatic services'
    }
    catch {
        Write-ToolkitLog "Failed service health check. $_" 'ERROR'
    }
}

function Get-StartupPrograms {
    [CmdletBinding()]
    param()

    try {
        Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User | ForEach-Object {
            Add-ToolkitResult ([PSCustomObject]@{
                Time     = Get-Date
                Category = 'System Health'
                Action   = 'Get Startup Programs'
                Name     = $_.Name
                Command  = $_.Command
                Location = $_.Location
                User     = $_.User
            })
        }
        Write-ToolkitLog 'Retrieved startup programs'
    }
    catch {
        Write-ToolkitLog "Failed startup program check. $_" 'ERROR'
    }
}

function Get-InstalledSoftware {
    [CmdletBinding()]
    param()

    try {
        $paths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        Get-ItemProperty $paths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
            Sort-Object DisplayName |
            ForEach-Object {
                Add-ToolkitResult ([PSCustomObject]@{
                    Time           = Get-Date
                    Category       = 'Software'
                    Action         = 'Get Installed Software'
                    DisplayName    = $_.DisplayName
                    DisplayVersion = $_.DisplayVersion
                    Publisher      = $_.Publisher
                    InstallDate    = $_.InstallDate
                })
            }
        Write-ToolkitLog 'Retrieved installed software inventory'
    }
    catch {
        Write-ToolkitLog "Failed software inventory. $_" 'ERROR'
    }
}

function Find-InstalledSoftware {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    try {
        $paths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        Get-ItemProperty $paths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like "*$Name*" } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
            ForEach-Object {
                Add-ToolkitResult ([PSCustomObject]@{
                    Time           = Get-Date
                    Category       = 'Software'
                    Action         = 'Find Installed Software'
                    Search         = $Name
                    DisplayName    = $_.DisplayName
                    DisplayVersion = $_.DisplayVersion
                    Publisher      = $_.Publisher
                    InstallDate    = $_.InstallDate
                })
            }
        Write-ToolkitLog "Completed software search for $Name"
    }
    catch {
        Write-ToolkitLog "Failed software search for $Name. $_" 'ERROR'
    }
}

function Export-LastResults {
    [CmdletBinding()]
    param()

    if (-not $script:LastResults -or $script:LastResults.Count -eq 0) {
        Write-Host 'No results to export yet.'
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $exportPath = Join-Path $script:ExportFolder "HelpDeskResults_$timestamp.csv"
    try {
        $script:LastResults | Export-Csv -Path $exportPath -NoTypeInformation -Force
        Write-ToolkitLog "Exported results to $exportPath"
        Write-Host "Results exported to $exportPath"
    }
    catch {
        Write-ToolkitLog "Failed to export results. $_" 'ERROR'
    }
}

function Clear-LastResults {
    [CmdletBinding()]
    param()
    $script:LastResults = @()
    Write-ToolkitLog 'Cleared in-memory results'
    Write-Host 'Collected results have been cleared from memory.'
}

function Show-LastResults {
    [CmdletBinding()]
    param()
    if (-not $script:LastResults -or $script:LastResults.Count -eq 0) {
        Write-Host 'No results collected yet.'
        return
    }
    $script:LastResults | Format-Table -AutoSize -Wrap
}

function Show-ComputerMenu {
    do {
        Clear-Host
        Write-Host '--- Computer Tools ---'
        Write-Host '1. Ping computer'
        Write-Host '2. Get logged-in user'
        Write-Host '3. Get uptime'
        Write-Host '4. Get disk space'
        Write-Host '5. Get full system info'
        Write-Host '6. Reboot computer'
        Write-Host '7. Back'
        $choice = Read-Host 'Choose an option'
        switch ($choice) {
            '1' { $c = Get-TargetComputer; Test-PCOnline -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '2' { $c = Get-TargetComputer; Get-LoggedInUser -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '3' { $c = Get-TargetComputer; Get-PCUptime -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '4' { $c = Get-TargetComputer; Get-DiskSpace -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '5' { $c = Get-TargetComputer; Get-PCSystemInfo -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '6' { $c = Get-TargetComputer; Restart-RemoteComputer -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '7' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '7')
}

function Show-PrinterMenu {
    do {
        Clear-Host
        Write-Host '--- Printer Toolkit v2 ---'
        Write-Host '1. Restart spooler on remote PC'
        Write-Host '2. Clear print queue on remote PC'
        Write-Host '3. List local installed printers'
        Write-Host '4. Get remote printers'
        Write-Host '5. Remove remote printer'
        Write-Host '6. Back'
        $choice = Read-Host 'Choose an option'
        switch ($choice) {
            '1' { $c = Get-TargetComputer; Restart-PrintSpooler -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '2' { $c = Get-TargetComputer; Clear-PrintQueue -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '3' { Get-InstalledPrinters | Format-Table -AutoSize; Pause-Toolkit }
            '4' { $c = Get-TargetComputer; Get-RemotePrinters -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '5' { $c = Get-TargetComputer; $p = Read-Host 'Enter printer name'; Remove-RemotePrinter -ComputerName $c -PrinterName $p | Format-Table -AutoSize; Pause-Toolkit }
            '6' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '6')
}

function Show-NetworkMenu {
    do {
        Clear-Host
        Write-Host '--- Network Tools ---'
        Write-Host '1. Test common ports on remote PC'
        Write-Host '2. Show local network adapters'
        Write-Host '3. Back'
        $choice = Read-Host 'Choose an option'
        switch ($choice) {
            '1' { $c = Get-TargetComputer; Test-CommonPorts -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '2' { Get-NetworkAdapters | Format-Table -AutoSize; Pause-Toolkit }
            '3' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '3')
}

function Show-SoftwareHealthMenu {
    do {
        Clear-Host
        Write-Host '--- Software / System Health ---'
        Write-Host '1. Find process by name'
        Write-Host '2. Stop process by name'
        Write-Host '3. Show top memory processes'
        Write-Host '4. Get local service status'
        Write-Host '5. Restart local service'
        Write-Host '6. Get installed software'
        Write-Host '7. Search installed software'
        Write-Host '8. Get startup programs'
        Write-Host '9. Get failed services'
        Write-Host '10. Back'
        $choice = Read-Host 'Choose an option'
        switch ($choice) {
            '1' { $p = Read-Host 'Enter process name'; Find-ProcessByName -ProcessName $p | Format-Table -AutoSize; Pause-Toolkit }
            '2' { $p = Read-Host 'Enter process name to stop'; Stop-ProcessByName -ProcessName $p | Format-Table -AutoSize; Pause-Toolkit }
            '3' { Get-TopMemoryProcesses | Format-Table -AutoSize; Pause-Toolkit }
            '4' { $s = Read-Host 'Enter service name'; Get-ServiceStatus -ServiceName $s | Format-Table -AutoSize; Pause-Toolkit }
            '5' { $s = Read-Host 'Enter service name to restart'; Restart-LocalServiceByName -ServiceName $s | Format-Table -AutoSize; Pause-Toolkit }
            '6' { Get-InstalledSoftware | Format-Table -AutoSize -Wrap; Pause-Toolkit }
            '7' { $n = Read-Host 'Enter software name'; Find-InstalledSoftware -Name $n | Format-Table -AutoSize -Wrap; Pause-Toolkit }
            '8' { Get-StartupPrograms | Format-Table -AutoSize -Wrap; Pause-Toolkit }
            '9' { Get-FailedServices | Format-Table -AutoSize -Wrap; Pause-Toolkit }
            '10' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '10')
}

function Show-ReportingMenu {
    do {
        Clear-Host
        Write-Host '--- Reporting Tools ---'
        Write-Host '1. Show collected results'
        Write-Host '2. Export collected results to CSV'
        Write-Host '3. Clear collected results'
        Write-Host '4. Back'
        $choice = Read-Host 'Choose an option'
        switch ($choice) {
            '1' { Show-LastResults; Pause-Toolkit }
            '2' { Export-LastResults; Pause-Toolkit }
            '3' { Clear-LastResults; Pause-Toolkit }
            '4' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '4')
}

try {
    Start-Transcript -Path $script:TranscriptFile -Append | Out-Null
}
catch {
    Write-ToolkitLog "Could not start transcript. $_" 'WARN'
}

Write-ToolkitLog 'Toolkit started'

try {
    do {
        Show-ToolkitHeader
        Write-Host '1. Computer Tools'
        Write-Host '2. Printer Toolkit v2'
        Write-Host '3. Network Tools'
        Write-Host '4. Software / System Health'
        Write-Host '5. Reporting Tools'
        Write-Host '6. Exit'
        Write-Host ''
        $mainChoice = Read-Host 'Choose an option'
        switch ($mainChoice) {
            '1' { Show-ComputerMenu }
            '2' { Show-PrinterMenu }
            '3' { Show-NetworkMenu }
            '4' { Show-SoftwareHealthMenu }
            '5' { Show-ReportingMenu }
            '6' { Write-Host 'Exiting toolkit...' }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($mainChoice -eq '6')
}
finally {
    Write-ToolkitLog 'Toolkit exited'
    try { Stop-Transcript | Out-Null } catch {}
}
