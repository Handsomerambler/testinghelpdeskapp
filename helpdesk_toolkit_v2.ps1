# Help Desk Toolkit v2
# PowerShell toolkit for common help desk tasks
# Run in a PowerShell session with the permissions required for your environment.

$ErrorActionPreference = 'Stop'
$LogFile = 'C:\Temp\HelpDeskToolkit.log'
$ExportFolder = 'C:\Temp'
$Global:LastResults = @()

if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogFile -Value "$timestamp [$Level] $Message"
}

function Add-ToolkitResult {
    param(
        [Parameter(Mandatory)]
        [psobject]$InputObject
    )

    $Global:LastResults += $InputObject
    $InputObject
}

function Show-Header {
    Clear-Host
    Write-Host '========================================='
    Write-Host '         Help Desk Toolkit v2            '
    Write-Host '========================================='
    Write-Host "Log File: $LogFile"
    Write-Host ''
}

function Pause-Toolkit {
    Write-Host ''
    Read-Host 'Press Enter to continue'
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

function Test-PCOnline {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $online = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop
        Write-Log "Connectivity check for $ComputerName: $online"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Ping Computer'
            ComputerName = $ComputerName
            Online       = $online
        })
    }
    catch {
        Write-Log "Connectivity check failed for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Ping Computer'
            ComputerName = $ComputerName
            Online       = $false
            Error        = $_.Exception.Message
        })
    }
}

function Confirm-Action {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    $response = Read-Host "$Prompt (Y/N)"
    return $response -match '^(Y|y)$'
}

function Get-LoggedInUser {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $user = (Get-CimInstance Win32_ComputerSystem -ComputerName $ComputerName).UserName
        Write-Log "Retrieved logged-in user for $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get Logged In User'
            ComputerName = $ComputerName
            LoggedInUser = if ($user) { $user } else { 'No user detected' }
        })
    }
    catch {
        Write-Log "Failed to retrieve logged-in user for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get Logged In User'
            ComputerName = $ComputerName
            LoggedInUser = 'Unavailable'
            Error        = $_.Exception.Message
        })
    }
}

function Get-PCUptime {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName
        $uptime = (Get-Date) - $os.LastBootUpTime

        Write-Log "Retrieved uptime for $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get Uptime'
            ComputerName = $ComputerName
            LastBoot     = $os.LastBootUpTime
            UptimeDays   = [math]::Round($uptime.TotalDays, 2)
        })
    }
    catch {
        Write-Log "Failed to retrieve uptime for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get Uptime'
            ComputerName = $ComputerName
            LastBoot     = $null
            UptimeDays   = $null
            Error        = $_.Exception.Message
        })
    }
}

function Get-DiskSpace {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $ComputerName -Filter "DriveType=3"
        Write-Log "Retrieved disk space for $ComputerName"

        foreach ($disk in $disks) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Action       = 'Get Disk Space'
                ComputerName = $ComputerName
                Drive        = $disk.DeviceID
                SizeGB       = [math]::Round($disk.Size / 1GB, 2)
                FreeGB       = [math]::Round($disk.FreeSpace / 1GB, 2)
            })
        }
    }
    catch {
        Write-Log "Failed to retrieve disk space for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get Disk Space'
            ComputerName = $ComputerName
            Drive        = 'N/A'
            SizeGB       = $null
            FreeGB       = $null
            Error        = $_.Exception.Message
        })
    }
}

function Get-PCSystemInfo {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $ComputerName
        $bios = Get-CimInstance Win32_BIOS -ComputerName $ComputerName
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $ComputerName

        Write-Log "Retrieved system info for $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time          = Get-Date
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
        Write-Log "Failed to retrieve system info for $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Get System Info'
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
        })
    }
}

function Restart-RemoteComputer {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    if (-not (Confirm-Action "Are you sure you want to reboot $ComputerName?")) {
        Write-Host 'Action canceled.'
        Write-Log "Reboot canceled for $ComputerName" 'WARN'
        return
    }

    try {
        Restart-Computer -ComputerName $ComputerName -Force
        Write-Log "Reboot command sent to $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Reboot Computer'
            ComputerName = $ComputerName
            Status       = 'Reboot command sent'
        })
    }
    catch {
        Write-Log "Failed to reboot $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Reboot Computer'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Restart-PrintSpooler {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Restart-Service -Name Spooler -Force
        }

        Write-Log "Spooler restarted on $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Restart Spooler'
            ComputerName = $ComputerName
            Status       = 'Success'
        })
    }
    catch {
        Write-Log "Failed to restart Spooler on $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Restart Spooler'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Clear-PrintQueue {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    if (-not (Confirm-Action "Clear all print jobs on $ComputerName?")) {
        Write-Host 'Action canceled.'
        Write-Log "Clear print queue canceled for $ComputerName" 'WARN'
        return
    }

    try {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Stop-Service Spooler -Force
            Remove-Item 'C:\Windows\System32\spool\PRINTERS\*' -Force -ErrorAction SilentlyContinue
            Start-Service Spooler
        }

        Write-Log "Print queue cleared on $ComputerName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Clear Print Queue'
            ComputerName = $ComputerName
            Status       = 'Success'
        })
    }
    catch {
        Write-Log "Failed to clear print queue on $ComputerName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Clear Print Queue'
            ComputerName = $ComputerName
            Status       = 'Failed'
            Error        = $_.Exception.Message
        })
    }
}

function Test-CommonPorts {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )

    $ports = 3389, 445, 443, 80

    foreach ($port in $ports) {
        try {
            $result = Test-NetConnection -ComputerName $ComputerName -Port $port -WarningAction SilentlyContinue
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Action       = 'Test Port'
                ComputerName = $ComputerName
                Port         = $port
                Open         = $result.TcpTestSucceeded
            })
        }
        catch {
            Add-ToolkitResult ([PSCustomObject]@{
                Time         = Get-Date
                Action       = 'Test Port'
                ComputerName = $ComputerName
                Port         = $port
                Open         = $false
                Error        = $_.Exception.Message
            })
        }
    }

    Write-Log "Completed common port test on $ComputerName"
}

function Get-NetworkAdapters {
    try {
        $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress
        Write-Log 'Retrieved local network adapters'
        $adapters | ForEach-Object { Add-ToolkitResult $_ }
    }
    catch {
        Write-Log "Failed to retrieve local network adapters. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Network Adapters'
            Error  = $_.Exception.Message
        })
    }
}

function Get-RecentSystemErrors {
    try {
        $events = Get-WinEvent -LogName System -MaxEvents 50 |
            Where-Object { $_.LevelDisplayName -eq 'Error' } |
            Select-Object TimeCreated, Id, ProviderName, Message

        Write-Log 'Retrieved recent system errors'
        $events | ForEach-Object { Add-ToolkitResult $_ }
    }
    catch {
        Write-Log "Failed to retrieve system errors. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Recent System Errors'
            Error  = $_.Exception.Message
        })
    }
}

function Get-RecentApplicationErrors {
    try {
        $events = Get-WinEvent -LogName Application -MaxEvents 50 |
            Where-Object { $_.LevelDisplayName -eq 'Error' } |
            Select-Object TimeCreated, Id, ProviderName, Message

        Write-Log 'Retrieved recent application errors'
        $events | ForEach-Object { Add-ToolkitResult $_ }
    }
    catch {
        Write-Log "Failed to retrieve application errors. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Recent Application Errors'
            Error  = $_.Exception.Message
        })
    }
}

function Find-ProcessByName {
    $processName = Read-Host 'Enter process name'

    if ([string]::IsNullOrWhiteSpace($processName)) {
        Write-Host 'Process name cannot be blank.'
        return
    }

    try {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue |
            Select-Object Name, Id, CPU, StartTime

        if (-not $processes) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time    = Get-Date
                Action  = 'Find Process'
                Process = $processName
                Status  = 'Not found'
            })
        }
        else {
            $processes | ForEach-Object { Add-ToolkitResult $_ }
        }

        Write-Log "Process search completed for $processName"
    }
    catch {
        Write-Log "Failed process lookup for $processName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Find Process'
            Process = $processName
            Error   = $_.Exception.Message
        })
    }
}

function Stop-ProcessByName {
    $processName = Read-Host 'Enter process name to stop'

    if ([string]::IsNullOrWhiteSpace($processName)) {
        Write-Host 'Process name cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Stop process $processName?")) {
        Write-Host 'Action canceled.'
        Write-Log "Stop process canceled for $processName" 'WARN'
        return
    }

    try {
        Stop-Process -Name $processName -Force
        Write-Log "Stopped process $processName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Stop Process'
            Process = $processName
            Status  = 'Stopped'
        })
    }
    catch {
        Write-Log "Failed to stop process $processName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Stop Process'
            Process = $processName
            Status  = 'Failed'
            Error   = $_.Exception.Message
        })
    }
}

function Get-TopMemoryProcesses {
    try {
        $processes = Get-Process |
            Sort-Object WorkingSet -Descending |
            Select-Object -First 10 Name, Id,
            @{Name='MemoryMB';Expression={[math]::Round($_.WorkingSet / 1MB, 2)}}

        Write-Log 'Retrieved top memory processes'
        $processes | ForEach-Object { Add-ToolkitResult $_ }
    }
    catch {
        Write-Log "Failed to retrieve top memory processes. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Top Memory Processes'
            Error  = $_.Exception.Message
        })
    }
}

function Get-ServiceStatus {
    $serviceName = Read-Host 'Enter service name'

    if ([string]::IsNullOrWhiteSpace($serviceName)) {
        Write-Host 'Service name cannot be blank.'
        return
    }

    try {
        $service = Get-Service -Name $serviceName | Select-Object Name, DisplayName, Status, StartType
        Write-Log "Retrieved status for service $serviceName"
        Add-ToolkitResult $service
    }
    catch {
        Write-Log "Failed to retrieve service $serviceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Get Service Status'
            Service = $serviceName
            Error   = $_.Exception.Message
        })
    }
}

function Restart-LocalServiceByName {
    $serviceName = Read-Host 'Enter service name to restart'

    if ([string]::IsNullOrWhiteSpace($serviceName)) {
        Write-Host 'Service name cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Restart service $serviceName?")) {
        Write-Host 'Action canceled.'
        Write-Log "Service restart canceled for $serviceName" 'WARN'
        return
    }

    try {
        Restart-Service -Name $serviceName -Force
        Write-Log "Restarted service $serviceName"

        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Restart Service'
            Service = $serviceName
            Status  = 'Restarted'
        })
    }
    catch {
        Write-Log "Failed to restart service $serviceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time    = Get-Date
            Action  = 'Restart Service'
            Service = $serviceName
            Status  = 'Failed'
            Error   = $_.Exception.Message
        })
    }
}

function Get-InstalledPrinters {
    try {
        $printers = Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published
        Write-Log 'Retrieved installed printers'
        $printers | ForEach-Object { Add-ToolkitResult $_ }
    }
    catch {
        Write-Log "Failed to retrieve printers. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Installed Printers'
            Error  = $_.Exception.Message
        })
    }
}

function Test-ModuleAvailable {
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module is not installed or not available in this session."
        Write-Log "$ModuleName module not available" 'WARN'
        return $false
    }

    return $true
}

function Test-ADModuleInstalled {
    try {
        $module = Get-Module -ListAvailable -Name ActiveDirectory
        $rsat = Get-WindowsCapability -Online | Where-Object Name -like '*ActiveDirectory*'

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Check AD Module'
            ModuleFound  = [bool]$module
            ModulePath   = if ($module) { $module.Path } else { $null }
            RSATName     = if ($rsat) { $rsat.Name } else { $null }
            RSATState    = if ($rsat) { $rsat.State } else { 'Unknown' }
        })
    }
    catch {
        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Check AD Module'
            ModuleFound  = $false
            RSATState    = 'Unknown'
            Error        = $_.Exception.Message
        })
    }
}

function Install-ADModuleRSAT {
    if (-not (Confirm-Action 'Install RSAT Active Directory tools on this computer?')) {
        Write-Host 'Action canceled.'
        Write-Log 'RSAT Active Directory install canceled' 'WARN'
        return
    }

    try {
        Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
        Write-Log 'Installed RSAT Active Directory tools'

        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Install AD Module RSAT'
            Status = 'Install command completed'
        })
    }
    catch {
        Write-Log "Failed RSAT Active Directory install. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Install AD Module RSAT'
            Status = 'Failed'
            Error  = $_.Exception.Message
        })
    }
}

function Test-EntraModuleInstalled {
    try {
        $entra = Get-Module -ListAvailable -Name Microsoft.Entra*
        $graph = Get-Module -ListAvailable -Name Microsoft.Graph*

        Add-ToolkitResult ([PSCustomObject]@{
            Time         = Get-Date
            Action       = 'Check Entra Module'
            EntraFound   = [bool]$entra
            EntraModules = if ($entra) { ($entra.Name | Sort-Object -Unique) -join ', ' } else { $null }
            GraphFound   = [bool]$graph
            GraphModules = if ($graph) { ($graph.Name | Sort-Object -Unique | Select-Object -First 10) -join ', ' } else { $null }
        })
    }
    catch {
        Add-ToolkitResult ([PSCustomObject]@{
            Time       = Get-Date
            Action     = 'Check Entra Module'
            EntraFound = $false
            GraphFound = $false
            Error      = $_.Exception.Message
        })
    }
}

function Install-EntraModule {
    if (-not (Confirm-Action 'Install Microsoft.Entra module for the current user?')) {
        Write-Host 'Action canceled.'
        Write-Log 'Microsoft.Entra install canceled' 'WARN'
        return
    }

    try {
        Install-Module -Name Microsoft.Entra -Scope CurrentUser -Force -AllowClobber
        Write-Log 'Installed Microsoft.Entra module'

        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Install Entra Module'
            Status = 'Installed'
        })
    }
    catch {
        Write-Log "Failed Microsoft.Entra install. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Install Entra Module'
            Status = 'Failed'
            Error  = $_.Exception.Message
        })
    }
}

function Get-ADUserStatus {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    try {
        $user = Get-ADUser -Identity $username -Properties LockedOut, Enabled, LastLogonDate, PasswordExpired, PasswordLastSet, EmailAddress
        Write-Log "Retrieved AD user status for $username"

        Add-ToolkitResult ([PSCustomObject]@{
            Time            = Get-Date
            Action          = 'Get AD User Status'
            SamAccountName  = $user.SamAccountName
            Enabled         = $user.Enabled
            LockedOut       = $user.LockedOut
            LastLogonDate   = $user.LastLogonDate
            PasswordExpired = $user.PasswordExpired
            PasswordLastSet = $user.PasswordLastSet
            EmailAddress    = $user.EmailAddress
        })
    }
    catch {
        Write-Log "Failed AD user status lookup for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Get AD User Status'
            UserName = $username
            Error    = $_.Exception.Message
        })
    }
}

function Unlock-ADUserAccount {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username to unlock'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Unlock AD account $username?")) {
        Write-Host 'Action canceled.'
        Write-Log "Unlock AD account canceled for $username" 'WARN'
        return
    }

    try {
        Unlock-ADAccount -Identity $username
        Write-Log "Unlocked AD account for $username"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Unlock AD Account'
            UserName = $username
            Status   = 'Unlocked'
        })
    }
    catch {
        Write-Log "Failed to unlock AD account for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Unlock AD Account'
            UserName = $username
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Reset-ADUserPassword {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    $newPassword = Read-Host 'Enter new password' -AsSecureString

    if (-not (Confirm-Action "Reset password for $username and require password change at next logon?")) {
        Write-Host 'Action canceled.'
        Write-Log "Password reset canceled for $username" 'WARN'
        return
    }

    try {
        Set-ADAccountPassword -Identity $username -NewPassword $newPassword -Reset
        Set-ADUser -Identity $username -ChangePasswordAtLogon $true
        Unlock-ADAccount -Identity $username -ErrorAction SilentlyContinue
        Write-Log "Password reset completed for $username"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Reset AD Password'
            UserName = $username
            Status   = 'Password reset and change at logon enabled'
        })
    }
    catch {
        Write-Log "Failed password reset for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Reset AD Password'
            UserName = $username
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Get-ADUserGroups {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $username | Sort-Object Name | Select-Object Name
        Write-Log "Retrieved AD group membership for $username"

        if (-not $groups) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time     = Get-Date
                Action   = 'Get AD User Groups'
                UserName = $username
                Status   = 'No groups returned'
            })
        }
        else {
            $groups | ForEach-Object {
                Add-ToolkitResult ([PSCustomObject]@{
                    Time      = Get-Date
                    Action    = 'Get AD User Groups'
                    UserName  = $username
                    GroupName = $_.Name
                })
            }
        }
    }
    catch {
        Write-Log "Failed AD group lookup for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Get AD User Groups'
            UserName = $username
            Error    = $_.Exception.Message
        })
    }
}

function Add-ADUserToGroup {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username'
    $group = Read-Host 'Enter AD group name'

    if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($group)) {
        Write-Host 'Username and group cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Add $username to AD group $group?")) {
        Write-Host 'Action canceled.'
        Write-Log "Add AD group membership canceled for $username -> $group" 'WARN'
        return
    }

    try {
        Add-ADGroupMember -Identity $group -Members $username
        Write-Log "Added $username to AD group $group"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Add AD User To Group'
            UserName = $username
            Group    = $group
            Status   = 'Added'
        })
    }
    catch {
        Write-Log "Failed to add $username to AD group $group. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Add AD User To Group'
            UserName = $username
            Group    = $group
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Remove-ADUserFromGroup {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username'
    $group = Read-Host 'Enter AD group name'

    if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($group)) {
        Write-Host 'Username and group cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Remove $username from AD group $group?")) {
        Write-Host 'Action canceled.'
        Write-Log "Remove AD group membership canceled for $username -> $group" 'WARN'
        return
    }

    try {
        Remove-ADGroupMember -Identity $group -Members $username -Confirm:$false
        Write-Log "Removed $username from AD group $group"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Remove AD User From Group'
            UserName = $username
            Group    = $group
            Status   = 'Removed'
        })
    }
    catch {
        Write-Log "Failed to remove $username from AD group $group. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Remove AD User From Group'
            UserName = $username
            Group    = $group
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Disable-ADUserAccount {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username to disable'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Disable AD account $username?")) {
        Write-Host 'Action canceled.'
        Write-Log "Disable AD account canceled for $username" 'WARN'
        return
    }

    try {
        Disable-ADAccount -Identity $username
        Write-Log "Disabled AD account for $username"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Disable AD Account'
            UserName = $username
            Status   = 'Disabled'
        })
    }
    catch {
        Write-Log "Failed to disable AD account for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Disable AD Account'
            UserName = $username
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Enable-ADUserAccount {
    if (-not (Test-ModuleAvailable -ModuleName 'ActiveDirectory')) { return }

    Import-Module ActiveDirectory
    $username = Read-Host 'Enter username to enable'

    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host 'Username cannot be blank.'
        return
    }

    if (-not (Confirm-Action "Enable AD account $username?")) {
        Write-Host 'Action canceled.'
        Write-Log "Enable AD account canceled for $username" 'WARN'
        return
    }

    try {
        Enable-ADAccount -Identity $username
        Write-Log "Enabled AD account for $username"

        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Enable AD Account'
            UserName = $username
            Status   = 'Enabled'
        })
    }
    catch {
        Write-Log "Failed to enable AD account for $username. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Enable AD Account'
            UserName = $username
            Status   = 'Failed'
            Error    = $_.Exception.Message
        })
    }
}

function Connect-M365Graph {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Graph')) { return }

    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $context) {
            Connect-MgGraph -Scopes 'User.Read.All','Group.Read.All','Directory.Read.All','Device.Read.All'
            Write-Log 'Connected to Microsoft Graph'
        }
        else {
            Write-Log 'Microsoft Graph session already available'
        }

        $newContext = Get-MgContext
        Add-ToolkitResult ([PSCustomObject]@{
            Time        = Get-Date
            Action      = 'Connect Microsoft Graph'
            TenantId    = $newContext.TenantId
            Account     = $newContext.Account
            Environment = $newContext.Environment
            Status      = 'Connected'
        })
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Connect Microsoft Graph'
            Status = 'Failed'
            Error  = $_.Exception.Message
        })
    }
}

function Connect-EntraSession {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.Authentication')) { return }

    try {
        Connect-Entra
        $context = Get-EntraContext
        Write-Log 'Connected to Microsoft Entra'

        Add-ToolkitResult ([PSCustomObject]@{
            Time        = Get-Date
            Action      = 'Connect Entra'
            TenantId    = $context.TenantId
            Account     = $context.Account
            Environment = $context.Environment
            Status      = 'Connected'
        })
    }
    catch {
        Write-Log "Failed to connect to Entra. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Connect Entra'
            Status = 'Failed'
            Error  = $_.Exception.Message
        })
    }
}

function Get-EntraUserLookup {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.Users')) { return }

    $userId = Read-Host 'Enter user email / UPN'
    if ([string]::IsNullOrWhiteSpace($userId)) {
        Write-Host 'User value cannot be blank.'
        return
    }

    try {
        $user = Get-EntraUser -UserId $userId
        Write-Log "Retrieved Entra user $userId"

        Add-ToolkitResult ([PSCustomObject]@{
            Time              = Get-Date
            Action            = 'Get Entra User'
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
            AccountEnabled    = $user.AccountEnabled
            Id                = $user.Id
        })
    }
    catch {
        Write-Log "Failed Entra user lookup for $userId. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Entra User'
            User   = $userId
            Error  = $_.Exception.Message
        })
    }
}

function Get-EntraDeviceLookup {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.DirectoryManagement')) { return }

    $deviceName = Read-Host 'Enter device name'
    if ([string]::IsNullOrWhiteSpace($deviceName)) {
        Write-Host 'Device name cannot be blank.'
        return
    }

    try {
        $device = Get-EntraDevice -Filter "displayName eq '$deviceName'"
        if (-not $device) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time       = Get-Date
                Action     = 'Get Entra Device'
                DeviceName = $deviceName
                Status     = 'Not found'
            })
            return
        }

        foreach ($item in @($device)) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time           = Get-Date
                Action         = 'Get Entra Device'
                DeviceName     = $item.DisplayName
                DeviceId       = $item.DeviceId
                Id             = $item.Id
                OS             = $item.OperatingSystem
                OSVersion      = $item.OperatingSystemVersion
                TrustType      = $item.TrustType
                AccountEnabled = $item.AccountEnabled
            })
        }

        Write-Log "Retrieved Entra device $deviceName"
    }
    catch {
        Write-Log "Failed Entra device lookup for $deviceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Entra Device'
            Device = $deviceName
            Error  = $_.Exception.Message
        })
    }
}

function Get-EntraDeviceOwners {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.DirectoryManagement')) { return }

    $deviceId = Read-Host 'Enter Entra object Id for the device'
    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        Write-Host 'Device Id cannot be blank.'
        return
    }

    try {
        $owners = Get-EntraDeviceRegisteredOwner -DeviceId $deviceId
        if (-not $owners) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time     = Get-Date
                Action   = 'Get Entra Device Owners'
                DeviceId = $deviceId
                Status   = 'No owners found'
            })
            return
        }

        $owners | ForEach-Object {
            Add-ToolkitResult ([PSCustomObject]@{
                Time      = Get-Date
                Action    = 'Get Entra Device Owners'
                DeviceId  = $deviceId
                OwnerName = $_.DisplayName
                OwnerId   = $_.Id
            })
        }

        Write-Log "Retrieved Entra device owners for $deviceId"
    }
    catch {
        Write-Log "Failed Entra device owner lookup for $deviceId. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time     = Get-Date
            Action   = 'Get Entra Device Owners'
            DeviceId = $deviceId
            Error    = $_.Exception.Message
        })
    }
}

function Get-EntraUserDevices {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.Users')) { return }

    $userId = Read-Host 'Enter user email / UPN'
    if ([string]::IsNullOrWhiteSpace($userId)) {
        Write-Host 'User value cannot be blank.'
        return
    }

    try {
        $devices = Get-EntraUserRegisteredDevice -UserId $userId
        if (-not $devices) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time   = Get-Date
                Action = 'Get Entra User Devices'
                User   = $userId
                Status = 'No devices found'
            })
            return
        }

        $devices | ForEach-Object {
            Add-ToolkitResult ([PSCustomObject]@{
                Time      = Get-Date
                Action    = 'Get Entra User Devices'
                User      = $userId
                DeviceId  = $_.Id
                DeviceRef = $_.AdditionalProperties.displayName
            })
        }

        Write-Log "Retrieved Entra devices for $userId"
    }
    catch {
        Write-Log "Failed Entra user devices lookup for $userId. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Entra User Devices'
            User   = $userId
            Error  = $_.Exception.Message
        })
    }
}

function Search-EntraDeviceByName {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Entra.DirectoryManagement')) { return }

    $deviceName = Read-Host 'Enter full or partial device name'
    if ([string]::IsNullOrWhiteSpace($deviceName)) {
        Write-Host 'Device name cannot be blank.'
        return
    }

    try {
        $devices = Get-EntraDevice -All | Where-Object {
            $_.DisplayName -like "*$deviceName*"
        }

        if (-not $devices) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time       = Get-Date
                Action     = 'Search Entra Device'
                DeviceName = $deviceName
                Status     = 'Not found'
            })
            return
        }

        foreach ($item in $devices) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time            = Get-Date
                Action          = 'Search Entra Device'
                DeviceName      = $item.DisplayName
                DeviceId        = $item.DeviceId
                Id              = $item.Id
                OperatingSystem = $item.OperatingSystem
                OSVersion       = $item.OperatingSystemVersion
                TrustType       = $item.TrustType
                AccountEnabled  = $item.AccountEnabled
            })
        }

        Write-Log "Searched Entra devices for pattern $deviceName"
    }
    catch {
        Write-Log "Failed Entra device search for $deviceName. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time       = Get-Date
            Action     = 'Search Entra Device'
            DeviceName = $deviceName
            Error      = $_.Exception.Message
        })
    }
}

function Get-M365UserInfo {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Graph')) { return }

    $upn = Read-Host 'Enter user email / UPN'

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host 'UPN cannot be blank.'
        return
    }

    try {
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All'
        }

        $user = Get-MgUser -UserId $upn -Property DisplayName,UserPrincipalName,AccountEnabled,Mail,Department,JobTitle
        Write-Log "Retrieved Microsoft 365 user info for $upn"

        Add-ToolkitResult ([PSCustomObject]@{
            Time              = Get-Date
            Action            = 'Get Microsoft 365 User Info'
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
            Department        = $user.Department
            JobTitle          = $user.JobTitle
            AccountEnabled    = $user.AccountEnabled
        })
    }
    catch {
        Write-Log "Failed Microsoft 365 user lookup for $upn. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Microsoft 365 User Info'
            User   = $upn
            Error  = $_.Exception.Message
        })
    }
}

function Get-M365UserLicenses {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Graph')) { return }

    $upn = Read-Host 'Enter user email / UPN'

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host 'UPN cannot be blank.'
        return
    }

    try {
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -Scopes 'User.Read.All','Directory.Read.All'
        }

        $licenses = Get-MgUserLicenseDetail -UserId $upn
        Write-Log "Retrieved Microsoft 365 licenses for $upn"

        if (-not $licenses) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time   = Get-Date
                Action = 'Get Microsoft 365 User Licenses'
                User   = $upn
                Status = 'No licenses found'
            })
        }
        else {
            $licenses | ForEach-Object {
                Add-ToolkitResult ([PSCustomObject]@{
                    Time          = Get-Date
                    Action        = 'Get Microsoft 365 User Licenses'
                    User          = $upn
                    SkuPartNumber = $_.SkuPartNumber
                    ServicePlans  = ($_.ServicePlans.ServicePlanName -join ', ')
                })
            }
        }
    }
    catch {
        Write-Log "Failed Microsoft 365 license lookup for $upn. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Microsoft 365 User Licenses'
            User   = $upn
            Error  = $_.Exception.Message
        })
    }
}

function Get-M365UserGroups {
    if (-not (Test-ModuleAvailable -ModuleName 'Microsoft.Graph')) { return }

    $upn = Read-Host 'Enter user email / UPN'

    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Host 'UPN cannot be blank.'
        return
    }

    try {
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -Scopes 'User.Read.All','Group.Read.All','Directory.Read.All'
        }

        $groups = Get-MgUserMemberOf -UserId $upn -All
        Write-Log "Retrieved Microsoft 365 groups for $upn"

        if (-not $groups) {
            Add-ToolkitResult ([PSCustomObject]@{
                Time   = Get-Date
                Action = 'Get Microsoft 365 User Groups'
                User   = $upn
                Status = 'No groups found'
            })
        }
        else {
            $groups | ForEach-Object {
                $name = $_.AdditionalProperties.displayName
                Add-ToolkitResult ([PSCustomObject]@{
                    Time      = Get-Date
                    Action    = 'Get Microsoft 365 User Groups'
                    User      = $upn
                    GroupName = if ($name) { $name } else { 'Unknown' }
                })
            }
        }
    }
    catch {
        Write-Log "Failed Microsoft 365 group lookup for $upn. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Get Microsoft 365 User Groups'
            User   = $upn
            Error  = $_.Exception.Message
        })
    }
}

function Get-HybridUserReport {
    $userId = Read-Host 'Enter username or email / UPN'

    if ([string]::IsNullOrWhiteSpace($userId)) {
        Write-Host 'User value cannot be blank.'
        return
    }

    $adSam = $userId
    $entraUpn = $userId

    try {
        $adUser = $null
        $entraUser = $null
        $licenses = $null
        $adGroups = $null
        $entraGroups = $null

        if (Get-Module -ListAvailable -Name ActiveDirectory) {
            Import-Module ActiveDirectory
            $adUser = Get-ADUser -Identity $adSam -Properties DisplayName, Enabled, LockedOut, LastLogonDate, EmailAddress, Department, Title -ErrorAction SilentlyContinue
            if ($adUser) {
                $adGroups = Get-ADPrincipalGroupMembership -Identity $adUser.SamAccountName |
                    Sort-Object Name |
                    Select-Object -ExpandProperty Name
            }
        }

        if (Get-Module -ListAvailable -Name Microsoft.Graph) {
            if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
                Connect-MgGraph -Scopes 'User.Read.All','Group.Read.All','Directory.Read.All'
            }

            $entraUser = Get-MgUser -UserId $entraUpn -Property DisplayName,UserPrincipalName,Mail,AccountEnabled,Department,JobTitle -ErrorAction SilentlyContinue
            if ($entraUser) {
                $licenses = Get-MgUserLicenseDetail -UserId $entraUpn -ErrorAction SilentlyContinue
                $entraGroups = Get-MgUserMemberOf -UserId $entraUpn -All -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.AdditionalProperties.displayName } |
                    Where-Object { $_ }
            }
        }

        Add-ToolkitResult ([PSCustomObject]@{
            Time               = Get-Date
            Action             = 'Hybrid User Report'
            Input              = $userId
            AD_DisplayName     = $adUser.DisplayName
            AD_SamAccountName  = $adUser.SamAccountName
            AD_Enabled         = $adUser.Enabled
            AD_LockedOut       = $adUser.LockedOut
            AD_LastLogonDate   = $adUser.LastLogonDate
            AD_Email           = $adUser.EmailAddress
            AD_Department      = $adUser.Department
            AD_Title           = $adUser.Title
            AD_Groups          = if ($adGroups) { $adGroups -join '; ' } else { $null }
            Entra_DisplayName  = $entraUser.DisplayName
            Entra_UPN          = $entraUser.UserPrincipalName
            Entra_Mail         = $entraUser.Mail
            Entra_Enabled      = $entraUser.AccountEnabled
            Entra_Department   = $entraUser.Department
            Entra_JobTitle     = $entraUser.JobTitle
            Entra_Licenses     = if ($licenses) { ($licenses.SkuPartNumber -join '; ') } else { $null }
            Entra_Groups       = if ($entraGroups) { $entraGroups -join '; ' } else { $null }
        })

        Write-Log "Generated hybrid user report for $userId"
    }
    catch {
        Write-Log "Failed hybrid user report for $userId. $_" 'ERROR'
        Add-ToolkitResult ([PSCustomObject]@{
            Time   = Get-Date
            Action = 'Hybrid User Report'
            Input  = $userId
            Error  = $_.Exception.Message
        })
    }
}

function Export-LastResults {
    if (-not $Global:LastResults -or $Global:LastResults.Count -eq 0) {
        Write-Host 'No results to export yet.'
        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $exportPath = Join-Path $ExportFolder "HelpDeskResults_$timestamp.csv"

    try {
        $Global:LastResults | Export-Csv -Path $exportPath -NoTypeInformation -Force
        Write-Log "Exported results to $exportPath"
        Write-Host "Results exported to $exportPath"
    }
    catch {
        Write-Log "Failed to export results. $_" 'ERROR'
        Write-Host "Export failed: $($_.Exception.Message)"
    }
}

function Clear-LastResults {
    $Global:LastResults = @()
    Write-Log 'Cleared in-memory results'
    Write-Host 'Saved results have been cleared from memory.'
}

function Show-LastResults {
    if (-not $Global:LastResults -or $Global:LastResults.Count -eq 0) {
        Write-Host 'No results collected yet.'
        return
    }

    $Global:LastResults | Format-Table -AutoSize
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
        Write-Host '--- Printer Tools ---'
        Write-Host '1. Restart spooler on remote PC'
        Write-Host '2. Clear print queue on remote PC'
        Write-Host '3. List local installed printers'
        Write-Host '4. Back'

        $choice = Read-Host 'Choose an option'

        switch ($choice) {
            '1' { $c = Get-TargetComputer; Restart-PrintSpooler -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '2' { $c = Get-TargetComputer; Clear-PrintQueue -ComputerName $c | Format-Table -AutoSize; Pause-Toolkit }
            '3' { Get-InstalledPrinters | Format-Table -AutoSize; Pause-Toolkit }
            '4' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '4')
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

function Show-SoftwareMenu {
    do {
        Clear-Host
        Write-Host '--- Software and Process Tools ---'
        Write-Host '1. Find process by name'
        Write-Host '2. Stop process by name'
        Write-Host '3. Show top memory processes'
        Write-Host '4. Get local service status'
        Write-Host '5. Restart local service'
        Write-Host '6. Back'

        $choice = Read-Host 'Choose an option'

        switch ($choice) {
            '1' { Find-ProcessByName | Format-Table -AutoSize; Pause-Toolkit }
            '2' { Stop-ProcessByName | Format-Table -AutoSize; Pause-Toolkit }
            '3' { Get-TopMemoryProcesses | Format-Table -AutoSize; Pause-Toolkit }
            '4' { Get-ServiceStatus | Format-Table -AutoSize; Pause-Toolkit }
            '5' { Restart-LocalServiceByName | Format-Table -AutoSize; Pause-Toolkit }
            '6' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '6')
}

function Show-EventLogMenu {
    do {
        Clear-Host
        Write-Host '--- Event Log Tools ---'
        Write-Host '1. Show recent system errors'
        Write-Host '2. Show recent application errors'
        Write-Host '3. Back'

        $choice = Read-Host 'Choose an option'

        switch ($choice) {
            '1' { Get-RecentSystemErrors | Format-Table -Wrap; Pause-Toolkit }
            '2' { Get-RecentApplicationErrors | Format-Table -Wrap; Pause-Toolkit }
            '3' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '3')
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

function Show-ADMenu {
    do {
        Clear-Host
        Write-Host '--- Active Directory Tools ---'
        Write-Host '1. Check AD module / RSAT status'
        Write-Host '2. Install RSAT Active Directory tools'
        Write-Host '3. Get AD user status'
        Write-Host '4. Unlock AD account'
        Write-Host '5. Reset AD password'
        Write-Host '6. Get AD user groups'
        Write-Host '7. Add AD user to group'
        Write-Host '8. Remove AD user from group'
        Write-Host '9. Disable AD account'
        Write-Host '10. Enable AD account'
        Write-Host '11. Back'

        $choice = Read-Host 'Choose an option'

        switch ($choice) {
            '1'  { Test-ADModuleInstalled | Format-Table -AutoSize; Pause-Toolkit }
            '2'  { Install-ADModuleRSAT | Format-Table -AutoSize; Pause-Toolkit }
            '3'  { Get-ADUserStatus | Format-Table -AutoSize; Pause-Toolkit }
            '4'  { Unlock-ADUserAccount | Format-Table -AutoSize; Pause-Toolkit }
            '5'  { Reset-ADUserPassword | Format-Table -AutoSize; Pause-Toolkit }
            '6'  { Get-ADUserGroups | Format-Table -AutoSize; Pause-Toolkit }
            '7'  { Add-ADUserToGroup | Format-Table -AutoSize; Pause-Toolkit }
            '8'  { Remove-ADUserFromGroup | Format-Table -AutoSize; Pause-Toolkit }
            '9'  { Disable-ADUserAccount | Format-Table -AutoSize; Pause-Toolkit }
            '10' { Enable-ADUserAccount | Format-Table -AutoSize; Pause-Toolkit }
            '11' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '11')
}

function Show-M365Menu {
    do {
        Clear-Host
        Write-Host '--- Microsoft 365 / Entra / Graph Tools ---'
        Write-Host '1. Check Entra / Graph modules'
        Write-Host '2. Install Microsoft.Entra module'
        Write-Host '3. Connect to Microsoft Entra'
        Write-Host '4. Connect to Microsoft Graph'
        Write-Host '5. Get Entra user lookup'
        Write-Host '6. Get Entra device lookup'
        Write-Host '7. Search Entra device by partial name'
        Write-Host '8. Get Entra device owners'
        Write-Host '9. Get Entra user devices'
        Write-Host '10. Get Microsoft 365 user info'
        Write-Host '11. Get Microsoft 365 user licenses'
        Write-Host '12. Get Microsoft 365 user groups'
        Write-Host '13. Run hybrid AD + Entra user report'
        Write-Host '14. Back'

        $choice = Read-Host 'Choose an option'

        switch ($choice) {
            '1'  { Test-EntraModuleInstalled | Format-Table -AutoSize; Pause-Toolkit }
            '2'  { Install-EntraModule | Format-Table -AutoSize; Pause-Toolkit }
            '3'  { Connect-EntraSession | Format-Table -AutoSize; Pause-Toolkit }
            '4'  { Connect-M365Graph | Format-Table -AutoSize; Pause-Toolkit }
            '5'  { Get-EntraUserLookup | Format-Table -AutoSize; Pause-Toolkit }
            '6'  { Get-EntraDeviceLookup | Format-Table -AutoSize; Pause-Toolkit }
            '7'  { Search-EntraDeviceByName | Format-Table -AutoSize; Pause-Toolkit }
            '8'  { Get-EntraDeviceOwners | Format-Table -AutoSize; Pause-Toolkit }
            '9'  { Get-EntraUserDevices | Format-Table -AutoSize; Pause-Toolkit }
            '10' { Get-M365UserInfo | Format-Table -AutoSize; Pause-Toolkit }
            '11' { Get-M365UserLicenses | Format-Table -AutoSize; Pause-Toolkit }
            '12' { Get-M365UserGroups | Format-Table -AutoSize; Pause-Toolkit }
            '13' { Get-HybridUserReport | Format-Table -Wrap; Pause-Toolkit }
            '14' { }
            default { Write-Host 'Invalid choice.'; Pause-Toolkit }
        }
    } until ($choice -eq '14')
}

try {
    Start-Transcript -Path (Join-Path $ExportFolder 'ToolkitTranscript.txt') -Append | Out-Null
}
catch {
    Write-Log "Could not start transcript. $_" 'WARN'
}

Write-Log 'Toolkit started'

try {
    do {
        Show-Header
        Write-Host '1. Computer Tools'
        Write-Host '2. Printer Tools'
        Write-Host '3. Network Tools'
        Write-Host '4. Software and Process Tools'
        Write-Host '5. Event Log Tools'
        Write-Host '6. Active Directory Tools'
        Write-Host '7. Microsoft 365 / Graph Tools'
        Write-Host '8. Reporting Tools'
        Write-Host '9. Exit'
        Write-Host ''

        $mainChoice = Read-Host 'Choose an option'

        switch ($mainChoice) {
            '1' { Show-ComputerMenu }
            '2' { Show-PrinterMenu }
            '3' { Show-NetworkMenu }
            '4' { Show-SoftwareMenu }
            '5' { Show-EventLogMenu }
            '6' { Show-ADMenu }
            '7' { Show-M365Menu }
            '8' { Show-ReportingMenu }
            '9' { Write-Host 'Exiting toolkit...' }
            default {
                Write-Host 'Invalid choice.'
                Pause-Toolkit
            }
        }
    } until ($mainChoice -eq '9')
}
finally {
    Write-Log 'Toolkit exited'
    try {
        Stop-Transcript | Out-Null
    }
    catch {
    }
}
