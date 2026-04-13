#Requires -RunAsAdministrator

<#
.SYNOPSIS
Checks installed Zebra printers, verifies the driver is approved,
and replaces it with an approved driver if needed.

.NOTES
Place this script in the same folder as ZBRN.inf and run as Administrator.
Approved drivers:
- ZDesigner ZD421-300dpi ZPL
- ZDesigner ZT411-203dpi ZPL
- ZDesigner ZT411-300dpi ZPL
#>

$ApprovedDrivers = @(
    "ZDesigner ZD421-300dpi ZPL",
    "ZDesigner ZT411-203dpi ZPL",
    "ZDesigner ZT411-300dpi ZPL"
)

# Pick the driver you want to install if replacement is needed
$PreferredDriver = "ZDesigner ZT411-203dpi ZPL"

$InfPath = Join-Path $PSScriptRoot "ZBRN.inf"
$LogFile = Join-Path $env:TEMP ("ZebraDriverCheck_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line

    switch ($Level) {
        "INFO"  { Write-Host $line -ForegroundColor Cyan }
        "OK"    { Write-Host $line -ForegroundColor Green }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
    }
}

function Remove-ZebraDriver {
    param(
        [Parameter(Mandatory)]
        [string]$DriverName
    )

    Write-Log "Removing driver: $DriverName" "WARN"

    $linkedPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object {
        $_.DriverName -eq $DriverName
    }

    foreach ($printer in $linkedPrinters) {
        try {
            Remove-Printer -Name $printer.Name -ErrorAction Stop
            Write-Log "Removed printer queue: $($printer.Name)" "OK"
        }
        catch {
            Write-Log "Failed to remove printer queue: $($printer.Name). $($_.Exception.Message)" "ERROR"
        }
    }

    try {
        Remove-PrinterDriver -Name $DriverName -ErrorAction Stop
        Write-Log "Removed printer driver: $DriverName" "OK"
        return
    }
    catch {
        Write-Log "Standard driver removal failed for $DriverName. Trying fallback." "WARN"
    }

    try {
        Start-Process -FilePath "rundll32.exe" `
            -ArgumentList "printui.dll,PrintUIEntry /dd /m `"$DriverName`"" `
            -Wait -NoNewWindow
        Write-Log "Fallback removal attempted for: $DriverName" "OK"
    }
    catch {
        Write-Log "Fallback removal failed for $DriverName. $($_.Exception.Message)" "ERROR"
    }
}

function Install-ApprovedDriver {
    param(
        [Parameter(Mandatory)]
        [string]$DriverName,

        [Parameter(Mandatory)]
        [string]$InfPath
    )

    if (-not (Test-Path $InfPath)) {
        throw "INF not found: $InfPath"
    }

    Write-Log "Using INF: $InfPath" "INFO"
    Write-Log "Staging driver package with pnputil." "INFO"

    $pnpOutput = & pnputil.exe /add-driver $InfPath /install 2>&1
    $pnpOutput | ForEach-Object { Write-Log $_ "INFO" }

    $existing = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "Approved driver already installed: $DriverName" "OK"
        return
    }

    try {
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
        Write-Log "Installed approved driver: $DriverName" "OK"
    }
    catch {
        throw "Failed to install approved driver $DriverName. $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Zebra Driver Check and Fix" -ForegroundColor White
Write-Host "Log file: $LogFile" -ForegroundColor DarkGray
Write-Host ""

Write-Log "Script started on $env:COMPUTERNAME by $env:USERNAME"

$zebraPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like "*Zebra*" -or
    $_.Name -like "*ZDesigner*" -or
    $_.DriverName -like "*ZDesigner*"
}

$zebraDrivers = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like "*ZDesigner*"
}

if (-not $zebraPrinters) {
    Write-Log "No Zebra printers found on this machine." "WARN"
    Write-Log "No changes made." "INFO"
    Write-Log "Done." "OK"
    exit 0
}

Write-Log "Found $(@($zebraPrinters).Count) Zebra printer(s)." "INFO"

$NeedsFix = $false
$CurrentDrivers = @()

foreach ($printer in $zebraPrinters) {
    $driverName = $printer.DriverName
    $CurrentDrivers += $driverName

    Write-Log "Printer found: $($printer.Name) | Driver: $driverName | Port: $($printer.PortName)" "INFO"

    if ($ApprovedDrivers -contains $driverName) {
        Write-Log "Nicee - approved driver detected on '$($printer.Name)': $driverName" "OK"
    }
    else {
        Write-Log "Unapproved driver detected on '$($printer.Name)': $driverName" "WARN"
        $NeedsFix = $true
    }
}

$CurrentDrivers = $CurrentDrivers | Sort-Object -Unique

if (-not $NeedsFix) {
    Write-Log "All Zebra printers are already using approved drivers. Done." "OK"
    exit 0
}

foreach ($driver in $CurrentDrivers) {
    if ($ApprovedDrivers -notcontains $driver) {
        Remove-ZebraDriver -DriverName $driver
    }
}

try {
    Install-ApprovedDriver -DriverName $PreferredDriver -InfPath $InfPath
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    exit 1
}

Write-Log "Driver replacement complete." "OK"
Write-Log "Approved driver ready: $PreferredDriver" "OK"
Write-Log "Done." "OK"
