Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Help Desk Toolkit GUI v1'
$form.Size = New-Object System.Drawing.Size(1180,760)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1180,760)

$font = New-Object System.Drawing.Font('Segoe UI',10)
$form.Font = $font

$lblComputer = New-Object System.Windows.Forms.Label
$lblComputer.Text = 'Computer Name:'
$lblComputer.Location = New-Object System.Drawing.Point(20,20)
$lblComputer.Size = New-Object System.Drawing.Size(120,24)
$form.Controls.Add($lblComputer)

$txtComputer = New-Object System.Windows.Forms.TextBox
$txtComputer.Location = New-Object System.Drawing.Point(145,18)
$txtComputer.Size = New-Object System.Drawing.Size(220,28)
$form.Controls.Add($txtComputer)

$lblSoftware = New-Object System.Windows.Forms.Label
$lblSoftware.Text = 'Software Search:'
$lblSoftware.Location = New-Object System.Drawing.Point(390,20)
$lblSoftware.Size = New-Object System.Drawing.Size(120,24)
$form.Controls.Add($lblSoftware)

$txtSoftware = New-Object System.Windows.Forms.TextBox
$txtSoftware.Location = New-Object System.Drawing.Point(515,18)
$txtSoftware.Size = New-Object System.Drawing.Size(220,28)
$form.Controls.Add($txtSoftware)

$lblService = New-Object System.Windows.Forms.Label
$lblService.Text = 'Service Name:'
$lblService.Location = New-Object System.Drawing.Point(760,20)
$lblService.Size = New-Object System.Drawing.Size(100,24)
$form.Controls.Add($lblService)

$txtService = New-Object System.Windows.Forms.TextBox
$txtService.Location = New-Object System.Drawing.Point(860,18)
$txtService.Size = New-Object System.Drawing.Size(180,28)
$form.Controls.Add($txtService)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(260,60)
$outputBox.Size = New-Object System.Drawing.Size(890,620)
$outputBox.Multiline = $true
$outputBox.ScrollBars = 'Vertical'
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font('Consolas',10)
$form.Controls.Add($outputBox)

function Add-Output {
    param([string]$Text)
    $outputBox.AppendText($Text + [Environment]::NewLine)
}

function Add-Separator {
    $outputBox.AppendText(('=' * 80) + [Environment]::NewLine)
}

function Require-ComputerName {
    if ([string]::IsNullOrWhiteSpace($txtComputer.Text)) {
        Add-Output 'Enter a computer name first.'
        return $false
    }
    return $true
}

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(20,60)
$panel.Size = New-Object System.Drawing.Size(220,620)
$panel.AutoScroll = $true
$form.Controls.Add($panel)

$buttonY = 0
function Add-ToolkitButton {
    param(
        [string]$Text,
        [scriptblock]$OnClick
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point(0,$script:buttonY)
    $button.Size = New-Object System.Drawing.Size(200,34)
    $button.Add_Click($OnClick)
    $panel.Controls.Add($button)
    $script:buttonY += 40
}

Add-ToolkitButton 'Ping Computer' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $online = Test-Connection -ComputerName $computer -Count 1 -Quiet
        Add-Output ("Ping result for {0}: {1}" -f $computer, $online)
        Add-Separator
    }
    catch {
        Add-Output ("Ping failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get Logged-In User' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $user = (Get-CimInstance Win32_ComputerSystem -ComputerName $computer).UserName
        Add-Output ("Logged-in user on {0}: {1}" -f $computer, $user)
        Add-Separator
    }
    catch {
        Add-Output ("Lookup failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get System Info' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $computer
        $bios = Get-CimInstance Win32_BIOS -ComputerName $computer
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $computer
        Add-Output ("Computer: {0}" -f $computer)
        Add-Output ("Manufacturer: {0}" -f $cs.Manufacturer)
        Add-Output ("Model: {0}" -f $cs.Model)
        Add-Output ("Serial: {0}" -f $bios.SerialNumber)
        Add-Output ("OS: {0}" -f $os.Caption)
        Add-Output ("Version: {0}" -f $os.Version)
        Add-Output ("Last Boot: {0}" -f $os.LastBootUpTime)
        Add-Separator
    }
    catch {
        Add-Output ("System info failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get Disk Space' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $computer -Filter "DriveType=3"
        Add-Output ("Disk space for {0}:" -f $computer)
        foreach ($disk in $disks) {
            Add-Output ("{0}  Free: {1} GB  Size: {2} GB" -f $disk.DeviceID, [math]::Round($disk.FreeSpace/1GB,2), [math]::Round($disk.Size/1GB,2))
        }
        Add-Separator
    }
    catch {
        Add-Output ("Disk space lookup failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Installed Software' {
    try {
        $paths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        $apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Sort-Object DisplayName |
            Select-Object -First 40 DisplayName, DisplayVersion, Publisher
        Add-Output 'Installed software (top 40):'
        foreach ($app in $apps) {
            Add-Output ("{0} | {1} | {2}" -f $app.DisplayName, $app.DisplayVersion, $app.Publisher)
        }
        Add-Separator
    }
    catch {
        Add-Output ("Software inventory failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Search Software' {
    $name = $txtSoftware.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        Add-Output 'Enter a software search term first.'
        Add-Separator
        return
    }
    try {
        $paths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        $apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like "*$name*" } |
            Sort-Object DisplayName |
            Select-Object DisplayName, DisplayVersion, Publisher
        Add-Output ("Software search for: {0}" -f $name)
        foreach ($app in $apps) {
            Add-Output ("{0} | {1} | {2}" -f $app.DisplayName, $app.DisplayVersion, $app.Publisher)
        }
        if (-not $apps) { Add-Output 'No matching software found.' }
        Add-Separator
    }
    catch {
        Add-Output ("Software search failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get Startup Programs' {
    try {
        $items = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location
        Add-Output 'Startup programs:'
        foreach ($item in $items) {
            Add-Output ("{0} | {1} | {2}" -f $item.Name, $item.Location, $item.Command)
        }
        Add-Separator
    }
    catch {
        Add-Output ("Startup check failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get Failed Services' {
    try {
        $services = Get-Service | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' }
        Add-Output 'Failed automatic services:'
        foreach ($svc in $services) {
            Add-Output ("{0} | {1}" -f $svc.Name, $svc.DisplayName)
        }
        if (-not $services) { Add-Output 'No failed automatic services found.' }
        Add-Separator
    }
    catch {
        Add-Output ("Failed services check failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Get Service Status' {
    $service = $txtService.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($service)) {
        Add-Output 'Enter a service name first.'
        Add-Separator
        return
    }
    try {
        $svc = Get-Service -Name $service | Select-Object Name, DisplayName, Status, StartType
        Add-Output ("Service: {0}" -f $svc.Name)
        Add-Output ("Display Name: {0}" -f $svc.DisplayName)
        Add-Output ("Status: {0}" -f $svc.Status)
        Add-Output ("Startup: {0}" -f $svc.StartType)
        Add-Separator
    }
    catch {
        Add-Output ("Service lookup failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'List Local Printers' {
    try {
        $printers = Get-Printer | Select-Object Name, DriverName, PortName
        Add-Output 'Local printers:'
        foreach ($printer in $printers) {
            Add-Output ("{0} | {1} | {2}" -f $printer.Name, $printer.DriverName, $printer.PortName)
        }
        Add-Separator
    }
    catch {
        Add-Output ("Printer lookup failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Remote Printers' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $printers = Invoke-Command -ComputerName $computer -ScriptBlock {
            Get-Printer | Select-Object Name, DriverName, PortName
        }
        Add-Output ("Remote printers on {0}:" -f $computer)
        foreach ($printer in $printers) {
            Add-Output ("{0} | {1} | {2}" -f $printer.Name, $printer.DriverName, $printer.PortName)
        }
        Add-Separator
    }
    catch {
        Add-Output ("Remote printer lookup failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Full Workstation Report' {
    if (-not (Require-ComputerName)) { return }
    $computer = $txtComputer.Text.Trim()
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $computer
        $bios = Get-CimInstance Win32_BIOS -ComputerName $computer
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $computer
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $computer -Filter "DriveType=3"
        $uptime = (Get-Date) - $os.LastBootUpTime
        Add-Output ("Workstation Report for {0}" -f $computer)
        Add-Output ("Manufacturer: {0}" -f $cs.Manufacturer)
        Add-Output ("Model: {0}" -f $cs.Model)
        Add-Output ("Serial: {0}" -f $bios.SerialNumber)
        Add-Output ("User: {0}" -f $cs.UserName)
        Add-Output ("OS: {0}" -f $os.Caption)
        Add-Output ("Version: {0}" -f $os.Version)
        Add-Output ("Last Boot: {0}" -f $os.LastBootUpTime)
        Add-Output ("Uptime Days: {0}" -f [math]::Round($uptime.TotalDays,2))
        foreach ($disk in $disks) {
            Add-Output ("Disk {0}  Free: {1} GB  Size: {2} GB" -f $disk.DeviceID, [math]::Round($disk.FreeSpace/1GB,2), [math]::Round($disk.Size/1GB,2))
        }
        Add-Separator
    }
    catch {
        Add-Output ("Workstation report failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Copy Output' {
    try {
        Set-Clipboard -Value $outputBox.Text
        Add-Output 'Output copied to clipboard.'
        Add-Separator
    }
    catch {
        Add-Output ("Copy failed: {0}" -f $_.Exception.Message)
        Add-Separator
    }
}

Add-ToolkitButton 'Clear Output' {
    $outputBox.Clear()
}

Add-ToolkitButton 'Exit' {
    $form.Close()
}

[void]$form.ShowDialog()
