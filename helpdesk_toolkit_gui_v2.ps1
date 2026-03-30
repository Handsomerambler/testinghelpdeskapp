Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-EmbeddedLogoImage {
    $base64 = @'
iVBORw0KGgoAAAANSUhEUgAABBoAAAIzCAYAAACur3qbAAAACXBIWXMAAC4jAAAuIwF4pT92AAAFFmlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4
'@
    try {
        $clean = ($base64 -replace '\s','')
        $bytes = [Convert]::FromBase64String($clean)
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        return [System.Drawing.Image]::FromStream($ms)
    }
    catch {
        return $null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'VSE Aviation Help Desk Toolkit GUI v2'
$form.Size = New-Object System.Drawing.Size(1280, 820)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1280, 820)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$script:BulkComputers = @()
$script:Credential = $null

function Set-Status {
    param(
        [string]$Message,
        [ValidateSet('Good','Bad','Neutral')]
        [string]$State = 'Neutral'
    )

    $lblStatus.Text = "Status: $Message"
    switch ($State) {
        'Good'    { $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen }
        'Bad'     { $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed }
        'Neutral' { $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue }
    }
}

function Get-EnteredComputer {
    return $txtComputer.Text.Trim()
}

function Require-Computer {
    $computer = Get-EnteredComputer
    if ([string]::IsNullOrWhiteSpace($computer)) {
        Set-Status -Message 'Enter a computer name first.' -State Bad
        return $false
    }
    return $true
}

function Invoke-RemoteAction {
    param(
        [string]$ComputerName,
        [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    if ($script:Credential) {
        Invoke-Command -ComputerName $ComputerName -Credential $script:Credential -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    }
    else {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    }
}

function Export-OutputToFile {
    param([string]$TextToSave,[string]$DefaultName='ToolkitOutput.txt')
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
    $saveDialog.Title = 'Save Output'
    $saveDialog.FileName = $DefaultName
    if ($saveDialog.ShowDialog() -eq 'OK') {
        $TextToSave | Out-File -FilePath $saveDialog.FileName -Encoding utf8
        Set-Status -Message "Output saved to $($saveDialog.FileName)" -State Good
    }
}

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(10, 10)
$headerPanel.Size = New-Object System.Drawing.Size(1240, 120)
$headerPanel.BorderStyle = 'FixedSingle'
$form.Controls.Add($headerPanel)

$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(10, 10)
$pictureBox.Size = New-Object System.Drawing.Size(260, 95)
$pictureBox.SizeMode = 'Zoom'
$headerPanel.Controls.Add($pictureBox)

$logoPath = Join-Path $PSScriptRoot 'VSE Aviation-Color.png'
if (Test-Path $logoPath) {
    try { $pictureBox.Image = [System.Drawing.Image]::FromFile($logoPath) } catch {}
}
if (-not $pictureBox.Image) {
    $pictureBox.Image = Get-EmbeddedLogoImage
}

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'VSE Aviation Help Desk Toolkit GUI v2'
$lblTitle.Location = New-Object System.Drawing.Point(290, 10)
$lblTitle.Size = New-Object System.Drawing.Size(520, 30)
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($lblTitle)

$lblComputer = New-Object System.Windows.Forms.Label
$lblComputer.Text = 'Computer Name:'
$lblComputer.Location = New-Object System.Drawing.Point(290, 50)
$lblComputer.Size = New-Object System.Drawing.Size(120, 25)
$headerPanel.Controls.Add($lblComputer)

$txtComputer = New-Object System.Windows.Forms.TextBox
$txtComputer.Location = New-Object System.Drawing.Point(420, 48)
$txtComputer.Size = New-Object System.Drawing.Size(220, 25)
$headerPanel.Controls.Add($txtComputer)

$lblSoftware = New-Object System.Windows.Forms.Label
$lblSoftware.Text = 'Software Search:'
$lblSoftware.Location = New-Object System.Drawing.Point(660, 50)
$lblSoftware.Size = New-Object System.Drawing.Size(120, 25)
$headerPanel.Controls.Add($lblSoftware)

$txtSoftware = New-Object System.Windows.Forms.TextBox
$txtSoftware.Location = New-Object System.Drawing.Point(785, 48)
$txtSoftware.Size = New-Object System.Drawing.Size(170, 25)
$headerPanel.Controls.Add($txtSoftware)

$lblService = New-Object System.Windows.Forms.Label
$lblService.Text = 'Service Name:'
$lblService.Location = New-Object System.Drawing.Point(970, 50)
$lblService.Size = New-Object System.Drawing.Size(90, 25)
$headerPanel.Controls.Add($lblService)

$txtService = New-Object System.Windows.Forms.TextBox
$txtService.Location = New-Object System.Drawing.Point(1060, 48)
$txtService.Size = New-Object System.Drawing.Size(150, 25)
$headerPanel.Controls.Add($txtService)

$btnCredential = New-Object System.Windows.Forms.Button
$btnCredential.Text = 'Set Admin Credential'
$btnCredential.Location = New-Object System.Drawing.Point(290, 82)
$btnCredential.Size = New-Object System.Drawing.Size(160, 28)
$btnCredential.Add_Click({
    try {
        $script:Credential = Get-Credential
        Set-Status -Message 'Credential loaded.' -State Good
    }
    catch {
        Set-Status -Message 'Credential prompt canceled.' -State Neutral
    }
})
$headerPanel.Controls.Add($btnCredential)

$btnBulkLoad = New-Object System.Windows.Forms.Button
$btnBulkLoad.Text = 'Load Bulk List'
$btnBulkLoad.Location = New-Object System.Drawing.Point(465, 82)
$btnBulkLoad.Size = New-Object System.Drawing.Size(120, 28)
$btnBulkLoad.Add_Click({
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = 'Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*'
    if ($openDialog.ShowDialog() -eq 'OK') {
        $script:BulkComputers = Get-Content $openDialog.FileName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        Set-Status -Message "$($script:BulkComputers.Count) bulk computers loaded." -State Good
    }
})
$headerPanel.Controls.Add($btnBulkLoad)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Status: Ready'
$lblStatus.Location = New-Object System.Drawing.Point(610, 84)
$lblStatus.Size = New-Object System.Drawing.Size(600, 24)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$headerPanel.Controls.Add($lblStatus)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 140)
$tabs.Size = New-Object System.Drawing.Size(1240, 620)
$form.Controls.Add($tabs)

$tabComputer  = New-Object System.Windows.Forms.TabPage
$tabPrinter   = New-Object System.Windows.Forms.TabPage
$tabSoftware  = New-Object System.Windows.Forms.TabPage
$tabReporting = New-Object System.Windows.Forms.TabPage
$tabComputer.Text  = 'Computer'
$tabPrinter.Text   = 'Printer'
$tabSoftware.Text  = 'Software'
$tabReporting.Text = 'Reporting'
$tabs.TabPages.AddRange(@($tabComputer,$tabPrinter,$tabSoftware,$tabReporting))

function New-OutputBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(350, 10)
    $box.Size = New-Object System.Drawing.Size(860, 540)
    $box.Multiline = $true
    $box.ScrollBars = 'Vertical'
    $box.ReadOnly = $true
    $box.Font = New-Object System.Drawing.Font('Consolas', 10)
    return $box
}

$txtComputerOut = New-OutputBox
$txtPrinterOut  = New-OutputBox
$txtSoftwareOut = New-OutputBox
$txtReportOut   = New-OutputBox
$tabComputer.Controls.Add($txtComputerOut)
$tabPrinter.Controls.Add($txtPrinterOut)
$tabSoftware.Controls.Add($txtSoftwareOut)
$tabReporting.Controls.Add($txtReportOut)

function New-Button {
    param([System.Windows.Forms.Control]$Parent,[string]$Text,[int]$X,[int]$Y,[scriptblock]$Action)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size(300, 36)
    $btn.Add_Click($Action)
    $Parent.Controls.Add($btn)
}

function Add-ComputerOutput { param([string]$Text) $txtComputerOut.AppendText($Text + [Environment]::NewLine) }
function Add-PrinterOutput  { param([string]$Text) $txtPrinterOut.AppendText($Text + [Environment]::NewLine) }
function Add-SoftwareOutput { param([string]$Text) $txtSoftwareOut.AppendText($Text + [Environment]::NewLine) }
function Add-ReportOutput   { param([string]$Text) $txtReportOut.AppendText($Text + [Environment]::NewLine) }

New-Button $tabComputer 'Ping Computer' 20 20 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $online = Test-Connection -ComputerName $computer -Count 1 -Quiet
        Add-ComputerOutput ("Ping result for {0}: {1}" -f $computer, $online)
        Set-Status -Message "Ping complete for $computer" -State Good
    } catch {
        Add-ComputerOutput ("Ping failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "Ping failed for $computer" -State Bad
    }
}

New-Button $tabComputer 'Get Logged-In User' 20 65 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $user = (Get-CimInstance Win32_ComputerSystem -ComputerName $computer).UserName
        Add-ComputerOutput ("Logged-in user on {0}: {1}" -f $computer, $user)
        Set-Status -Message "Logged-in user pulled for $computer" -State Good
    } catch {
        Add-ComputerOutput ("Lookup failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "User lookup failed for $computer" -State Bad
    }
}

New-Button $tabComputer 'Get System Info' 20 110 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $computer
        $bios = Get-CimInstance Win32_BIOS -ComputerName $computer
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $computer
        Add-ComputerOutput ("Computer: {0}" -f $computer)
        Add-ComputerOutput ("Manufacturer: {0}" -f $cs.Manufacturer)
        Add-ComputerOutput ("Model: {0}" -f $cs.Model)
        Add-ComputerOutput ("Serial: {0}" -f $bios.SerialNumber)
        Add-ComputerOutput ("OS: {0}" -f $os.Caption)
        Add-ComputerOutput ("Version: {0}" -f $os.Version)
        Add-ComputerOutput ("Last Boot: {0}" -f $os.LastBootUpTime)
        Add-ComputerOutput ''
        Set-Status -Message "System info complete for $computer" -State Good
    } catch {
        Add-ComputerOutput ("System info failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "System info failed for $computer" -State Bad
    }
}

New-Button $tabComputer 'Get Disk Space' 20 155 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $computer -Filter "DriveType=3"
        foreach ($disk in $disks) {
            Add-ComputerOutput ("{0} | Free: {1} GB | Size: {2} GB" -f $disk.DeviceID, [math]::Round($disk.FreeSpace/1GB,2), [math]::Round($disk.Size/1GB,2))
        }
        Add-ComputerOutput ''
        Set-Status -Message "Disk space pulled for $computer" -State Good
    } catch {
        Add-ComputerOutput ("Disk space failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "Disk lookup failed for $computer" -State Bad
    }
}

New-Button $tabComputer 'Full Workstation Report' 20 200 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $computer
        $bios = Get-CimInstance Win32_BIOS -ComputerName $computer
        $os = Get-CimInstance Win32_OperatingSystem -ComputerName $computer
        $disks = Get-CimInstance Win32_LogicalDisk -ComputerName $computer -Filter "DriveType=3"
        $uptime = (Get-Date) - $os.LastBootUpTime
        Add-ComputerOutput ("Workstation Report for {0}" -f $computer)
        Add-ComputerOutput ("Manufacturer: {0}" -f $cs.Manufacturer)
        Add-ComputerOutput ("Model: {0}" -f $cs.Model)
        Add-ComputerOutput ("Serial: {0}" -f $bios.SerialNumber)
        Add-ComputerOutput ("User: {0}" -f $cs.UserName)
        Add-ComputerOutput ("OS: {0}" -f $os.Caption)
        Add-ComputerOutput ("Version: {0}" -f $os.Version)
        Add-ComputerOutput ("Last Boot: {0}" -f $os.LastBootUpTime)
        Add-ComputerOutput ("Uptime Days: {0}" -f [math]::Round($uptime.TotalDays,2))
        foreach ($disk in $disks) {
            Add-ComputerOutput ("Disk {0} | Free: {1} GB | Size: {2} GB" -f $disk.DeviceID, [math]::Round($disk.FreeSpace/1GB,2), [math]::Round($disk.Size/1GB,2))
        }
        Add-ComputerOutput ''
        Set-Status -Message "Workstation report complete for $computer" -State Good
    } catch {
        Add-ComputerOutput ("Workstation report failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "Workstation report failed for $computer" -State Bad
    }
}

New-Button $tabComputer 'Export This Tab' 20 255 {
    Export-OutputToFile -TextToSave $txtComputerOut.Text -DefaultName 'ComputerTabOutput.txt'
}

New-Button $tabPrinter 'List Local Printers' 20 20 {
    try {
        $printers = Get-Printer | Select-Object Name, DriverName, PortName
        foreach ($printer in $printers) {
            Add-PrinterOutput ("{0} | {1} | {2}" -f $printer.Name, $printer.DriverName, $printer.PortName)
        }
        Add-PrinterOutput ''
        Set-Status -Message 'Local printers listed.' -State Good
    } catch {
        Add-PrinterOutput ("Printer lookup failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Printer lookup failed.' -State Bad
    }
}

New-Button $tabPrinter 'List Remote Printers' 20 65 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        $printers = Invoke-RemoteAction -ComputerName $computer -ScriptBlock { Get-Printer | Select-Object Name, DriverName, PortName }
        foreach ($printer in $printers) {
            Add-PrinterOutput ("{0} | {1} | {2}" -f $printer.Name, $printer.DriverName, $printer.PortName)
        }
        Add-PrinterOutput ''
        Set-Status -Message "Remote printers listed for $computer" -State Good
    } catch {
        Add-PrinterOutput ("Remote printer lookup failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "Remote printer lookup failed for $computer" -State Bad
    }
}

New-Button $tabPrinter 'Restart Spooler' 20 110 {
    if (-not (Require-Computer)) { return }
    $computer = Get-EnteredComputer
    try {
        Invoke-RemoteAction -ComputerName $computer -ScriptBlock { Restart-Service -Name Spooler -Force }
        Add-PrinterOutput ("Spooler restarted on {0}" -f $computer)
        Set-Status -Message "Spooler restarted on $computer" -State Good
    } catch {
        Add-PrinterOutput ("Spooler restart failed for {0}: {1}" -f $computer, $_.Exception.Message)
        Set-Status -Message "Spooler restart failed for $computer" -State Bad
    }
}

New-Button $tabPrinter 'Export This Tab' 20 155 {
    Export-OutputToFile -TextToSave $txtPrinterOut.Text -DefaultName 'PrinterTabOutput.txt'
}

New-Button $tabSoftware 'Installed Software' 20 20 {
    try {
        $paths = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        $apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | Sort-Object DisplayName | Select-Object -First 50 DisplayName, DisplayVersion, Publisher
        foreach ($app in $apps) { Add-SoftwareOutput ("{0} | {1} | {2}" -f $app.DisplayName, $app.DisplayVersion, $app.Publisher) }
        Add-SoftwareOutput ''
        Set-Status -Message 'Installed software loaded.' -State Good
    } catch {
        Add-SoftwareOutput ("Software inventory failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Software inventory failed.' -State Bad
    }
}

New-Button $tabSoftware 'Search Software' 20 65 {
    $name = $txtSoftware.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        Add-SoftwareOutput 'Enter a software search term first.'
        Set-Status -Message 'Software search term missing.' -State Bad
        return
    }
    try {
        $paths = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        $apps = Get-ItemProperty $paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -like "*$name*" } | Sort-Object DisplayName | Select-Object DisplayName, DisplayVersion, Publisher
        foreach ($app in $apps) { Add-SoftwareOutput ("{0} | {1} | {2}" -f $app.DisplayName, $app.DisplayVersion, $app.Publisher) }
        if (-not $apps) { Add-SoftwareOutput 'No matching software found.' }
        Add-SoftwareOutput ''
        Set-Status -Message "Software search complete for $name" -State Good
    } catch {
        Add-SoftwareOutput ("Software search failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Software search failed.' -State Bad
    }
}

New-Button $tabSoftware 'Startup Programs' 20 110 {
    try {
        $items = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location
        foreach ($item in $items) { Add-SoftwareOutput ("{0} | {1} | {2}" -f $item.Name, $item.Location, $item.Command) }
        Add-SoftwareOutput ''
        Set-Status -Message 'Startup programs loaded.' -State Good
    } catch {
        Add-SoftwareOutput ("Startup lookup failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Startup lookup failed.' -State Bad
    }
}

New-Button $tabSoftware 'Failed Services' 20 155 {
    try {
        $services = Get-Service | Where-Object { $_.Status -eq 'Stopped' -and $_.StartType -eq 'Automatic' }
        foreach ($svc in $services) { Add-SoftwareOutput ("{0} | {1}" -f $svc.Name, $svc.DisplayName) }
        if (-not $services) { Add-SoftwareOutput 'No failed automatic services found.' }
        Add-SoftwareOutput ''
        Set-Status -Message 'Failed services check complete.' -State Good
    } catch {
        Add-SoftwareOutput ("Failed services check failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Failed services check failed.' -State Bad
    }
}

New-Button $tabSoftware 'Get Service Status' 20 200 {
    $service = $txtService.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($service)) {
        Add-SoftwareOutput 'Enter a service name first.'
        Set-Status -Message 'Service name missing.' -State Bad
        return
    }
    try {
        $svc = Get-Service -Name $service | Select-Object Name, DisplayName, Status, StartType
        Add-SoftwareOutput ("{0} | {1} | {2} | {3}" -f $svc.Name, $svc.DisplayName, $svc.Status, $svc.StartType)
        Add-SoftwareOutput ''
        Set-Status -Message "Service status loaded for $service" -State Good
    } catch {
        Add-SoftwareOutput ("Service lookup failed: {0}" -f $_.Exception.Message)
        Set-Status -Message 'Service lookup failed.' -State Bad
    }
}

New-Button $tabSoftware 'Export This Tab' 20 245 {
    Export-OutputToFile -TextToSave $txtSoftwareOut.Text -DefaultName 'SoftwareTabOutput.txt'
}

New-Button $tabReporting 'Bulk Ping' 20 20 {
    if (-not $script:BulkComputers -or $script:BulkComputers.Count -eq 0) {
        Add-ReportOutput 'No bulk computer list loaded.'
        Set-Status -Message 'No bulk list loaded.' -State Bad
        return
    }
    foreach ($computer in $script:BulkComputers) {
        try {
            $online = Test-Connection -ComputerName $computer -Count 1 -Quiet
            Add-ReportOutput ("{0} | Online: {1}" -f $computer, $online)
        } catch {
            Add-ReportOutput ("{0} | Failed: {1}" -f $computer, $_.Exception.Message)
        }
    }
    Add-ReportOutput ''
    Set-Status -Message 'Bulk ping complete.' -State Good
}

New-Button $tabReporting 'Bulk Logged-In User' 20 65 {
    if (-not $script:BulkComputers -or $script:BulkComputers.Count -eq 0) {
        Add-ReportOutput 'No bulk computer list loaded.'
        Set-Status -Message 'No bulk list loaded.' -State Bad
        return
    }
    foreach ($computer in $script:BulkComputers) {
        try {
            $user = (Get-CimInstance Win32_ComputerSystem -ComputerName $computer).UserName
            Add-ReportOutput ("{0} | User: {1}" -f $computer, $user)
        } catch {
            Add-ReportOutput ("{0} | Failed: {1}" -f $computer, $_.Exception.Message)
        }
    }
    Add-ReportOutput ''
    Set-Status -Message 'Bulk user lookup complete.' -State Good
}

New-Button $tabReporting 'Export This Tab' 20 110 {
    Export-OutputToFile -TextToSave $txtReportOut.Text -DefaultName 'ReportingTabOutput.txt'
}

New-Button $tabReporting 'Clear Reporting Output' 20 155 {
    $txtReportOut.Clear()
    Set-Status -Message 'Reporting output cleared.' -State Neutral
}

[void]$form.ShowDialog()
