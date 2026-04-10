# Combined Installer v1
# Installs common apps with winget, then remaps Kellstrom drives.
# Run PowerShell as Administrator for best results.

Write-Host "Starting combined installer..." -ForegroundColor Cyan

# -----------------------------
# App installs via winget
# -----------------------------
$apps = @(
    "Discord.Discord",
    "Microsoft.PowerShell",
    "Microsoft.VisualStudioCode",
    "Git.Git",
    "GitHub.GitHubDesktop",
    "Python.Python.3.12",
    "Microsoft.WindowsTerminal",
    "Notepad++.Notepad++"
)

foreach ($app in $apps) {
    Write-Host "Installing $app ..." -ForegroundColor Yellow
    try {
        winget install -e --id $app --accept-source-agreements --accept-package-agreements --silent
    }
    catch {
        Write-Warning "Failed to install $app"
    }
}

# -----------------------------
# Clear old stored credentials
# -----------------------------
$targets = @(
    "kmfs.almchi.airalliance.com",
    "kafs1.almchi.airalliance.com",
    "kmfs2.almchi.airalliance.com",
    "kafs2",
    "KMFS"
)

foreach ($target in $targets) {
    cmdkey /delete:$target 2>$null | Out-Null
}

# -----------------------------
# Remove old drive mappings
# -----------------------------
$driveLetters = "G:","J:","P:","Q:","S:","T:"
foreach ($drive in $driveLetters) {
    net use $drive /delete /y 2>$null | Out-Null
}

# -----------------------------
# Prompt once for username
# -----------------------------
$UserName = Read-Host "Enter Kellstrom username (example: ALMCHI\username)"

# First mapping prompts securely for password
Write-Host "You may be prompted once for your Kellstrom password..." -ForegroundColor Cyan
net use G: "\\kmfs.almchi.airalliance.com\groups\sales" /user:$UserName * /persistent:yes

# Remaining mappings reuse the same authenticated session
net use J: "\\kmfs.almchi.airalliance.com\KM_Departments" /persistent:yes
net use P: "\\kafs1.almchi.airalliance.com\applications" /persistent:yes
net use Q: "\\kmfs.almchi.airalliance.com\groups\qadata" /persistent:yes
net use S: "\\kmfs2.almchi.airalliance.com\manuals" /persistent:yes
net use T: "\\kmfs.almchi.airalliance.com\groups\transfer" /persistent:yes

Write-Host "Combined installer complete." -ForegroundColor Green
pause
