# --- Clear old creds, remove mappings, ask for username once, then remap persistently ---

# Remove old stored credentials
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

# Remove old drive mappings
$driveLetters = "G:","J:","P:","Q:","S:","T:"
foreach ($drive in $driveLetters) {
    net use $drive /delete /y 2>$null | Out-Null
}

# Ask once for username
$UserName = Read-Host "Enter username (example: ALMCHI\username)"

# First mapping will securely prompt for password
net use G: "\\kmfs.almchi.airalliance.com\groups\sales" /user:$UserName * /persistent:yes

# Remaining mappings reuse the same authenticated session
net use J: "\\kmfs.almchi.airalliance.com\KM_Departments" /persistent:yes
net use P: "\\kafs1.almchi.airalliance.com\applications" /persistent:yes
net use Q: "\\kmfs.almchi.airalliance.com\groups\qadata" /persistent:yes
net use S: "\\kmfs2.almchi.airalliance.com\manuals" /persistent:yes
net use T: "\\kmfs.almchi.airalliance.com\groups\transfer" /persistent:yes

Write-Host "Done."
pause
