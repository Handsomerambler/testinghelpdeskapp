# Kellstrom Drive Mapping v3
# Clears old mappings, asks for username once, then maps drives persistently.
# Password is prompted securely by net use.

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

$driveLetters = "G:","J:","P:","Q:","S:","T:"
foreach ($drive in $driveLetters) {
    net use $drive /delete /y 2>$null | Out-Null
}

$UserName = Read-Host "Enter username (example: ALMCHI\username)"

net use G: "\\kmfs.almchi.airalliance.com\groups\sales" /user:$UserName * /persistent:yes
net use J: "\\kmfs.almchi.airalliance.com\KM_Departments" /persistent:yes
net use P: "\\kafs1.almchi.airalliance.com\applications" /persistent:yes
net use Q: "\\kmfs.almchi.airalliance.com\groups\qadata" /persistent:yes
net use S: "\\kmfs2.almchi.airalliance.com\manuals" /persistent:yes
net use T: "\\kmfs.almchi.airalliance.com\groups\transfer" /persistent:yes

Write-Host "Kellstrom drive mapping v3 complete."
pause
