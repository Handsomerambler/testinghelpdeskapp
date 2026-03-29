# Help Desk Toolkit v2

PowerShell toolkit for common help desk and junior sysadmin tasks.

## Current features
- Computer diagnostics
- Printer tools
- Network tools
- Software and process tools
- Event log tools
- Active Directory tools
- Microsoft 365 / Entra / Graph tools
- Reporting and CSV export

## Requirements
- Windows PowerShell 5.1+ or PowerShell 7
- Admin rights for some local and remote actions
- RSAT Active Directory tools for AD features
- Microsoft.Entra and/or Microsoft.Graph modules for cloud features
- PowerShell remoting / WinRM for some remote tasks

## Run
```powershell
.\helpdesk_toolkit_v2.ps1
```

## Notes
Some functions require environment-specific permissions and may fail if remoting, modules, or tenant permissions are not available.
