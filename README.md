# SharePoint Channel File Migration (PowerShell 5.1)

Recursively moves or copies all files and folders from one 
Microsoft Teams SharePoint channel to another using PnP PowerShell.

## Requirements
- PowerShell 5.1
- SharePointPnPPowerShellOnline module

## Setup
```powershell
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
```

## Usage
```powershell
# Dry run - no changes made
.\Move-SharePointChannelFiles.ps1 -WhatIf

# Copy only - keeps source files
.\Move-SharePointChannelFiles.ps1 -CopyOnly

# Full move - copies to destination and deletes source
.\Move-SharePointChannelFiles.ps1
```

## Configuration
Edit these variables at the top of the script:
- `$SourceSiteUrl` — Source SharePoint site URL
- `$SourceFolderRel` — Source channel folder path
- `$DestSiteUrl` — Destination SharePoint site URL  
- `$DestFolderAbs` — Destination channel folder path

## Features
- Recursive copy (all files, subfolders, nested levels)
- Dry run mode (`-WhatIf`)
- Copy-only mode (no deletion)
- Timestamped log file saved locally
- Confirmation prompt before any changes
