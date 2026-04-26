# ============================================================
#  Move-SharePointChannelFiles.ps1  (PowerShell 5.1 Compatible)
#  Recursively copies ALL files AND folders from:
#    Old Team -> Reliability Engineering
#  To:
#    New Team -> Reliability Engineering
#
#  REQUIREMENTS:
#    PowerShell 5.1
#    Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
#
#  RUN:
#    .\Move-SharePointChannelFiles.ps1              # Move (copy + delete source)
#    .\Move-SharePointChannelFiles.ps1 -WhatIf     # Dry run (no changes made)
#    .\Move-SharePointChannelFiles.ps1 -CopyOnly   # Copy without deleting source
# ============================================================

param(
    [switch]$WhatIf,
    [switch]$CopyOnly
)

# ─── CONFIG ────────────────────────────────────────────────
$SourceSiteUrl      = "https://<Tenant_Name>.sharepoint.com/sites/<SourceTeamName>"
$SourceFolderRel    = "Shared Documents/Reliability Engineering"   # Site-relative

$DestSiteUrl        = "https://<Tenant_Name>.sharepoint.com/teams/Newteam"
$DestFolderAbs      = "/teams/Newteam/Shared Documents/Reliability Engineering"  # Server-relative

$LogFile            = "$PSScriptRoot\MoveLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

$TotalCopied  = 0
$TotalFailed  = 0
$TotalDeleted = 0
# ───────────────────────────────────────────────────────────

# ─── LOGGING ───────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line      = "[$timestamp] [$Level] $Message"
    $color     = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Cyan" }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line
}
# ───────────────────────────────────────────────────────────

# ─── RECURSIVE COPY FUNCTION ───────────────────────────────
# Walks every folder level and copies files + subfolders
function Copy-FolderRecursive {
    param(
        [string]$SrcFolderRelUrl,   # e.g. "Shared Documents/Reliability Engineering/Automation"
        [string]$DstFolderAbsUrl    # e.g. "/teams/.../Shared Documents/Reliability Engineering/Automation"
    )

    # ── Copy files in current folder ──────────────────────
    try {
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType File -ErrorAction Stop
    } catch {
        Write-Log "  Could not list files in '$SrcFolderRelUrl': $_" "ERROR"
        $files = @()
    }

    foreach ($file in $files) {
        $fileName   = $file.Name
        $srcPath    = "$SrcFolderRelUrl/$fileName"
        $dstPath    = "$DstFolderAbsUrl/$fileName"

        Write-Log "  FILE: $srcPath"

        if ($WhatIf) {
            Write-Log "    [WHATIF] Would copy file -> $dstPath" "WARN"
            $script:TotalCopied++
            continue
        }

        try {
            Copy-PnPFile `
                -SourceUrl  $srcPath `
                -TargetUrl  $dstPath `
                -Force `
                -OverwriteIfAlreadyExists `
                -ErrorAction Stop

            Write-Log "    Copied file '$fileName'" "SUCCESS"
            $script:TotalCopied++

        } catch {
            Write-Log "    FAILED to copy file '$fileName': $_" "ERROR"
            $script:TotalFailed++
        }
    }

    # ── Recurse into subfolders ────────────────────────────
    try {
        $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $SrcFolderRelUrl -ItemType Folder -ErrorAction Stop
    } catch {
        Write-Log "  Could not list subfolders in '$SrcFolderRelUrl': $_" "ERROR"
        $subFolders = @()
    }

    foreach ($sub in $subFolders) {
        $subName       = $sub.Name
        $subSrcRelUrl  = "$SrcFolderRelUrl/$subName"
        $subDstAbsUrl  = "$DstFolderAbsUrl/$subName"

        Write-Log "  FOLDER: $subSrcRelUrl"

        if ($WhatIf) {
            Write-Log "    [WHATIF] Would create folder -> $subDstAbsUrl" "WARN"
            # Still recurse to list children in WhatIf
            Copy-FolderRecursive -SrcFolderRelUrl $subSrcRelUrl -DstFolderAbsUrl $subDstAbsUrl
            continue
        }

        # Ensure destination subfolder exists
        try {
            Resolve-PnPFolder -SiteRelativePath ($subDstAbsUrl.TrimStart('/')) -ErrorAction Stop | Out-Null
        } catch {
            Write-Log "    Could not ensure destination folder '$subName': $_" "WARN"
        }

        # Recurse into subfolder
        Copy-FolderRecursive -SrcFolderRelUrl $subSrcRelUrl -DstFolderAbsUrl $subDstAbsUrl
    }
}
# ───────────────────────────────────────────────────────────

# ─── CHECK / INSTALL MODULE ────────────────────────────────
$moduleName = "SharePointPnPPowerShellOnline"
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Log "$moduleName not found. Installing..." "WARN"
    Install-Module $moduleName -Scope CurrentUser -Force -AllowClobber
}
Import-Module $moduleName -ErrorAction Stop -WarningAction SilentlyContinue
Write-Log "Module '$moduleName' loaded." "SUCCESS"
# ───────────────────────────────────────────────────────────

Write-Log "========================================"
Write-Log "  Old -> New Full Migration Script"
Write-Log "  PS Version : $($PSVersionTable.PSVersion)"
Write-Log "  Source     : $SourceSiteUrl/$SourceFolderRel"
Write-Log "  Destination: $DestSiteUrl$DestFolderAbs"
if ($WhatIf)   { Write-Log "  MODE: DRY RUN (no changes will be made)" "WARN" }
if ($CopyOnly) { Write-Log "  MODE: COPY ONLY (source files kept)"     "WARN" }
Write-Log "========================================"

# ─── CONNECT TO SOURCE ─────────────────────────────────────
Write-Log "Connecting to SOURCE site (browser login will open)..."
try {
    Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin -ErrorAction Stop
    Write-Log "Connected to SOURCE site." "SUCCESS"
} catch {
    Write-Log "Failed to connect to SOURCE: $_" "ERROR"
    exit 1
}

# ─── GET TOP-LEVEL ITEMS ───────────────────────────────────
Write-Log "Scanning top-level items in: $SourceFolderRel"

try {
    $topFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $SourceFolderRel -ItemType Folder -ErrorAction Stop
    $topFiles   = Get-PnPFolderItem -FolderSiteRelativeUrl $SourceFolderRel -ItemType File   -ErrorAction Stop
} catch {
    Write-Log "Failed to scan source folder: $_" "ERROR"
    Disconnect-PnPOnline
    exit 1
}

$totalItems = $topFolders.Count + $topFiles.Count
Write-Log "Found $($topFolders.Count) folder(s) and $($topFiles.Count) file(s) at root level."

if ($totalItems -eq 0) {
    Write-Log "Nothing found in source. Exiting." "WARN"
    Disconnect-PnPOnline
    exit 0
}

Write-Log "Top-level folders:"
foreach ($f in $topFolders) { Write-Log "   [FOLDER] $($f.Name)" }
Write-Log "Top-level files:"
foreach ($f in $topFiles)   { Write-Log "   [FILE]   $($f.Name)" }

# ─── CONFIRM ───────────────────────────────────────────────
if (-not $WhatIf) {
    Write-Log ""
    $confirm = Read-Host "Proceed with copying ALL files and folders (including subfolders) to New Team? (yes/no)"
    if ($confirm -ne "yes") {
        Write-Log "Cancelled by user." "WARN"
        Disconnect-PnPOnline
        exit 0
    }
}

# ─── COPY TOP-LEVEL FILES ──────────────────────────────────
foreach ($file in $topFiles) {
    $fileName = $file.Name
    $srcPath  = "$SourceFolderRel/$fileName"
    $dstPath  = "$DestFolderAbs/$fileName"

    Write-Log "Copying root file: '$fileName'"

    if ($WhatIf) {
        Write-Log "  [WHATIF] Would copy -> $dstPath" "WARN"
        $TotalCopied++
        continue
    }

    try {
        Copy-PnPFile `
            -SourceUrl  $srcPath `
            -TargetUrl  $dstPath `
            -Force `
            -OverwriteIfAlreadyExists `
            -ErrorAction Stop

        Write-Log "  Copied '$fileName'" "SUCCESS"
        $TotalCopied++

    } catch {
        Write-Log "  FAILED for '$fileName': $_" "ERROR"
        $TotalFailed++
    }
}

# ─── RECURSIVELY COPY ALL FOLDERS ──────────────────────────
foreach ($folder in $topFolders) {
    $folderName    = $folder.Name
    $srcFolderRel  = "$SourceFolderRel/$folderName"
    $dstFolderAbs  = "$DestFolderAbs/$folderName"

    Write-Log ""
    Write-Log "====[ FOLDER: $folderName ]===="

    if (-not $WhatIf) {
        # Ensure top-level destination folder exists
        try {
            Resolve-PnPFolder -SiteRelativePath ($dstFolderAbs.TrimStart('/')) -ErrorAction Stop | Out-Null
            Write-Log "  Destination folder ready." "SUCCESS"
        } catch {
            Write-Log "  Could not create destination folder '$folderName': $_" "WARN"
        }
    }

    # Recursively copy everything inside
    Copy-FolderRecursive -SrcFolderRelUrl $srcFolderRel -DstFolderAbsUrl $dstFolderAbs

    # Delete source folder after copy (if not CopyOnly)
    if (-not $CopyOnly -and -not $WhatIf) {
        try {
            Remove-PnPFolder -Name $folderName -Folder $SourceFolderRel -Force -ErrorAction Stop
            Write-Log "  Deleted source folder '$folderName'." "SUCCESS"
            $TotalDeleted++
        } catch {
            Write-Log "  Could not delete source folder '$folderName': $_" "WARN"
        }
    }
}

# ─── DISCONNECT ────────────────────────────────────────────
Disconnect-PnPOnline
Write-Log ""
Write-Log "Disconnected from SharePoint."

# ─── SUMMARY ───────────────────────────────────────────────
Write-Log "========================================"
Write-Log "  MIGRATION COMPLETE"
Write-Log "  Items Copied  : $TotalCopied"
Write-Log "  Items Failed  : $TotalFailed"
Write-Log "  Folders Deleted (source): $TotalDeleted"
Write-Log "  Log saved     : $LogFile"
Write-Log "========================================"

if ($TotalFailed -gt 0) {
    Write-Log "Some items failed. Check the log: $LogFile" "WARN"
    exit 1
} else {
    Write-Log "All files and folders migrated successfully!" "SUCCESS"
    exit 0
}
