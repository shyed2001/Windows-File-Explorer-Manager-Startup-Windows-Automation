# =============================================================================
# Obsidian_Backup_V3.ps1
# Purpose:
#   Safe timestamped Obsidian Vault backup to Google Drive.
#
# Features:
#   1. Creates date-time stamped backup folders.
#   2. Keeps latest 3 backup folders.
#   3. Logs every backup run to JSON.
#   4. Warns before cleaning old backups.
#   5. Moves old backups to _Old_Backups_To_Delete before permanent deletion.
#   6. Does NOT use /PURGE against the active backup root.
#
# Project:
#   Windows Server WinVPS Local Cloud-storage RDP
# =============================================================================

# ---------------- USER CONFIG ----------------

$sourcePath = "F:\GitHubDesktop\GitHubCloneFiles\Obsidian_Vault"

$backupRoot = "F:\Google_Drive_Directory_Streaming\My Drive\Obsidian_Backup"

$keepLatestBackups = 3

# Safer option: "Quarantine"
# Other option: "PermanentDelete"
# Recommendation: keep "Quarantine" unless you are 100% sure.
$cleanupMode = "Quarantine"

# Ask before cleanup?
$requireCleanupApproval = $true

# Backup folder name prefix
$backupPrefix = "Obsidian_Backup_"

# ---------------- INTERNAL PATHS ----------------

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$humanTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$currentBackupName = "$backupPrefix$timestamp"
$currentBackupPath = Join-Path $backupRoot $currentBackupName

$logPath = Join-Path $backupRoot "_backup_log.json"
$robocopyLogRoot = Join-Path $backupRoot "_robocopy_logs"
$oldBackupHoldRoot = Join-Path $backupRoot "_Old_Backups_To_Delete"

$robocopyLogPath = Join-Path $robocopyLogRoot "robocopy_$timestamp.log"

# ---------------- FUNCTIONS ----------------

function Get-FolderSizeBytes {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return 0
    }

    $size = 0
    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
        ForEach-Object {
            if (-not $_.PSIsContainer) {
                $size += $_.Length
            }
        }

    return $size
}

function Read-BackupLog {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return @()
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @()
        }

        $data = $raw | ConvertFrom-Json
        if ($null -eq $data) {
            return @()
        }

        if ($data -is [System.Array]) {
            return @($data)
        }

        return @($data)
    }
    catch {
        Write-Host "WARNING: Could not read existing backup log. A new log structure will be written." -ForegroundColor Yellow
        return @()
    }
}

function Write-BackupLog {
    param(
        [string]$Path,
        [array]$Records
    )

    $Records |
        ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $Path -Encoding UTF8
}

function Confirm-Yes {
    param([string]$Message)

    $answer = Read-Host "$Message Type YES to continue"
    return ($answer -eq "YES")
}

# ---------------- START ----------------

Write-Host ""
Write-Host "======================================================================" -ForegroundColor White
Write-Host " OBSIDIAN GOOGLE DRIVE TIMESTAMPED BACKUP - V3" -ForegroundColor White
Write-Host "======================================================================" -ForegroundColor White
Write-Host "Time:        $humanTime"
Write-Host "Source:      $sourcePath"
Write-Host "Backup Root: $backupRoot"
Write-Host "New Backup:  $currentBackupPath"
Write-Host ""

# ---------------- VALIDATION ----------------

if (-not (Test-Path $sourcePath)) {
    Write-Host "ERROR: Source path does not exist:" -ForegroundColor Red
    Write-Host $sourcePath -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null
New-Item -ItemType Directory -Force -Path $robocopyLogRoot | Out-Null
New-Item -ItemType Directory -Force -Path $oldBackupHoldRoot | Out-Null
New-Item -ItemType Directory -Force -Path $currentBackupPath | Out-Null

# ---------------- BACKUP COPY ----------------

Write-Host "Starting timestamped backup copy..." -ForegroundColor Cyan

# IMPORTANT:
# No /PURGE here.
# This is a fresh timestamped snapshot folder.
robocopy $sourcePath $currentBackupPath /E /FFT /R:2 /W:5 /NP /TEE /LOG+:$robocopyLogPath

$robocopyExitCode = $LASTEXITCODE
$backupSuccess = ($robocopyExitCode -lt 8)

if ($backupSuccess) {
    Write-Host ""
    Write-Host "SUCCESS: Backup copy completed. Robocopy exit code: $robocopyExitCode" -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host "ERROR: Backup copy failed or had serious errors. Robocopy exit code: $robocopyExitCode" -ForegroundColor Red
}

$currentBackupSizeBytes = Get-FolderSizeBytes -Path $currentBackupPath

# ---------------- FIND OLD BACKUPS ----------------

$allBackupFolders =
    Get-ChildItem -LiteralPath $backupRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -like "$backupPrefix*" -and
        $_.FullName -ne $currentBackupPath
    } |
    Sort-Object Name -Descending

# Include current folder in retention calculation
$allBackupFoldersIncludingCurrent =
    Get-ChildItem -LiteralPath $backupRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$backupPrefix*" } |
    Sort-Object Name -Descending

$foldersToCleanup = @($allBackupFoldersIncludingCurrent | Select-Object -Skip $keepLatestBackups)

$cleanupRecords = @()
$cleanupApproved = $false

if ($foldersToCleanup.Count -gt 0) {
    Write-Host ""
    Write-Host "Older backup folders detected beyond latest $keepLatestBackups:" -ForegroundColor Yellow

    foreach ($folder in $foldersToCleanup) {
        $sizeBytes = Get-FolderSizeBytes -Path $folder.FullName
        $sizeMB = [math]::Round($sizeBytes / 1MB, 2)

        Write-Host " - $($folder.FullName) [$sizeMB MB]" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "WARNING:" -ForegroundColor Red
    Write-Host "These old backup folders are outside the latest $keepLatestBackups retention limit." -ForegroundColor Red
    Write-Host "Recommended safe action: move them to _Old_Backups_To_Delete first, then review manually." -ForegroundColor Yellow
    Write-Host ""

    if ($requireCleanupApproval) {
        $cleanupApproved = Confirm-Yes "Approve cleanup of old backup folders?"
    }
    else {
        $cleanupApproved = $true
    }

    if ($cleanupApproved) {
        foreach ($folder in $foldersToCleanup) {
            $folderSizeBytes = Get-FolderSizeBytes -Path $folder.FullName
            $targetPath = Join-Path $oldBackupHoldRoot $folder.Name

            $record = [ordered]@{
                folder_name       = $folder.Name
                original_path     = $folder.FullName
                size_bytes        = $folderSizeBytes
                cleanup_mode      = $cleanupMode
                cleanup_status    = "Pending"
                cleanup_time      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                target_path       = $targetPath
                error             = $null
            }

            try {
                if ($cleanupMode -eq "Quarantine") {
                    if (Test-Path $targetPath) {
                        $targetPath = Join-Path $oldBackupHoldRoot "$($folder.Name)_moved_$timestamp"
                        $record.target_path = $targetPath
                    }

                    Move-Item -LiteralPath $folder.FullName -Destination $targetPath -Force
                    $record.cleanup_status = "MovedToQuarantine"
                    Write-Host "Moved old backup to quarantine: $targetPath" -ForegroundColor Green
                }
                elseif ($cleanupMode -eq "PermanentDelete") {
                    Remove-Item -LiteralPath $folder.FullName -Recurse -Force
                    $record.cleanup_status = "PermanentlyDeleted"
                    Write-Host "Permanently deleted old backup: $($folder.FullName)" -ForegroundColor Red
                }
                else {
                    $record.cleanup_status = "SkippedInvalidCleanupMode"
                    $record.error = "Unknown cleanup mode: $cleanupMode"
                    Write-Host "Skipped cleanup due to invalid cleanup mode: $cleanupMode" -ForegroundColor Yellow
                }
            }
            catch {
                $record.cleanup_status = "Failed"
                $record.error = $_.Exception.Message
                Write-Host "ERROR cleaning folder: $($folder.FullName)" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            $cleanupRecords += [pscustomobject]$record
        }
    }
    else {
        Write-Host "Cleanup skipped by user approval decision." -ForegroundColor Yellow
    }
}
else {
    Write-Host ""
    Write-Host "No old backup folders need cleanup. Latest $keepLatestBackups retention is clean." -ForegroundColor Green
}

# ---------------- WRITE JSON LOG ----------------

$existingLog = Read-BackupLog -Path $logPath

$runRecord = [ordered]@{
    run_time                 = $humanTime
    source_path              = $sourcePath
    backup_root              = $backupRoot
    backup_folder_name       = $currentBackupName
    backup_path              = $currentBackupPath
    backup_size_bytes        = $currentBackupSizeBytes
    robocopy_exit_code       = $robocopyExitCode
    robocopy_log_path        = $robocopyLogPath
    backup_success           = $backupSuccess
    keep_latest_backups      = $keepLatestBackups
    cleanup_mode             = $cleanupMode
    cleanup_approval_needed  = $requireCleanupApproval
    cleanup_approved         = $cleanupApproved
    cleanup_records          = $cleanupRecords
}

$newLog = @($existingLog) + @([pscustomobject]$runRecord)

Write-BackupLog -Path $logPath -Records $newLog

Write-Host ""
Write-Host "Backup log updated:" -ForegroundColor Cyan
Write-Host $logPath

Write-Host ""
Write-Host "Robocopy log saved:" -ForegroundColor Cyan
Write-Host $robocopyLogPath

Write-Host ""
if ($backupSuccess) {
    Write-Host "FINAL STATUS: BACKUP SUCCESSFUL" -ForegroundColor Green
}
else {
    Write-Host "FINAL STATUS: BACKUP HAD ERRORS - REVIEW ROBOCOPY LOG" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to close"