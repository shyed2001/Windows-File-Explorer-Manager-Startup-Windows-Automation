# =============================================================================
# PowerShell Script to EFFICIENTLY SYNCHRONIZE Obsidian Vault to Google Drive
# TRULY FINAL VERSION - Corrected for cloud sync timestamp precision issues.
#
# FEATURES:
# 1. Deletes files from the backup if they are deleted from the source (/PURGE).
# 2. Skips unchanged files.
# 3. Ignores timestamp differences of up to 2 seconds, fixing the re-copy loop (/FFT).
# =============================================================================

# --- User Variables ---
# Source: Your original Obsidian vault in your Git directory
$sourcePath = "F:\GitHubDesktop\GitHubCloneFiles\Obsidian_Vault"

# Destination: Your mirrored Google Drive directory for the backup
$destinationPath = "F:\Google_Drive_Directory_Streaming\My Drive\Obsidian_Backup"
# ----------------------

# --- Script Logic ---
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$timestamp] Starting robust vault synchronization..."
Write-Host "Source: $sourcePath"
Write-Host "Destination: $destinationPath"
Write-Host "NOTE: Using /FFT to handle cloud timestamp differences."
Write-Host ""

# Check if the source directory exists before trying to sync
if (-not (Test-Path $sourcePath)) {
    Write-Host "[$timestamp] ERROR: Source path not found."
    Write-Host "The script will close in 15 seconds."
    Start-Sleep -Seconds 15
    exit
}

# --- Robocopy Command ---
# /E       :: Copies subdirectories, including empty ones.
# /PURGE   :: Deletes destination files/dirs that no longer exist in the source.
# /XO      :: eXcludes Older files (to work peacefully with the sync client).
# /FFT     :: Assume FAT File Times (2-second granularity). THIS IS THE KEY FIX.
# /R:2     :: Retry failed copies 2 times.
# /W:5     :: Wait 5 seconds between retries.
# /NP      :: No Progress - doesn't display the % copied.
# /NJH     :: No Job Header.
# /NJS     :: No Job Summary.
robocopy $sourcePath $destinationPath /E /PURGE /XO /FFT /R:2 /W:5 /NP /NJH /NJS

# Robocopy returns exit codes. 0-7 are success, 8+ are failures.
if ($LASTEXITCODE -lt 8) {
    Write-Host ""
    Write-Host "[$timestamp] SUCCESS: Vault synchronization complete."
} else {
    Write-Host ""
    Write-Host "[$timestamp] ERROR: Robocopy finished with an error code ($LASTEXITCODE). Please review the output above."
}

Write-Host "Process finished. The window will close in 7 seconds."
Start-Sleep -Seconds 7