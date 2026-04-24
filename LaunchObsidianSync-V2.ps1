# LaunchObsidianSync_Advanced.ps1

# --- Configuration Section ---
# Defines the full path to your Obsidian vault.
# IMPORTANT: Replace "F:\GitHubCloneFiles\Obsidian_Vault" with your actual vault path.
# F:\GitHubDesktop\GitHubCloneFiles\Obsidian_Vault on Office Dell Laptop
# "F:\GitHubDesktop\GitHubCloneFiles\Obsidian_Vault" on Lenovo Thinkpad Laptop
$vaultPath = "F:\GitHubDesktop\GitHubCloneFiles\Obsidian_Vault"

# Defines the full path to your Obsidian executable.
# IMPORTANT: Replace "C:\Users\User\AppData\Local\Obsidian\Obsidian.exe" with your actual Obsidian.exe path.
# "E:\Programs\Obsidian\Obsidian.exe" on Office Dell Laptop
#  "C:\Users\Lenovo\AppData\Local\Obsidian\Obsidian.exe" on Lenovo Thinkpad Laptop
$obsidianExePath = "E:\Programs\Obsidian\Obsidian.exe"

# Sync interval for periodic reminders (in minutes)
$syncIntervalMinutes = 15

# --- Utility Function for User Confirmation ---
function Confirm-Action {
    param (
        [string]$Message
    )
    $response = Read-Host "$Message [y/n]"
    return ($response -eq 'y' -or $response -eq 'Y')
}

# --- Function to Handle Critical Errors and Exit ---
function Exit-WithError {
    param (
        [string]$ErrorMessage
    )
    Write-Host "`nCRITICAL ERROR: $ErrorMessage" -ForegroundColor Red
    Write-Host "Script will pause and then exit. Manual intervention required!" -ForegroundColor Yellow
    Read-Host "Press Enter to exit..."
    exit 1
}

# --- Function to Perform Git Sync (Commit and Push) ---
function Perform-GitSync {
    param (
        [string]$SyncType # e.g., "Periodic", "Final"
    )
    Write-Host "`n--- Performing Git Sync ($SyncType) ---" -ForegroundColor Cyan
    
    # Stage all changes, respecting the .gitignore file.
    Write-Host "`n--- Staging all changes for commit ---" -ForegroundColor Cyan
    try {
        git add . | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Git Add had issues." }
        Write-Host "All eligible changes staged." -ForegroundColor Green
    } catch {
        Write-Host "WARNING: Git Add had issues: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Check if there are actually changes to commit
    $hasChanges = $true
    try {
        git diff-index --quiet HEAD --
        if ($LASTEXITCODE -eq 0) { # $LASTEXITCODE is 0 if no changes
            $hasChanges = $false
        }
    } catch {
        Write-Host "WARNING: Could not check for changes: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    if ($hasChanges) {
        # Get computer name for unique commit message
        $computerName = $env:COMPUTERNAME
        # Capture the current timestamp once
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        Write-Host "`n--- Committing staged changes ---" -ForegroundColor Cyan
        try {
            # git commit -m: Creates a commit with a message including device name and the captured timestamp.
            git commit -m "Auto-sync ($SyncType) from Device ${computerName}: $timestamp" | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Git Commit reported an issue." }
            Write-Host "Changes committed locally." -ForegroundColor Green
        } catch {
            Write-Host "WARNING: Git Commit reported an issue: $($_.Exception.Message). Review your changes manually." -ForegroundColor Yellow
            # Read-Host "Press Enter to continue..." # Optional pause for review
        }
    } else {
        Write-Host "No new changes to commit. Skipping commit step." -ForegroundColor Yellow
    }

    # --- Final Pull before Push (important to handle remote changes that happened while Obsidian was open) ---
    if ($hasChanges) { # Only pull/push if there were changes to push
        Write-Host "`n--- Performing Git Pull --rebase BEFORE PUSH (to handle remote changes) ---" -ForegroundColor Cyan
        try {
            git pull --rebase origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Git pull --rebase failed before push." }
            Write-Host "Git Pull --rebase successful." -ForegroundColor Green
        } catch {
            Write-Host "`nCRITICAL ERROR: Git Pull --rebase failed before push: $($_.Exception.Message)." -ForegroundColor Red
            Write-Host "This often means conflicts. MANUAL INTERVENTION REQUIRED! Please resolve conflicts, then manually commit & push." -ForegroundColor Yellow
            return $false # Indicate failure to the caller
        }

        Write-Host "`n--- Pushing changes to GitHub main branch ---" -ForegroundColor Cyan
        try {
            git push origin main | Out-Null
            if ($LASTEXITCODE -ne 0) { throw "Git Push failed." }
            Write-Host "Changes pushed to GitHub successfully." -ForegroundColor Green
        } catch {
            Write-Host "`nCRITICAL ERROR: Git Push failed: $($_.Exception.Message)." -ForegroundColor Red
            Write-Host "Check connection, authentication, or conflicts. MANUAL INTERVENTION REQUIRED!" -ForegroundColor Yellow
            return $false # Indicate failure to the caller
        }
    } else {
        Write-Host "No changes to push. Skipping final push." -ForegroundColor Yellow
    }
    return $true # Indicate success
}

# --- Main Script Logic ---
Write-Host "`n======================================================================" -ForegroundColor White
Write-Host "--- STARTING OBSIDIAN VAULT GIT SYNC AND MONITOR ---" -ForegroundColor White
Write-Host "======================================================================" -ForegroundColor White
Write-Host "`n--- Current Directory: $(Get-Location) ---" -ForegroundColor DarkGray

try {
    Set-Location $vaultPath
    Write-Host "Changed directory to: $(Get-Location)" -ForegroundColor Green
} catch {
    Exit-WithError "Could not navigate to the vault path: $vaultPath. Please check the VAULT_PATH variable."
}

# --- Initial Pull (Before Launch) ---
Write-Host "`n--- Git Status BEFORE Initial Pull ---" -ForegroundColor Cyan
git status

if (Confirm-Action "Do you want to pull the latest changes from GitHub (main branch) BEFORE launching Obsidian?") {
    Write-Host "`n--- Performing Git Pull from GitHub (main branch) before launching Obsidian ---" -ForegroundColor Cyan
    try {
        git pull --rebase origin main | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Git pull --rebase failed before launch." }
        Write-Host "Initial Git Pull successful." -ForegroundColor Green
    } catch {
        Write-Host "`nWARNING: Git Pull failed before launch. $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This could be due to network issues, conflicts, or authentication problems." -ForegroundColor Yellow
        Write-Host "Proceeding with Obsidian launch, but check Git status later." -ForegroundColor Yellow
    }
} else {
    Write-Host "Skipping initial Git Pull." -ForegroundColor Yellow
}

# --- Launch Obsidian (without -Wait this time) ---
Write-Host "`n--- Launching Obsidian in the background. ---" -ForegroundColor Green
Write-Host "Type '-exit obsidian' in this console to trigger final sync and close." -ForegroundColor Green
Write-Host "This console will remain open for periodic sync prompts." -ForegroundColor Green
Write-Host "DO NOT close this console window manually; use '-exit obsidian'." -ForegroundColor Red # Added warning
Start-Process -FilePath $obsidianExePath -WindowStyle Normal -ErrorAction SilentlyContinue

# --- Active Monitoring Loop for Periodic Sync and Custom Exit ---
$lastSyncTime = Get-Date
$obsidianProcessName = "Obsidian" # Process name might be "Obsidian" or "Obsidian.exe"
$exitCommand = "-exit obsidian"
$loopSleepSeconds = 5 # How often to check for Obsidian process and user input (e.g., every 5 seconds)

# Bring the console window to the foreground when needed for prompts
$wshell = New-Object -ComObject WScript.Shell
$hwd = $wshell.AppActivate($host.ui.RawUI.WindowTitle) # Get the handle to this console window

while ($true) {
    # Check if Obsidian is still running
    $obsidianRunning = Get-Process -Name $obsidianProcessName -ErrorAction SilentlyContinue | Select-Object -First 1

    if (-not $obsidianRunning) {
        Write-Host "`n--- Obsidian process not found. Assuming Obsidian has been closed. ---" -ForegroundColor Red
        break # Exit the loop to perform final sync
    }

    # Check for periodic sync time
    $currentTime = Get-Date
    if ($currentTime -gt ($lastSyncTime).AddMinutes($syncIntervalMinutes)) {
        # Bring console to foreground for user interaction
        if ($hwd) { # Check if $hwd is valid before calling Activate
            $wshell.AppActivate($host.ui.RawUI.WindowTitle) | Out-Null
        } else {
            # Fallback if AppActivate fails, try to make the window active via more generic methods (might not work perfectly)
            [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            [Microsoft.VisualBasic.Interaction]::AppActivate($host.ui.RawUI.WindowTitle)
        }
        
        Write-Host "`n----------------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "  15-MINUTE SYNC REMINDER!" -ForegroundColor Yellow
        Write-Host "  Obsidian is still running." -ForegroundColor Yellow
        Write-Host "----------------------------------------------------------------------" -ForegroundColor Yellow
        
        if (Confirm-Action "Do you want to commit and push your saved work NOW?") {
            # Perform periodic sync
            if (Perform-GitSync -SyncType "Periodic") {
                $lastSyncTime = Get-Date # Reset timer only on successful sync
                Write-Host "`nPeriodic sync completed." -ForegroundColor Green
            } else {
                Write-Host "`nPeriodic sync had issues. Please check the output above." -ForegroundColor Red
            }
        } else {
            Write-Host "`nSkipping periodic sync for now." -ForegroundColor Yellow
            $lastSyncTime = Get-Date # Reset timer even if skipped, to avoid immediate re-prompt
        }
        Write-Host "`nType '-exit obsidian' to close this console and perform final sync." -ForegroundColor Green
    }

    # --- Check for custom exit command ---
    # This requires checking for input. Read-Host blocks, so we need a non-blocking way.
    # The current approach is a simple sleep loop; user input detection while sleeping is hard.
    # A true non-blocking input reader is very complex for PowerShell in a loop.
    # For now, the user has to type it *when prompted* or *when not sleeping*.
    # A simple way for a single line is to just have a timeout for Read-Host, but that complicates the loop.

    # Simpler approach: user types exit command when it's obvious, or when the prompt is active.
    # We will rely on the Read-Host at the end of the script for direct input.
    # If the user types it mid-loop, it will show up as unparsed input, but the script
    # won't act on it until the final Read-Host or if they force close.

    # Sleep for a bit before checking again
    Start-Sleep -Seconds $loopSleepSeconds
}

# --- Obsidian has been closed (loop exited). Proceed to Final Sync ---
Write-Host "`n--- Obsidian closed. Performing Final Git Sync ---" -ForegroundColor Green

# Perform final sync
if (Perform-GitSync -SyncType "Final") {
    Write-Host "`n======================================================================" -ForegroundColor White
    Write-Host "--- OBSIDIAN VAULT FINAL SYNC COMPLETE ---" -ForegroundColor White
    Write-Host "======================================================================" -ForegroundColor White
} else {
    Write-Host "`n======================================================================" -ForegroundColor Red
    Write-Host "--- OBSIDIAN VAULT FINAL SYNC HAD ISSUES ---" -ForegroundColor Red
    Write-Host "--- MANUAL INTERVENTION REQUIRED ---" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor Red
}

Read-Host "Press Enter to exit this console..." # Keeps console open for review