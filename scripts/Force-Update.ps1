# -------------------------------------------------
# Path to PsExec
# -------------------------------------------------
$psexecPath = "C:\...\PSTools\PsExec.exe"

# -------------------------------------------------
# Remote script blocks (saved as temp .ps1 files)
# -------------------------------------------------
$moduleScriptContent = @'
# Force NuGet provider install silently before anything else
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null

# Trust the PSGallery so Install-Module doesn't prompt
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-Module PSWindowsUpdate -Force -Scope AllUsers -ErrorAction Stop
    Write-Output 'MODULE_INSTALLED'
} else {
    Write-Output 'MODULE_ALREADY_PRESENT'
}
'@

$updateScriptContent = @'
Import-Module PSWindowsUpdate -ErrorAction Stop
$updates = Get-WindowsUpdate -AcceptAll -ErrorAction Stop
if ($updates.Count -eq 0) {
    Write-Output 'NO_UPDATES_NEEDED'
} else {
    Write-Output "UPDATES_FOUND:$($updates.Count)"
    Install-WindowsUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop |
        ForEach-Object { Write-Output "PATCH:$($_.Title)" }
    Write-Output 'INSTALL_COMPLETE'
}
'@

# -------------------------------------------------
# Import Hosts
# -------------------------------------------------
$hosts   = Get-Content "C:\Temp\hosts.txt"
$results = @()

foreach ($computer in $hosts) {

    Write-Host "Processing $computer..." -ForegroundColor Cyan
    $status      = "Success"
    $description = ""

    # Paths on the REMOTE machine
    $remoteTemp        = "\\$computer\C$\Temp"
    $remoteModScript   = "\\$computer\C$\Temp\wu_install_module.ps1"
    $remoteUpdScript   = "\\$computer\C$\Temp\wu_run_updates.ps1"
    $localModScript    = "C:\Temp\wu_install_module.ps1"
    $localUpdScript    = "C:\Temp\wu_run_updates.ps1"

    try {
        # -----------------------------------------
        # Stage 1: Connectivity Check
        # -----------------------------------------
        if (-not (Test-Connection $computer -Count 1 -Quiet)) {
            $status      = "Offline"
            $description = "Host unreachable"
            throw "Host unreachable"
        }
        Write-Host "  [1/3] $computer is online" -ForegroundColor Green

        # -----------------------------------------
        # Stage 2: Install PSWindowsUpdate Module
        # -----------------------------------------
        Write-Host "  [2/3] Installing PSWindowsUpdate on $computer..."

        # Write script locally then copy via UNC share
        $moduleScriptContent | Set-Content -Path $localModScript -Encoding UTF8
        if (-not (Test-Path $remoteTemp)) { New-Item -Path $remoteTemp -ItemType Directory -Force | Out-Null }
        Copy-Item -Path $localModScript -Destination $remoteModScript -Force

        $installOutput = & $psexecPath "\\$computer" -accepteula -s -nobanner `
            powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass `
            -File "C:\Temp\wu_install_module.ps1" 2>&1

        $installExit = $LASTEXITCODE

        if ($installExit -ne 0) {
            $status      = "Failed - Module Install"
            $description = "PsExec exit $installExit | $($installOutput -join ' ')"
            throw $description
        }

        $moduleNote = if ($installOutput -match "MODULE_INSTALLED") { "Module freshly installed" }
                      elseif ($installOutput -match "MODULE_ALREADY_PRESENT") { "Module already present" }
                      else { "Module status unknown" }

        Write-Host "  [2/3] $moduleNote on $computer" -ForegroundColor Green

        # -----------------------------------------
        # Stage 3: Run Windows Updates
        # -----------------------------------------
        Write-Host "  [3/3] Running Windows Update on $computer..."

        $updateScriptContent | Set-Content -Path $localUpdScript -Encoding UTF8
        Copy-Item -Path $localUpdScript -Destination $remoteUpdScript -Force

        $updateOutput = & $psexecPath "\\$computer" -accepteula -s -nobanner `
            powershell -NonInteractive -NoProfile -ExecutionPolicy Bypass `
            -File "C:\Temp\wu_run_updates.ps1" 2>&1

        $updateExit = $LASTEXITCODE

        if ($updateExit -ne 0) {
            $status      = "Failed - Update Run"
            $description = "PsExec exit $updateExit | $($updateOutput -join ' ')"
            throw $description
        }

        # Parse update output for reporting
        if ($updateOutput -match "NO_UPDATES_NEEDED") {
            $status      = "Up To Date"
            $description = "No updates required"
        }
        elseif ($updateOutput -match "INSTALL_COMPLETE") {
            $countLine = ($updateOutput | Where-Object { $_ -match "UPDATES_FOUND:" }) -replace "UPDATES_FOUND:", ""
            $patches   = ($updateOutput | Where-Object { $_ -match "^PATCH:" }) -replace "PATCH:", ""
            $status      = "Updated"
            $description = "$moduleNote | $countLine update(s) installed: $($patches -join '; ')"
        }
        else {
            $status      = "Unknown"
            $description = "Unexpected output: $($updateOutput -join ' ')"
        }

        Write-Host "  [3/3] $status - $description" -ForegroundColor Green

    }
    catch {
        if ($status -eq "Success") {
            $status      = "Error"
            $description = $_.Exception.Message
        }
        Write-Host "  ERROR on $computer`: $description" -ForegroundColor Red
    }
    finally {
        # -----------------------------------------
        # Cleanup temp scripts from remote machine
        # -----------------------------------------
        foreach ($f in @($remoteModScript, $remoteUpdScript)) {
            if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
        }
    }

    # -----------------------------------------
    # Save Result
    # -----------------------------------------
    $results += [PSCustomObject]@{
        Hostname    = $computer
        Status      = $status
        Description = $description
        Timestamp   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
}

# -------------------------------------------------
# Export Results
# -------------------------------------------------
$results | Export-Csv "C:\Temp\pswindowsupdate_results.csv" -NoTypeInformation
Write-Host "`nDone. Results saved to C:\Temp\pswindowsupdate_results.csv" -ForegroundColor Cyan
$results | Format-Table -AutoSize