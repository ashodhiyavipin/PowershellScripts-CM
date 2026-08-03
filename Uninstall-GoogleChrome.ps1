<#
.SYNOPSIS
    Robust Google Chrome Enterprise Removal Script v3.7
    Designed for SCCM Task Sequence execution under SYSTEM context.

.DESCRIPTION
    Discovers and removes all Google Chrome Enterprise (MSI) installations
    (machine-wide and per-user), even when standard uninstall registry entries
    are missing. Optimized for environments where only Chrome Enterprise .msi
    packages are deployed.
    Phases: Initialization → Discovery → Process Kill → Uninstall → Cleanup → Validation.

.PARAMETER WhatIf
    Dry-run mode. Logs every action without modifying the system.

.PARAMETER LogPath
    Path to the log file. 
    Default: C:\Windows\fndr\logs\ChromeCleanup.log
    MSIUninstall Log File Path: C:\Windows\fndr\logs\GoogleChromeUninstall.log

.NOTES
    Exit Codes:
      0    — Chrome not found or successfully removed.
      1    — Removal attempted but chrome.exe still exists.
      3010 — Removal succeeded but reboot is recommended.

    Changelog:
      v3.0 — Initial release.
      v3.1 — Added exit code evaluation after msiexec invocations.
      v3.2 — Refactored for Chrome Enterprise MSI-only environments.
      v3.3 — Removed Win32_Product (WMI) entirely. Registry-based only.
      v3.4 — Bug fixes for Set-StrictMode compatibility.
      v3.5 — Simplified return types and collection handling.
      v3.6 — Removed Set-StrictMode entirely. Stripped all workarounds
              (Get-SafeProperty, @() wrappers, [string[]] casts).
              Cleaner, shorter, more maintainable code throughout.
      v3.7 — Added MSI uninstall log entries.
#>

[CmdletBinding(SupportsShouldProcess = $false)]
param(
    [switch]$WhatIf,
    [string]$LogPath = 'C:\Windows\fndr\logs\ChromeCleanup.log'
)

$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL STATE
# ─────────────────────────────────────────────────────────────────────────────
$script:ExitCode = 0
$script:RebootRequired = $false
$script:LoadedHives = @()
$script:FoundRegistryKeys = @()
$script:FoundChromeExePaths = @()
$script:AttemptedProductCodes = @()
$script:ScriptStartTime = Get-Date

# ─────────────────────────────────────────────────────────────────────────────
# P/INVOKE — MoveFileEx
# ─────────────────────────────────────────────────────────────────────────────
$MoveFileExSignature = @'
[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern bool MoveFileEx(
    string lpExistingFileName,
    string lpNewFileName,
    int dwFlags
);
'@
try { $null = [Win32.FileOps] } catch {
    Add-Type -MemberDefinition $MoveFileExSignature -Name 'FileOps' -Namespace 'Win32' -ErrorAction SilentlyContinue
}
$MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004

# ═════════════════════════════════════════════════════════════════════════════
#  FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════

$logDir = Split-Path -Path $LogPath -Parent
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $entry = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'))] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
    switch ($Level) {
        'ERROR' { Write-Warning $entry }
        'WARN' { Write-Warning $entry }
        default { Write-Verbose $entry -Verbose }
    }
}

function Invoke-ProcessWithTimeout {
    param([string]$FilePath, [string]$Arguments, [int]$TimeoutSeconds = 120)
    Write-Log "Launching: `"$FilePath`" $Arguments (timeout ${TimeoutSeconds}s)"
    if ($WhatIf.IsPresent) {
        Write-Log "[WhatIf] Would execute: `"$FilePath`" $Arguments"; return 0
    }
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $FilePath; $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $exited = $proc.WaitForExit($TimeoutSeconds * 1000)
        if (-not $exited) { Write-Log "Timeout — killing." -Level WARN; $proc.Kill(); return -1 }
        Write-Log "Process exited with code $($proc.ExitCode)."
        return $proc.ExitCode
    }
    catch { Write-Log "Failed to launch: $_" -Level ERROR; return -1 }
}

function Evaluate-InstallerExitCode {
    param([int]$ExitCode, [string]$InstallerDescription = 'msiexec')
    switch ($ExitCode) {
        0 { Write-Log "$InstallerDescription → success." }
        3010 { Write-Log "$InstallerDescription → success, reboot required." -Level WARN; $script:RebootRequired = $true }
        1641 { Write-Log "$InstallerDescription → reboot initiated." -Level WARN; $script:RebootRequired = $true }
        1605 { Write-Log "$InstallerDescription → product not installed (non-fatal)." -Level WARN }
        1603 { Write-Log "$InstallerDescription → fatal MSI error." -Level ERROR }
        -1 { Write-Log "$InstallerDescription → timeout or launch failure." -Level ERROR }
        default { Write-Log "$InstallerDescription → unexpected exit code: $ExitCode" -Level WARN }
    }
}

function Invoke-MsiUninstall {
    param([string]$ProductCode, [string]$Source = 'Unknown')
    $ProductCode = $ProductCode.Trim()
    if ([string]::IsNullOrWhiteSpace($ProductCode)) {
        Write-Log "Empty product code from '$Source'. Skipping." -Level WARN; return $false
    }
    if ($script:AttemptedProductCodes -contains $ProductCode) {
        Write-Log "Product code $ProductCode already attempted. Skipping from '$Source'."; return $false
    }
    $script:AttemptedProductCodes += $ProductCode
    $desc = "msiexec /x $ProductCode (source: $Source)"
    Write-Log "Attempting: $desc"
    $rc = Invoke-ProcessWithTimeout -FilePath 'msiexec.exe' -Arguments "/x $ProductCode /qn /norestart /L*V C:\Windows\fndr\logs\GoogleChromeUninstall.log" -TimeoutSeconds 120
    Evaluate-InstallerExitCode -ExitCode $rc -InstallerDescription $desc
    return $true
}

function Convert-CompressedGuidToStandard {
    param([string]$CompressedGuid)
    if ($CompressedGuid.Length -ne 32 -or $CompressedGuid -notmatch '^[A-Fa-f0-9]{32}$') {
        Write-Log "Invalid compressed GUID: $CompressedGuid" -Level WARN; return $null
    }
    try {
        $c = $CompressedGuid.ToCharArray()
        $b1 = ($c[7..0]) -join ''; $b2 = ($c[11..8]) -join ''; $b3 = ($c[15..12]) -join ''
        $b4 = ''; for ($i = 16; $i -lt 32; $i += 2) { $b4 += $c[$i + 1]; $b4 += $c[$i] }
        return "{$b1-$b2-$b3-$($b4.Substring(0,4))-$($b4.Substring(4,12))}".ToUpper()
    }
    catch { Write-Log "GUID conversion failed: $_" -Level WARN; return $null }
}

function Test-ChromeExeExists {
    $candidates = @(
        'C:\Program Files\Google\Chrome\Application\chrome.exe',
        'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    )
    $userDirs = Get-ChildItem -Path 'C:\Users' -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @('Public', 'Default', 'defaultuser0', 'All Users', 'Default User') }
    foreach ($ud in $userDirs) {
        $candidates += Join-Path $ud.FullName 'AppData\Local\Google\Chrome\Application\chrome.exe'
    }
    $found = @()
    foreach ($c in $candidates) {
        if (Test-Path -Path $c -PathType Leaf) { $found += $c }
    }
    return $found
}

function Remove-ItemSafe {
    param([string]$Path, [int]$MaxRetries = 3, [int]$DelaySeconds = 2)
    if (-not (Test-Path -Path $Path)) { Write-Log "Not found, skipping: $Path"; return $true }
    if ($WhatIf.IsPresent) { Write-Log "[WhatIf] Would remove: $Path"; return $true }

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-Log "Removing (attempt $attempt/$MaxRetries): $Path"
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            Write-Log "Removed: $Path"; return $true
        }
        catch {
            Write-Log "Attempt $attempt failed: $_" -Level WARN
            if ($attempt -lt $MaxRetries) { Start-Sleep -Seconds $DelaySeconds }
        }
    }
    Write-Log "Scheduling reboot deletion: $Path" -Level WARN
    try {
        if (Test-Path -Path $Path -PathType Container) {
            $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Sort-Object { $_.FullName.Length } -Descending
            foreach ($item in $items) {
                [Win32.FileOps]::MoveFileEx($item.FullName, [NullString]::Value, $MOVEFILE_DELAY_UNTIL_REBOOT) | Out-Null
            }
        }
        [Win32.FileOps]::MoveFileEx($Path, [NullString]::Value, $MOVEFILE_DELAY_UNTIL_REBOOT) | Out-Null
        $script:RebootRequired = $true
    }
    catch { Write-Log "Failed to schedule reboot deletion: $_" -Level ERROR }
    return $false
}

function Remove-RegistryKeySafe {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) { return }
    if ($WhatIf.IsPresent) { Write-Log "[WhatIf] Would remove registry key: $Path"; return }
    try {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Log "Removed registry key: $Path"
    }
    catch {
        if ("$_" -like '*does not exist*') { Write-Log "Registry key already absent: $Path" }
        else { Write-Log "Failed to remove registry key '$Path': $_" -Level WARN }
    }
}

function Unload-AllHives {
    foreach ($hiveName in $script:LoadedHives) {
        Write-Log "Unloading hive: HKU\$hiveName"
        if (-not $WhatIf.IsPresent) {
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 500
                $null = & reg.exe unload "HKU\$hiveName" 2>&1
                if ($LASTEXITCODE -eq 0) { Write-Log "Unloaded: HKU\$hiveName"; break }
                Write-Log "Unload attempt $attempt/3 failed." -Level WARN
                if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
                else { Write-Log "FAILED to unload HKU\$hiveName after 3 attempts." -Level ERROR }
            }
        }
        else { Write-Log "[WhatIf] Would unload: HKU\$hiveName" }
    }
}

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN EXECUTION
# ═════════════════════════════════════════════════════════════════════════════

Write-Log "═══════════════════════════════════════════════════════════════"
Write-Log "Google Chrome Enterprise Removal Script v3.6 — Starting"
Write-Log "═══════════════════════════════════════════════════════════════"
Write-Log "Hostname       : $($env:COMPUTERNAME)"
Write-Log "Executing User : $($env:USERNAME)"
Write-Log "User Domain    : $($env:USERDOMAIN)"
Write-Log "LogPath        : $LogPath"
Write-Log "WhatIf Mode    : $($WhatIf.IsPresent)"
if ($WhatIf.IsPresent) { Write-Log "*** DRY-RUN MODE — No changes will be made ***" }

# ─── PHASE 1: Discovery ────────────────────────────────────────────────────
Write-Log ""
Write-Log "─── PHASE 1: Discovery ───────────────────────────────────────"

if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null
}

# 1.1 Load offline user hives
Write-Log 'Loading offline user registry hives...'
$excludedProfiles = @('Public', 'Default', 'defaultuser0', 'All Users', 'Default User')
$userProfiles = Get-ChildItem -Path 'C:\Users' -Directory -ErrorAction SilentlyContinue |
Where-Object { $_.Name -notin $excludedProfiles }

$loadedSIDs = Get-ChildItem -Path 'HKU:\' -ErrorAction SilentlyContinue |
Select-Object -ExpandProperty PSChildName

foreach ($profile in $userProfiles) {
    $ntUserDat = Join-Path $profile.FullName 'NTUSER.DAT'
    if (-not (Test-Path -Path $ntUserDat)) {
        Write-Log "No NTUSER.DAT for $($profile.Name), skipping."
        continue
    }
    $hiveName = "TEMP_$($profile.Name)"
    $alreadyLoaded = $false

    if (Test-Path -Path "HKU:\$hiveName") {
        $alreadyLoaded = $true
        Write-Log "Hive already loaded: HKU\$hiveName"
    }
    else {
        foreach ($sid in $loadedSIDs) {
            try {
                $pp = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
                if ($pp -eq $profile.FullName) {
                    $alreadyLoaded = $true; $hiveName = $sid
                    Write-Log "Profile $($profile.Name) already loaded under SID $sid."
                    break
                }
            }
            catch { }
        }
    }

    if (-not $alreadyLoaded) {
        Write-Log "Loading hive: $($profile.Name) → HKU\$hiveName"
        if (-not $WhatIf.IsPresent) {
            $null = & reg.exe load "HKU\$hiveName" "$ntUserDat" 2>&1
            if ($LASTEXITCODE -ne 0) { Write-Log "Failed to load hive for $($profile.Name)" -Level WARN; continue }
            $script:LoadedHives += $hiveName
        }
        else {
            Write-Log "[WhatIf] Would load hive for $($profile.Name)."
            $script:LoadedHives += $hiveName
        }
    }
}

# 1.2 Registry scan
Write-Log 'Scanning registry for Chrome uninstall entries...'
$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$hkuChildren = Get-ChildItem -Path 'HKU:\' -ErrorAction SilentlyContinue |
Select-Object -ExpandProperty PSChildName
foreach ($sid in $hkuChildren) {
    $uninstallPaths += "HKU:\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
}

foreach ($regPath in $uninstallPaths) {
    try {
        $keys = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*Google Chrome*' }
        foreach ($key in $keys) {
            Write-Log "Found Chrome uninstall key: $($key.PSPath)"
            $productCode = $null

            $keyName = Split-Path -Path $key.PSPath -Leaf
            if ($keyName -match '^\{[A-Fa-f0-9\-]+\}$') { $productCode = $keyName }

            if (-not $productCode) {
                $cmd = if ($key.QuietUninstallString) { $key.QuietUninstallString } else { $key.UninstallString }
                if ($cmd -match '\{[A-Fa-f0-9\-]+\}') { $productCode = $Matches }
            }

            $script:FoundRegistryKeys += [PSCustomObject]@{
                Path                 = $key.PSPath
                DisplayName          = $key.DisplayName
                UninstallString      = $key.UninstallString
                QuietUninstallString = $key.QuietUninstallString
                ProductCode          = $productCode
            }
            if ($productCode) { Write-Log "  → Product code: $productCode" }
            else { Write-Log "  → Could not extract MSI product code." -Level WARN }
        }
    }
    catch {
        Write-Log "Error scanning '$regPath': $_" -Level WARN
    }
}
Write-Log "Registry scan complete. Found $($script:FoundRegistryKeys.Count) uninstall key(s)."

# 1.3 File system scan
Write-Log 'Scanning file system for chrome.exe...'
$script:FoundChromeExePaths = Test-ChromeExeExists
foreach ($p in $script:FoundChromeExePaths) { Write-Log "Found chrome.exe: $p" }
Write-Log "File system scan complete. Found $($script:FoundChromeExePaths.Count) chrome.exe instance(s)."

# 1.4 Decision gate
$hasRegistryKeys = $script:FoundRegistryKeys.Count -gt 0
$hasFiles = $script:FoundChromeExePaths.Count -gt 0

if (-not $hasRegistryKeys -and -not $hasFiles) {
    Write-Log 'Chrome not detected on this machine. Nothing to do.'
    Write-Log ''
    Unload-AllHives
    $duration = (New-TimeSpan -Start $script:ScriptStartTime -End (Get-Date)).TotalSeconds
    Write-Log ''
    Write-Log '═══════════════════════════════════════════════════════════════'
    Write-Log "Script completed in $([math]::Round($duration, 2)) seconds. Exit code: 0"
    Write-Log '═══════════════════════════════════════════════════════════════'
    exit 0
}

Write-Log "Decision gate → Registry keys: $hasRegistryKeys | Files on disk: $hasFiles"

# ─── PHASE 2: Process Termination ──────────────────────────────────────────
Write-Log ''
Write-Log '─── PHASE 2: Process Termination ─────────────────────────────'

foreach ($procName in @('chrome', 'GoogleUpdate', 'GoogleCrashHandler', 'GoogleCrashHandler64')) {
    $running = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($running) {
        Write-Log "Terminating: $procName (PID: $($running.Id -join ', '))"
        if (-not $WhatIf.IsPresent) { $running | Stop-Process -Force -ErrorAction SilentlyContinue }
        else { Write-Log "[WhatIf] Would terminate $procName." }
    }
    else { Write-Log "Not running: $procName" }
}
if (-not $WhatIf.IsPresent) { Write-Log 'Waiting 2s for handles to release...'; Start-Sleep -Seconds 2 }

# ─── PHASE 3: Uninstallation ───────────────────────────────────────────────
Write-Log ''
Write-Log '─── PHASE 3: Uninstallation (MSI-Only Strategy) ──────────────'
$chromeRemoved = $false

# Step 3A: Registry-discovered product codes
if ($hasRegistryKeys) {
    Write-Log '── Step 3A: MSI uninstall via registry product codes ──'
    $codesFromRegistry = $script:FoundRegistryKeys | Where-Object { $_.ProductCode }

    if (-not $codesFromRegistry) {
        Write-Log 'No product codes extracted from registry. Skipping 3A.' -Level WARN
    }
    else {
        foreach ($rk in $codesFromRegistry) {
            Invoke-MsiUninstall -ProductCode $rk.ProductCode -Source "Registry: $($rk.Path)"
        }
        $remaining = Test-ChromeExeExists
        if ($remaining.Count -eq 0) { Write-Log '3A succeeded.'; $chromeRemoved = $true }
        else { Write-Log "3A incomplete. Remaining: $($remaining -join ', ')" -Level WARN }
    }
}

# Step 3B: Windows Installer registry discovery
if (-not $chromeRemoved) {
    Write-Log '── Step 3B: MSI uninstall via Installer registry ──'
    $msiCodes = @()

    # Source 1: Installer\Products
    $installerProductsPath = 'HKLM:\SOFTWARE\Classes\Installer\Products'
    if (Test-Path -Path $installerProductsPath) {
        try {
            foreach ($subKey in (Get-ChildItem -Path $installerProductsPath -ErrorAction SilentlyContinue)) {
                try {
                    $props = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
                    if ($props.ProductName -like '*Google Chrome*') {
                        $cg = Split-Path -Path $subKey.PSPath -Leaf
                        $sg = Convert-CompressedGuidToStandard -CompressedGuid $cg
                        if ($sg) {
                            Write-Log "Found in Installer\Products: '$($props.ProductName)' → $sg"
                            $msiCodes += [PSCustomObject]@{ ProductCode = $sg; Source = "Installer\Products\$cg" }
                        }
                    }
                }
                catch { }
            }
        }
        catch { Write-Log "Error scanning Installer\Products: $_" -Level WARN }
    }

    # Source 2: Installer\UserData
    $installerUserDataPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
    if (Test-Path -Path $installerUserDataPath) {
        try {
            foreach ($sidKey in (Get-ChildItem -Path $installerUserDataPath -ErrorAction SilentlyContinue)) {
                $prodPath = Join-Path $sidKey.PSPath 'Products'
                if (-not (Test-Path $prodPath)) { continue }
                foreach ($pk in (Get-ChildItem -Path $prodPath -ErrorAction SilentlyContinue)) {
                    $ipPath = Join-Path $pk.PSPath 'InstallProperties'
                    if (-not (Test-Path $ipPath)) { continue }
                    try {
                        $ip = Get-ItemProperty -Path $ipPath -ErrorAction SilentlyContinue
                        if ($ip.DisplayName -like '*Google Chrome*') {
                            $cg = Split-Path -Path $pk.PSPath -Leaf
                            $sg = Convert-CompressedGuidToStandard -CompressedGuid $cg
                            if ($sg) {
                                Write-Log "Found in UserData: '$($ip.DisplayName)' v$($ip.DisplayVersion) → $sg"
                                $msiCodes += [PSCustomObject]@{ ProductCode = $sg; Source = "UserData\$cg" }
                            }
                        }
                    }
                    catch { }
                }
            }
        }
        catch { Write-Log "Error scanning Installer\UserData: $_" -Level WARN }
    }

    if ($msiCodes.Count -gt 0) {
        Write-Log "Discovered $($msiCodes.Count) MSI code(s)."
        foreach ($entry in $msiCodes) {
            Invoke-MsiUninstall -ProductCode $entry.ProductCode -Source "MSI Registry ($($entry.Source))"
        }
        $remaining = Test-ChromeExeExists
        if ($remaining.Count -eq 0) { Write-Log '3B succeeded.'; $chromeRemoved = $true }
        else { Write-Log "3B incomplete. Remaining: $($remaining -join ', ')" -Level WARN }
    }
    else { Write-Log 'No Chrome MSI codes found in Installer registry.' }
}

# Step 3C: Sledgehammer
if (-not $chromeRemoved) {
    Write-Log '── Step 3C: Manual deletion (sledgehammer) ──'

    $deletionTargets = @(
        'C:\Program Files\Google',
        'C:\Program Files (x86)\Google'
    )
    foreach ($profile in $userProfiles) {
        $ugd = Join-Path $profile.FullName 'AppData\Local\Google'
        if (Test-Path $ugd) { $deletionTargets += $ugd }
    }
    # Add parent dirs from discovered chrome.exe paths
    foreach ($cp in $script:FoundChromeExePaths) {
        if (-not $cp) { continue }
        try {
            # chrome.exe → Application → Chrome → Google
            $googleDir = Split-Path (Split-Path (Split-Path $cp -Parent) -Parent) -Parent
            if ($googleDir -and (Test-Path $googleDir) -and ($deletionTargets -notcontains $googleDir)) {
                Write-Log "Adding discovered path to targets: $googleDir"
                $deletionTargets += $googleDir
            }
        }
        catch { }
    }

    foreach ($target in $deletionTargets) {
        $null = Remove-ItemSafe -Path $target -MaxRetries 3 -DelaySeconds 2
    }

    # Registry cleanup
    Write-Log 'Removing Google registry keys...'
    Remove-RegistryKeySafe -Path 'HKLM:\SOFTWARE\Google'
    Remove-RegistryKeySafe -Path 'HKLM:\SOFTWARE\WOW6432Node\Google'
    foreach ($sid in $hkuChildren) { Remove-RegistryKeySafe -Path "HKU:\$sid\SOFTWARE\Google" }

    # MSI registration cleanup
    Write-Log 'Cleaning MSI installer registration...'
    if (Test-Path $installerProductsPath) {
        try {
            foreach ($sk in (Get-ChildItem $installerProductsPath -ErrorAction SilentlyContinue)) {
                try {
                    $pr = Get-ItemProperty $sk.PSPath -ErrorAction SilentlyContinue
                    if ($pr.ProductName -like '*Google Chrome*') { Remove-RegistryKeySafe -Path $sk.PSPath }
                }
                catch { }
            }
        }
        catch { Write-Log "Error cleaning Installer\Products: $_" -Level WARN }
    }
    if (Test-Path $installerUserDataPath) {
        try {
            foreach ($sk in (Get-ChildItem $installerUserDataPath -ErrorAction SilentlyContinue)) {
                $pp = Join-Path $sk.PSPath 'Products'
                if (-not (Test-Path $pp)) { continue }
                foreach ($pk in (Get-ChildItem $pp -ErrorAction SilentlyContinue)) {
                    $ipp = Join-Path $pk.PSPath 'InstallProperties'
                    if (-not (Test-Path $ipp)) { continue }
                    try {
                        $ip = Get-ItemProperty $ipp -ErrorAction SilentlyContinue
                        if ($ip.DisplayName -like '*Google Chrome*') { Remove-RegistryKeySafe -Path $pk.PSPath }
                    }
                    catch { }
                }
            }
        }
        catch { Write-Log "Error cleaning Installer\UserData: $_" -Level WARN }
    }

    $remaining = Test-ChromeExeExists
    if ($remaining.Count -eq 0) { Write-Log '3C succeeded.'; $chromeRemoved = $true }
    else {
        Write-Log "3C: chrome.exe still present at $($remaining.Count) location(s):" -Level ERROR
        foreach ($r in $remaining) { Write-Log "  → $r" -Level ERROR }
    }
}

# ─── PHASE 4: Residual Cleanup ─────────────────────────────────────────────
Write-Log ''
Write-Log '─── PHASE 4: Residual Cleanup ────────────────────────────────'

# Services
foreach ($svcName in @('gupdate', 'gupdatem')) {
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($svc) {
        Write-Log "Removing service: $svcName"
        if (-not $WhatIf.IsPresent) {
            Stop-Service $svcName -Force -ErrorAction SilentlyContinue
            $null = & sc.exe delete $svcName 2>&1
            if ($LASTEXITCODE -eq 0) { Write-Log "Service removed: $svcName" }
            else { Write-Log "Failed to remove service: $svcName" -Level WARN }
        }
        else { Write-Log "[WhatIf] Would remove service: $svcName" }
    }
    else { Write-Log "Service not found: $svcName" }
}

# Scheduled tasks
Write-Log 'Removing Google scheduled tasks...'
$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
Where-Object { $_.TaskName -like 'GoogleUpdateTaskMachine*' -or $_.TaskName -like 'GoogleUpdateTaskUser*' -or $_.TaskPath -like '\Google\*' }
foreach ($task in $tasks) {
    Write-Log "Removing task: $($task.TaskPath)$($task.TaskName)"
    if (-not $WhatIf.IsPresent) {
        Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Task removed: $($task.TaskName)"
    }
    else { Write-Log "[WhatIf] Would remove task: $($task.TaskName)" }
}

# Shortcuts
Write-Log 'Removing Chrome shortcuts...'
$shortcuts = @(
    'C:\Users\Public\Desktop\Google Chrome.lnk',
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk'
)
foreach ($profile in $userProfiles) {
    $shortcuts += Join-Path $profile.FullName 'Desktop\Google Chrome.lnk'
    $shortcuts += Join-Path $profile.FullName 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk'
}
foreach ($lnk in $shortcuts) {
    if (Test-Path $lnk) {
        Write-Log "Removing: $lnk"
        if (-not $WhatIf.IsPresent) { Remove-Item $lnk -Force -ErrorAction SilentlyContinue }
        else { Write-Log "[WhatIf] Would remove: $lnk" }
    }
}

# Leftover registry
Write-Log 'Final registry cleanup...'
Remove-RegistryKeySafe -Path 'HKLM:\SOFTWARE\Google'
Remove-RegistryKeySafe -Path 'HKLM:\SOFTWARE\WOW6432Node\Google'
foreach ($sid in $hkuChildren) { Remove-RegistryKeySafe -Path "HKU:\$sid\SOFTWARE\Google" }

# ─── PHASE 5: Final Validation ─────────────────────────────────────────────
Write-Log ''
Write-Log '─── PHASE 5: Final Validation ────────────────────────────────'

$finalCheck = Test-ChromeExeExists

if ($finalCheck.Count -eq 0) {
    if ($script:RebootRequired) {
        $script:ExitCode = 3010
        Write-Log 'SUCCESS — Chrome removed. Reboot recommended.'
    }
    else {
        $script:ExitCode = 0
        Write-Log 'SUCCESS — Chrome fully removed.'
    }
}
else {
    $script:ExitCode = 1
    Write-Log "FAILED — chrome.exe still exists at $($finalCheck.Count) location(s):" -Level ERROR
    foreach ($p in $finalCheck) { Write-Log "  → $p" -Level ERROR }
}

Write-Log ''
Unload-AllHives

$duration = (New-TimeSpan -Start $script:ScriptStartTime -End (Get-Date)).TotalSeconds
Write-Log ''
Write-Log '═══════════════════════════════════════════════════════════════'
Write-Log "Script completed in $([math]::Round($duration, 2)) seconds. Exit code: $($script:ExitCode)"
Write-Log '═══════════════════════════════════════════════════════════════'

exit $script:ExitCode