<#
.SYNOPSIS
    Removes all Microsoft .NET 5, 6, 7, 8, 9, and 10 products from Windows workstations.

.DESCRIPTION
    This script performs a silent, fully automated removal of all .NET 5 through 10 components
    including Runtimes, ASP.NET Core, Desktop Runtimes, Hosting Bundles, SDKs, Targeting Packs,
    and AppHost Packs. It uses registry-based discovery (no Win32_Product) and handles both
    MSI and EXE-based uninstallers. Designed for deployment via SCCM Task Sequence in SYSTEM context.

    The script will skip execution entirely if Visual Studio is detected on the machine to prevent
    breaking VS-bundled .NET components.

    The dotnet.exe host is intentionally left intact for the latest .NET version installer to handle.

    .NET Framework 4.x and .NET Core 2.x/3.x are explicitly excluded from scope.

.NOTES
    Exit Codes:
        0    = All targeted products removed successfully (or none found)
        1    = One or more uninstalls failed
        2    = Visual Studio detected -- machine skipped, no action taken
        3010 = Removal successful but a reboot is required

    Changelog:
        2026-01-XX  v1.0  Initial release -- registry-based discovery, MSI+EXE silent uninstall,
                          CMTrace-compatible logging, Visual Studio detection gate.
        2026-01-XX  v1.1  Bugfix -- $matches index references corrected in Invoke-Uninstall,
                          Invoke-ExeUninstall, and version parsing in Get-DotNetProducts.
        2026-07-02  v1.2  Bugfix -- CMTrace timestamp timezone offset fixed (+-420 -> +420),
                          em dash characters replaced with standard double-hyphen.
#>

# ============================================================================
# CONFIGURATION
# ============================================================================

$Script:LogPath = "C:\Windows\fndr\logs"
$Script:LogFile = Join-Path -Path $Script:LogPath -ChildPath "DotNetCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Script:Timeout = 120  # seconds per uninstall
$Script:TargetMajorVersions = @(5, 6, 7, 8, 9, 10)
$Script:RebootRequired = $false
$Script:FailureOccurred = $false

# ============================================================================
# FUNCTIONS
# ============================================================================

function Write-CMLog {
    <#
    .SYNOPSIS
        Writes a CMTrace-compatible log entry.
    .PARAMETER Message
        The log message.
    .PARAMETER Type
        1 = Informational, 2 = Warning, 3 = Error
    .PARAMETER Component
        The component name for CMTrace grouping.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet(1, 2, 3)]
        [int]$Type = 1,

        [Parameter(Mandatory = $false)]
        [string]$Component = "DotNetCleanup"
    )

    # Ensure log directory exists
    if (-not (Test-Path -Path $Script:LogPath)) {
        try {
            New-Item -Path $Script:LogPath -ItemType Directory -Force | Out-Null
        }
        catch {
            Write-Warning "Failed to create log directory: $($_.Exception.Message)"
            return
        }
    }

    # CMTrace-compatible timestamp format
    $Now = Get-Date
    $TimeZoneBias = [int][System.TimeZoneInfo]::Local.GetUtcOffset($Now).TotalMinutes
    $TimeGenerated = "$($Now.ToString('HH:mm:ss.fff'))+$([Math]::Abs($TimeZoneBias))"
    $DateGenerated = $Now.ToString('MM-dd-yyyy')
    $Thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    $Context = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Source = $Script:LogFile

    $LogLine = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="{7}">' -f `
        $Message, $TimeGenerated, $DateGenerated, $Component, $Context, $Type, $Thread, $Source

    try {
        Add-Content -Path $Script:LogFile -Value $LogLine -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to log: $($_.Exception.Message)"
    }
}

function Test-AdminContext {
    <#
    .SYNOPSIS
        Validates the script is running with administrative privileges.
    #>
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-VisualStudioInstalled {
    <#
    .SYNOPSIS
        Checks if any edition of Visual Studio is installed on the machine.
    .DESCRIPTION
        Scans the uninstall registry keys for products matching Visual Studio.
        Also checks the VS Setup instances registry path.
    #>

    Write-CMLog -Message "Checking for Visual Studio installations..." -Component "VSDetection"

    # Method 1: Check VS Setup Instances
    $vsSetupPath = "HKLM:\SOFTWARE\Microsoft\VisualStudio\Setup"
    if (Test-Path -Path $vsSetupPath) {
        Write-CMLog -Message "Visual Studio Setup registry path found: $vsSetupPath" -Type 2 -Component "VSDetection"
        return $true
    }

    # Method 2: Check uninstall registry for Visual Studio entries
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $uninstallPaths) {
        $vsProducts = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -like "Microsoft Visual Studio 20*" -and
            $_.Publisher -eq "Microsoft Corporation"
        }

        if ($vsProducts) {
            foreach ($vs in $vsProducts) {
                Write-CMLog -Message "Visual Studio detected: $($vs.DisplayName) (Version: $($vs.DisplayVersion))" -Type 2 -Component "VSDetection"
            }
            return $true
        }
    }

    Write-CMLog -Message "No Visual Studio installation detected. Proceeding with removal." -Component "VSDetection"
    return $false
}

function Get-DotNetProducts {
    <#
    .SYNOPSIS
        Discovers all .NET 5-10 products from the registry.
    .DESCRIPTION
        Scans both 64-bit and 32-bit uninstall registry keys, filters by known .NET product
        name patterns, publisher, and target major versions. Returns an array of product objects
        sorted by removal priority.
    #>

    Write-CMLog -Message "Starting .NET product discovery via registry scan..." -Component "Discovery"

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    # Name patterns to match
    $namePatterns = @(
        "Microsoft .NET SDK*",
        "Microsoft .NET Runtime*",
        "Microsoft .NET Core Runtime*",
        "Microsoft ASP.NET Core*",
        "Microsoft .NET AppHost Pack*",
        "Microsoft .NET Host*",
        "Microsoft .NET Core Host*",
        "Microsoft .NET Targeting Pack*",
        "Microsoft .NET Windows Desktop Runtime*",
        "Microsoft Windows Desktop Runtime*",
        "Microsoft .NET Core*Templates*",
        "Microsoft .NET Shared Framework*",
        "Microsoft ASP.NET Core*Module*",
        "Microsoft .NET Core*SDK*",
        "Microsoft .NET*Hosting*",
        "Microsoft .NET Runtime*Desktop*",
        "Microsoft .NET*workload*",
        "Microsoft .NET SDK Workload*"
    )

    $allProducts = @()
    $seenGUIDs = @{}

    foreach ($regPath in $uninstallPaths) {
        Write-CMLog -Message "Scanning registry path: $regPath" -Component "Discovery"

        $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue

        if (-not $entries) {
            Write-CMLog -Message "No entries found at: $regPath" -Type 2 -Component "Discovery"
            continue
        }

        foreach ($entry in $entries) {
            # Skip entries without a display name
            if ([string]::IsNullOrWhiteSpace($entry.DisplayName)) { continue }

            # Check publisher
            if ($entry.Publisher -ne "Microsoft Corporation") { continue }

            # Check if display name matches any of our patterns
            $matched = $false
            foreach ($pattern in $namePatterns) {
                if ($entry.DisplayName -like $pattern) {
                    $matched = $true
                    break
                }
            }
            if (-not $matched) { continue }

            # Parse major version from DisplayVersion
            $majorVersion = $null
            if (-not [string]::IsNullOrWhiteSpace($entry.DisplayVersion)) {
                $versionParts = $entry.DisplayVersion -split '\.'
                if ($versionParts.Count -ge 1) {
                    [int]$parsedMajor = 0
                    if ([int]::TryParse($versionParts[0], [ref]$parsedMajor)) {
                        $majorVersion = $parsedMajor
                    }
                }
            }

            # If we couldn't parse version from DisplayVersion, try from DisplayName
            if ($null -eq $majorVersion) {
                foreach ($ver in $Script:TargetMajorVersions) {
                    if ($entry.DisplayName -match "\b$ver\.\d") {
                        $majorVersion = $ver
                        break
                    }
                }
            }

            # Skip if not in our target versions
            if ($null -eq $majorVersion -or $majorVersion -notin $Script:TargetMajorVersions) { continue }

            # Extract the registry key name (product GUID)
            $keyName = $entry.PSChildName

            # Skip duplicates (same product in both 64-bit and WOW6432Node)
            if ($seenGUIDs.ContainsKey($keyName)) { continue }
            $seenGUIDs[$keyName] = $true

            # Determine uninstall type and command
            $uninstallType = "Unknown"
            $uninstallCommand = $null

            $quietUninstall = $entry.QuietUninstallString
            $standardUninstall = $entry.UninstallString

            if (-not [string]::IsNullOrWhiteSpace($quietUninstall)) {
                if ($quietUninstall -match "msiexec" -or $quietUninstall -match "MsiExec") {
                    $uninstallType = "MSI"
                }
                else {
                    $uninstallType = "EXE"
                }
                $uninstallCommand = $quietUninstall
            }
            elseif (-not [string]::IsNullOrWhiteSpace($standardUninstall)) {
                if ($standardUninstall -match "msiexec" -or $standardUninstall -match "MsiExec") {
                    $uninstallType = "MSI"
                }
                else {
                    $uninstallType = "EXE"
                }
                $uninstallCommand = $standardUninstall
            }

            # Determine removal priority
            $priority = Get-RemovalPriority -DisplayName $entry.DisplayName

            $product = [PSCustomObject]@{
                DisplayName          = $entry.DisplayName
                DisplayVersion       = $entry.DisplayVersion
                MajorVersion         = $majorVersion
                Publisher            = $entry.Publisher
                ProductGUID          = $keyName
                UninstallType        = $uninstallType
                UninstallCommand     = $uninstallCommand
                UninstallString      = $standardUninstall
                QuietUninstallString = $quietUninstall
                Priority             = $priority
                RegistryPath         = $regPath -replace '\\\*$', "\$keyName"
            }

            $allProducts += $product
        }
    }

    # Sort by priority descending (highest priority removed first)
    $allProducts = $allProducts | Sort-Object -Property Priority -Descending

    Write-CMLog -Message "Discovery complete. Found $($allProducts.Count) .NET product(s) matching criteria." -Component "Discovery"

    foreach ($p in $allProducts) {
        Write-CMLog -Message "  FOUND: $($p.DisplayName) | Version: $($p.DisplayVersion) | Type: $($p.UninstallType) | GUID: $($p.ProductGUID) | Priority: $($p.Priority)" -Component "Discovery"
    }

    return $allProducts
}

function Get-RemovalPriority {
    <#
    .SYNOPSIS
        Assigns a removal priority based on product type.
    .DESCRIPTION
        Higher number = removed first. SDKs and high-level components are removed before base runtimes.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $name = $DisplayName.ToLower()

    if ($name -like "*sdk*workload*" -or $name -like "*workload*") { return 90 }
    if ($name -like "*sdk*") { return 100 }
    if ($name -like "*template*") { return 90 }
    if ($name -like "*hosting*") { return 80 }
    if ($name -like "*targeting pack*") { return 70 }
    if ($name -like "*apphost*") { return 70 }
    if ($name -like "*asp.net*" -or $name -like "*aspnetcore*") { return 60 }
    if ($name -like "*desktop*") { return 50 }
    if ($name -like "*runtime*") { return 40 }
    if ($name -like "*host*") { return 30 }

    return 10
}

function Invoke-Uninstall {
    <#
    .SYNOPSIS
        Executes the uninstall for a single .NET product.
    .DESCRIPTION
        Handles both MSI-based and EXE-based uninstallers with timeout enforcement
        and exit code capture.
    .PARAMETER Product
        The product object from Get-DotNetProducts.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Product
    )

    $displayName = $Product.DisplayName
    $guid = $Product.ProductGUID

    Write-CMLog -Message "--- Starting uninstall: $displayName (Version: $($Product.DisplayVersion))" -Component "Uninstall"

    if ($Product.UninstallType -eq "MSI") {
        $msiGUID = $null

        if ($guid -match '^\{[0-9A-Fa-f\-]+\}$') {
            $msiGUID = $guid
        }
        elseif ($Product.UninstallString -match '(\{[0-9A-Fa-f\-]+\})') {
            $msiGUID = $matches[1]
        }
        elseif ($Product.QuietUninstallString -match '(\{[0-9A-Fa-f\-]+\})') {
            $msiGUID = $matches[1]
        }

        if ($msiGUID) {
            Write-CMLog -Message "  Using MSI uninstall: msiexec.exe /x $msiGUID /qn /norestart" -Component "Uninstall"

            try {
                $process = Start-Process -FilePath "msiexec.exe" `
                    -ArgumentList "/x $msiGUID /qn /norestart" `
                    -Wait:$false `
                    -PassThru `
                    -NoNewWindow `
                    -ErrorAction Stop

                $completed = $process.WaitForExit($Script:Timeout * 1000)

                if (-not $completed) {
                    Write-CMLog -Message "  TIMEOUT: Uninstall exceeded $($Script:Timeout) seconds. Killing process." -Type 3 -Component "Uninstall"
                    $process | Stop-Process -Force -ErrorAction SilentlyContinue
                    $Script:FailureOccurred = $true
                    return
                }

                $exitCode = $process.ExitCode
                Write-CMLog -Message "  Exit code: $exitCode" -Component "Uninstall"
                Resolve-ExitCode -ExitCode $exitCode -DisplayName $displayName
            }
            catch {
                Write-CMLog -Message "  ERROR: Failed to execute MSI uninstall for $displayName -- $($_.Exception.Message)" -Type 3 -Component "Uninstall"
                $Script:FailureOccurred = $true
            }
        }
        else {
            Write-CMLog -Message "  WARNING: MSI product detected but could not extract GUID. Attempting fallback..." -Type 2 -Component "Uninstall"
            Invoke-FallbackUninstall -Product $Product
        }
    }
    elseif ($Product.UninstallType -eq "EXE") {
        Invoke-ExeUninstall -Product $Product
    }
    else {
        Write-CMLog -Message "  SKIPPED: Unable to determine uninstall method for $displayName. Manual review required." -Type 2 -Component "Uninstall"
        Write-CMLog -Message "    UninstallString: $($Product.UninstallString)" -Type 2 -Component "Uninstall"
        Write-CMLog -Message "    QuietUninstallString: $($Product.QuietUninstallString)" -Type 2 -Component "Uninstall"
        $Script:FailureOccurred = $true
    }
}

function Invoke-ExeUninstall {
    <#
    .SYNOPSIS
        Handles EXE-based uninstallers (e.g., Hosting Bundles).
    #>
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Product
    )

    $displayName = $Product.DisplayName
    $command = $Product.QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($command)) {
        $command = $Product.UninstallString
    }

    if ([string]::IsNullOrWhiteSpace($command)) {
        Write-CMLog -Message "  SKIPPED: No uninstall command available for $displayName" -Type 2 -Component "Uninstall"
        $Script:FailureOccurred = $true
        return
    }

    # Parse the executable path and arguments
    $exePath = $null
    $exeArgs = ""

    if ($command -match '^"([^"]+)"\s*(.*)$') {
        $exePath = $matches[1]
        $exeArgs = $matches[2].Trim()
    }
    elseif ($command -match '^(\S+\.exe)\s*(.*)$') {
        $exePath = $matches[1]
        $exeArgs = $matches[2].Trim()
    }
    else {
        Write-CMLog -Message "  SKIPPED: Could not parse EXE uninstall command: $command" -Type 2 -Component "Uninstall"
        $Script:FailureOccurred = $true
        return
    }

    # Verify the executable exists
    if (-not (Test-Path -Path $exePath)) {
        Write-CMLog -Message "  SKIPPED: Uninstall executable not found at: $exePath" -Type 2 -Component "Uninstall"
        $Script:FailureOccurred = $true
        return
    }

    # Ensure silent and no-restart flags are present
    if ($exeArgs -notmatch "/quiet" -and $exeArgs -notmatch "--quiet") {
        $exeArgs = "$exeArgs /quiet"
    }
    if ($exeArgs -notmatch "/norestart" -and $exeArgs -notmatch "--norestart") {
        $exeArgs = "$exeArgs /norestart"
    }

    # Ensure /uninstall flag is present for hosting bundles and EXE installers
    if ($exeArgs -notmatch "/uninstall" -and $exeArgs -notmatch "--uninstall") {
        $exeArgs = "/uninstall $exeArgs"
    }

    $exeArgs = $exeArgs.Trim()

    Write-CMLog -Message "  Using EXE uninstall: `"$exePath`" $exeArgs" -Component "Uninstall"

    try {
        $process = Start-Process -FilePath $exePath `
            -ArgumentList $exeArgs `
            -Wait:$false `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        $completed = $process.WaitForExit($Script:Timeout * 1000)

        if (-not $completed) {
            Write-CMLog -Message "  TIMEOUT: Uninstall exceeded $($Script:Timeout) seconds. Killing process." -Type 3 -Component "Uninstall"
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            $Script:FailureOccurred = $true
            return
        }

        $exitCode = $process.ExitCode
        Write-CMLog -Message "  Exit code: $exitCode" -Component "Uninstall"
        Resolve-ExitCode -ExitCode $exitCode -DisplayName $displayName
    }
    catch {
        Write-CMLog -Message "  ERROR: Failed to execute EXE uninstall for $displayName -- $($_.Exception.Message)" -Type 3 -Component "Uninstall"
        $Script:FailureOccurred = $true
    }
}

function Invoke-FallbackUninstall {
    <#
    .SYNOPSIS
        Attempts uninstall using the raw UninstallString when GUID extraction fails.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Product
    )

    $displayName = $Product.DisplayName

    $command = $Product.QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($command)) {
        $command = $Product.UninstallString
    }

    if ([string]::IsNullOrWhiteSpace($command)) {
        Write-CMLog -Message "  SKIPPED: No uninstall command available for fallback on $displayName" -Type 2 -Component "Uninstall"
        $Script:FailureOccurred = $true
        return
    }

    # If it looks like an EXE command, delegate to EXE handler
    if ($command -notmatch "msiexec") {
        Write-CMLog -Message "  Fallback: Treating as EXE-based uninstaller." -Component "Uninstall"
        Invoke-ExeUninstall -Product $Product
        return
    }

    # It's msiexec-based but we couldn't extract a GUID -- try running the command directly
    if ($command -notmatch "/qn") {
        $command = $command -replace "/I", "/x"
        $command = "$command /qn /norestart"
    }

    Write-CMLog -Message "  Fallback MSI command: $command" -Component "Uninstall"

    try {
        $process = Start-Process -FilePath "cmd.exe" `
            -ArgumentList "/c $command" `
            -Wait:$false `
            -PassThru `
            -NoNewWindow `
            -ErrorAction Stop

        $completed = $process.WaitForExit($Script:Timeout * 1000)

        if (-not $completed) {
            Write-CMLog -Message "  TIMEOUT: Fallback uninstall exceeded $($Script:Timeout) seconds. Killing process." -Type 3 -Component "Uninstall"
            $process | Stop-Process -Force -ErrorAction SilentlyContinue
            $Script:FailureOccurred = $true
            return
        }

        $exitCode = $process.ExitCode
        Write-CMLog -Message "  Exit code: $exitCode" -Component "Uninstall"
        Resolve-ExitCode -ExitCode $exitCode -DisplayName $displayName
    }
    catch {
        Write-CMLog -Message "  ERROR: Fallback uninstall failed for $displayName -- $($_.Exception.Message)" -Type 3 -Component "Uninstall"
        $Script:FailureOccurred = $true
    }
}

function Resolve-ExitCode {
    <#
    .SYNOPSIS
        Evaluates an uninstall exit code and logs the result.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    switch ($ExitCode) {
        0 {
            Write-CMLog -Message "  SUCCESS: $DisplayName removed successfully." -Component "Uninstall"
        }
        3010 {
            Write-CMLog -Message "  SUCCESS: $DisplayName removed successfully. Reboot required." -Type 2 -Component "Uninstall"
            $Script:RebootRequired = $true
        }
        1605 {
            Write-CMLog -Message "  INFO: $DisplayName -- product not found (already removed or not installed). Exit code 1605." -Type 2 -Component "Uninstall"
        }
        1614 {
            Write-CMLog -Message "  INFO: $DisplayName -- product not found (invalid GUID). Exit code 1614." -Type 2 -Component "Uninstall"
        }
        default {
            Write-CMLog -Message "  FAILED: $DisplayName -- uninstall returned exit code $ExitCode" -Type 3 -Component "Uninstall"
            $Script:FailureOccurred = $true
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# --- Step 1: Initialize logging ---
Write-CMLog -Message "============================================================" -Component "Main"
Write-CMLog -Message "  .NET Cleanup Script -- v1.2" -Component "Main"
Write-CMLog -Message "  Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Component "Main"
Write-CMLog -Message "  Computer: $env:COMPUTERNAME" -Component "Main"
Write-CMLog -Message "  OS: $((Get-CimInstance -ClassName Win32_OperatingSystem).Caption)" -Component "Main"
Write-CMLog -Message "  OS Build: $([System.Environment]::OSVersion.Version)" -Component "Main"
Write-CMLog -Message "  Target Versions: $($Script:TargetMajorVersions -join ', ')" -Component "Main"
Write-CMLog -Message "  Timeout per product: $($Script:Timeout) seconds" -Component "Main"
Write-CMLog -Message "============================================================" -Component "Main"

# --- Step 2: Validate admin context ---
if (-not (Test-AdminContext)) {
    Write-CMLog -Message "ABORT: Script is not running with administrative privileges. Exiting." -Type 3 -Component "Main"
    exit 1
}
Write-CMLog -Message "Administrative context confirmed." -Component "Main"

# --- Step 3: Check for Visual Studio ---
if (Test-VisualStudioInstalled) {
    Write-CMLog -Message "============================================================" -Component "Main"
    Write-CMLog -Message "  SKIPPING MACHINE: Visual Studio is installed." -Type 2 -Component "Main"
    Write-CMLog -Message "  Reason: Removing .NET components could break Visual Studio." -Type 2 -Component "Main"
    Write-CMLog -Message "  Action: This machine requires manual handling or a separate process." -Type 2 -Component "Main"
    Write-CMLog -Message "============================================================" -Component "Main"
    exit 2
}

# --- Step 4: Discover .NET products ---
$products = Get-DotNetProducts

if (-not $products -or $products.Count -eq 0) {
    Write-CMLog -Message "No .NET 5-10 products found on this machine. Nothing to remove." -Component "Main"
    Write-CMLog -Message "Script completed successfully. Exiting with code 0." -Component "Main"
    exit 0
}

Write-CMLog -Message "============================================================" -Component "Main"
Write-CMLog -Message "  Total products to remove: $($products.Count)" -Component "Main"
Write-CMLog -Message "============================================================" -Component "Main"

# --- Step 5: Uninstall loop ---
$counter = 0
$total = $products.Count

foreach ($product in $products) {
    $counter++
    Write-CMLog -Message "[$counter / $total] Processing: $($product.DisplayName)" -Component "Main"
    Invoke-Uninstall -Product $product
    Write-CMLog -Message "[$counter / $total] Completed: $($product.DisplayName)" -Component "Main"
    Write-CMLog -Message "------------------------------------------------------------" -Component "Main"
}

# --- Step 6: Verification scan ---
Write-CMLog -Message "============================================================" -Component "Verification"
Write-CMLog -Message "  Running post-removal verification scan..." -Component "Verification"
Write-CMLog -Message "============================================================" -Component "Verification"

$remainingProducts = Get-DotNetProducts

if ($remainingProducts -and $remainingProducts.Count -gt 0) {
    Write-CMLog -Message "WARNING: $($remainingProducts.Count) product(s) still detected after removal:" -Type 2 -Component "Verification"
    foreach ($r in $remainingProducts) {
        Write-CMLog -Message "  REMAINING: $($r.DisplayName) | Version: $($r.DisplayVersion) | GUID: $($r.ProductGUID)" -Type 2 -Component "Verification"
    }
    $Script:FailureOccurred = $true
}
else {
    Write-CMLog -Message "Verification passed: No targeted .NET products remain." -Component "Verification"
}

# --- Step 7: Summary and exit ---
Write-CMLog -Message "============================================================" -Component "Summary"
Write-CMLog -Message "  .NET Cleanup Script -- Complete" -Component "Summary"
Write-CMLog -Message "  Ended: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Component "Summary"
Write-CMLog -Message "  Products processed: $total" -Component "Summary"
Write-CMLog -Message "  Failures occurred: $($Script:FailureOccurred)" -Component "Summary"
Write-CMLog -Message "  Reboot required: $($Script:RebootRequired)" -Component "Summary"

if ($Script:FailureOccurred) {
    Write-CMLog -Message "  Final exit code: 1 (one or more failures)" -Type 3 -Component "Summary"
    Write-CMLog -Message "============================================================" -Component "Summary"
    exit 1
}
elseif ($Script:RebootRequired) {
    Write-CMLog -Message "  Final exit code: 3010 (success, reboot required)" -Type 2 -Component "Summary"
    Write-CMLog -Message "============================================================" -Component "Summary"
    exit 3010
}
else {
    Write-CMLog -Message "  Final exit code: 0 (all clean)" -Component "Summary"
    Write-CMLog -Message "============================================================" -Component "Summary"
    exit 0
}