# Modular PowerShell Application Uninstaller System

## Overview

A modular, extensible PowerShell framework for uninstalling Windows applications with centralized logging, pattern-based dispatch, and automatic fallback logic.

**Version**: 1.0  
**Author**: V.Ashodhiya  
**Date**: 2026-01-27

## Features

- ✅ **Loose Pattern Matching**: Case-insensitive regex patterns match app names flexibly
- ✅ **Modular Architecture**: One file per app, easy to extend
- ✅ **CMTrace Logging**: Industry-standard log format compatible with CMTrace.exe
- ✅ **Automatic Fallback**: App-specific function → Registry search → Failure
- ✅ **Batch Uninstall**: Removes all installed versions of an app in one run
- ✅ **Structured Output**: Detailed summary with exit codes and version lists
- ✅ **Silent Operation**: No user prompts, ideal for automation

## Quick Start

### Basic Usage

```powershell
# Uninstall KeePass (uses app-specific function)
.\MainUninstallScript.ps1 -AppName "KeePass"

# Uninstall Adobe Reader (uses app-specific function)
.\MainUninstallScript.ps1 -AppName "Adobe"

# Uninstall unknown app (falls back to registry search)
.\MainUninstallScript.ps1 -AppName "SomeApp"

# Use custom log directory
.\MainUninstallScript.ps1 -AppName "Chrome" -LogPath "C:\Logs"
```

### Get Help

```powershell
Get-Help .\MainUninstallScript.ps1 -Full
```

## Directory Structure

```
Application-Uninstall/
├── Core/
│   └── Uninstall.Core.psm1         # Core dispatcher and utilities
├── Apps/
│   ├── Uninstall.KeePass.ps1       # KeePass uninstaller
│   ├── Uninstall.AdobeReader.ps1   # Adobe Reader uninstaller
│   └── Uninstall.Template.ps1      # Template for new apps
├── MainUninstallScript.ps1         # Entry point
└── README.md                        # This file
```

## How It Works

### 1. Pattern Matching

The core module contains a mapping table:

```powershell
$Global:AppMapping = @{
    '^keepass'         = 'Uninstall-KeePass'
    'adobe.*reader'    = 'Uninstall-AdobeReader'
    'adobe.*acrobat'   = 'Uninstall-AdobeReader'
}
```

Input is matched case-insensitively against these patterns.

### 2. Dispatch Flow

```
User Input → Pattern Match → App Function → Success/Failure
                         ↓
                    No Match → Registry Fallback → Success/Failure
                                              ↓
                                         App Function Failed → Registry Fallback → Success/Failure
```

### 3. Logging

All operations are logged to:

- **Location**: `C:\Windows\fndr\logs\` (or custom via `-LogPath`)
- **Format**: CMTrace-compatible
- **Naming**: `ModularUninstall_<AppName>_<timestamp>.log`

View logs with [CMTrace.exe](https://learn.microsoft.com/en-us/mem/configmgr/core/support/cmtrace) for best experience.

## Adding New Applications

Follow these 4 easy steps:

### Step 1: Copy Template

```powershell
Copy-Item "Apps\Uninstall.Template.ps1" "Apps\Uninstall.MyApp.ps1"
```

### Step 2: Edit Function

Open `Apps\Uninstall.MyApp.ps1` and:

- Rename function from `Uninstall-Template` to `Uninstall-MyApp`
- Update the `Where-Object` filter to match your app:

  ```powershell
  $_.DisplayName -like '*MyApp*'
  ```

- Add any app-specific uninstall logic

### Step 3: Update Mapping Table

Edit `Core\Uninstall.Core.psm1` and add to `$Global:AppMapping`:

```powershell
$Global:AppMapping = @{
    # ... existing entries ...
    'myapp'  = 'Uninstall-MyApp'  # Add this line
}
```

### Step 4: Test

```powershell
.\MainUninstallScript.ps1 -AppName "MyApp" -LogPath "$env:TEMP"
```

Check the log file for results!

## Command Reference

### MainUninstallScript.ps1

**Parameters**:

- `-AppName` (required): Name or pattern of app to uninstall
- `-LogPath` (optional): Directory for log files (default: `C:\Windows\fndr\logs`)

**Exit Codes**:

- `0`: Success (at least one version uninstalled)
- `1`: Failure (no versions found or all failed)

### Output Structure

```json
{
    "AppNameInput": "KeePass",
    "MatchedPattern": "^keepass",
    "SelectedFunction": "Uninstall-KeePass",
    "VersionsHandled": ["KeePass 2.47 (2.47)", "KeePass 2.46 (2.46)"],
    "PrimaryMethod": "AppSpecific",
    "FallbackUsed": false,
    "FinalExitCode": 0,
    "Status": "Success"
}
```

## Examples

### Example 1: Multiple Versions

```powershell
PS> .\MainUninstallScript.ps1 -AppName "keepass"

Found KeePass: KeePass 2.47 (2.47)
Successfully uninstalled: KeePass 2.47 (2.47)
Found KeePass: KeePass 2.46 (2.46)
Successfully uninstalled: KeePass 2.46 (2.46)

Status: Success
Versions Handled: 2
```

### Example 2: Pattern Fallback

```powershell
PS> .\MainUninstallScript.ps1 -AppName "unknownapp"

No pattern match found. Using registry-based uninstall.
No applications matching 'unknownapp' found in registry

Status: Failure
Exit Code: 1
```

### Example 3: Case-Insensitive Matching

```powershell
PS> .\MainUninstallScript.ps1 -AppName "ADOBE"
PS> .\MainUninstallScript.ps1 -AppName "adobe"
PS> .\MainUninstallScript.ps1 -AppName "AdObE"

# All three match the same pattern: 'adobe.*reader'
```

## Troubleshooting

### Permissions Error

**Problem**: `Access denied` when writing to `C:\Windows\fndr\logs`

**Solution**: Use a writable log directory:

```powershell
.\MainUninstallScript.ps1 -AppName "MyApp" -LogPath "$env:TEMP"
```

### App Not Found

**Problem**: Script reports "No applications found"

**Checklist**:

1. Is the app actually installed? Check registry:

   ```powershell
   Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | 
       Where-Object { $_.DisplayName -like "*AppName*" }
   ```

2. Is the app-specific filter correct?
3. Try the generic registry fallback to test

### Function Not Found

**Problem**: `Function 'Uninstall-MyApp' not found`

**Checklist**:

1. Is the .ps1 file in the `Apps\` directory?
2. Does the function name match the mapping table?
3. Did you reload the module? Use `-Force` when importing:

   ```powershell
   Import-Module .\Core\Uninstall.Core.psm1 -Force
   ```

### CMTrace Logs

To view logs properly:

1. Download CMTrace.exe from Microsoft
2. Run: `CMTrace.exe "C:\Windows\fndr\logs\ModularUninstall_*.log"`
3. Logs will show with:
   - Color-coded severity (Info/Warning/Error)
   - Timestamps with milliseconds
   - Component names
   - Thread IDs

## Technical Details

### Registry Paths Searched

1. `HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*` (64-bit apps)
2. `HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*` (32-bit apps on 64-bit Windows)
3. `HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*` (per-user apps)

### Uninstall Methods

1. **QuietUninstallString** (preferred)
2. **UninstallString** + silent switches
3. **Product Code** + `msiexec /x {GUID} /qn /norestart`

### MSI Silent Switches

- `/qn` - No UI
- `/norestart` - Suppress automatic reboot
- `/x` - Uninstall mode

## Best Practices

1. **Always test** in a non-production environment first
2. **Run as Administrator** for system-wide apps
3. **Check logs** after each run for details
4. **Use specific patterns** in mapping table to avoid false matches
5. **Document exclusions** (versions, editions) in app-specific functions
6. **Exit codes matter** - Use them in automation scripts

## Known Limitations

- Requires PowerShell 5.1 or higher
- Does not support:
  - Store apps (UWP/AppX packages)
  - Apps without registry entries
  - Apps requiring user interaction during uninstall
- Log file permissions depend on target directory

## Version History

**1.0** (2026-01-27)

- Initial release
- Core module with dispatcher
- CMTrace-compatible logging
- KeePass and Adobe Reader examples
- Template for new apps
- Registry fallback logic

## License

Internal use only. See LICENSE file for details.

## Support

For issues or questions:

1. Check the troubleshooting section
2. Review log files in CMTrace
3. Test with registry fallback mode
4. Contact: V.Ashodhiya
