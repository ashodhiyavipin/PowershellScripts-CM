# PowerShellScripts-CM

[![License](https://img.shields.io/github/license/ashodhiyavipin/PowershellScripts-CM)](https://github.com/ashodhiyavipin/PowershellScripts-CM/blob/main/LICENSE)

A comprehensive collection of production-ready PowerShell scripts primarily designed for system administration, Microsoft Configuration Manager (SCCM/MECM) environment management, Windows endpoint diagnostics, and automated software remediations.

## 📂 Repository Contents

This repository is organized into several key enterprise IT administrative categories.

### Windows Update & Maintenance
* `RepairResetWUAComponents-New.ps1` / `WUARemediations.ps1` - Advanced scripts to diagnose and repair broken Windows Update Agent components and soft-reset software distribution folders.
* `ApplyStandalonePatch.ps1` / `ApplyStandaloneMSU.ps1` - Automations for silently extracting and applying standalone Windows update packages (.msu/.cab).
* `MonthlySUGCleanup.ps1` - Maintenance script for Software Update Groups.

### Application Lifecycle Management
* `Download-UWPApps.ps1` - Robust, idempotent script for downloading AppX/UWP applications cleanly via Winget for offline provisioning.
* `ApplicationUninstallScript.ps1` & `Application-Uninstall/` - Framework and customizable templates for standardized native application uninstalls.
* `RemoveOldAppxVersions.ps1` - Cleans up staged, superseded AppX packages to reclaim OS storage space.
* `OfficeUninstallation.ps1` / `UninstallAdobeAcrobat.ps1` / `MicrosoftNET-SDKRemoval.ps1` - Targeted forced-removal tools to scrub complex enterprise software suites.

### Configuration Manager (SCCM) Operations
* `Add-CMDevice.ps1` / `Remove-CMDevice.ps1` - Scripted automation for managing device populations within specific Collections.
* `CcmClearPolicy.ps1` - Forces a deep wipe and dynamic refresh of the local CCM Client policy cache.

### Endpoint Optimization
* `UninstallLegacyPrinterDrivers.ps1` - Deep cleans stale and problematic legacy v3/v4 printer drivers from the system driver store.
* `FreeDiskSpace.ps1` - Enterprise-safe automated storage cleanup routines clearing temp directories and caches.
* `debloat.ps1` - Windows OS bloatware removal and baseline optimization.

## 🚀 Usage Guidelines

1. **Execution Policy**: Most operational environments restrict unsigned script execution. Ensure your policy permits execution for local files:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
2. **Privilege Limitations**: Scripts altering the Windows Update Agent, Driver Store, or the CCM namespace **must** be executed from an elevated PowerShell session (`Run as Administrator`).

## ⚖️ Legal Disclaimer & Limitation of Liability

The scripts, documentation, and tooling provided in this repository are supplied **"AS IS"**, without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement. 

In no event shall the author(s), publisher(s), or contributors be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.

**Risk Acknowledgment:** Many scripts in this repository perform high-impact administrative actions, including the irreversible manipulation of the Windows Registry, Component Store (WinSxS), standard applications, and the Windows Update Agent. 
* By executing these tools, you expressly assume all risks associated with their use. 
* It is **strictly mandated** that you independently review, audit, and validate the source code prior to execution.
* **Never** deploy these scripts to a production environment or across an enterprise fleet (via SCCM, Intune, RMM, or otherwise) without prior exhaustive testing in a secure, isolated sandbox environment.

**By downloading, cloning, or executing any code from this repository, you explicitly agree that assuming the operational risk is solely your responsibility, and you irrevocably waive any right to hold the author(s) liable for operational disruptions, data loss, or systemic failures.**

## 📝 License

This repository is licensed under the standard MIT/Open-Source terms defined in the local [LICENSE](LICENSE) file.
