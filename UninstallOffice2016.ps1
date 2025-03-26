<#
.SYNOPSIS
    Completely removes Office 2016 installations (both MSI-based and Click‑to‑Run) and cleans up leftover files, registry keys, scheduled tasks, and shortcuts.

.DESCRIPTION
    This script uses several techniques:
      • It backs up registry hives to a fixed folder under C:\temp\RegistryBackup.
      • It stops common Office processes and the Office Source Engine service.
      • It uses the Windows Installer COM object to enumerate and uninstall Office 2016 MSI products.
      • It explicitly checks for Office Standard 2016 (registry key Office16.STANDARD) and runs its uninstall command.
      • It calls OfficeC2RClient.exe if a Click‑to‑Run installation is detected.
      • Finally, it removes known Office folders, scheduled tasks, registry keys, and shortcuts.
      
    Logs are written to C:\temp.
    
.NOTES
    - Requires Administrator privileges (or running under SYSTEM with proper rights).
    - Tested on Windows 10 (64-bit) with Office 2016.
    - Adjust paths, registry keys, and process names as needed.
#>

#region Helper Functions

function Assert-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Error "This script must be run as Administrator!"
        exit 1
    }
}

function Write-Log {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "$timestamp - $Message"
    Write-Host $logLine
    Add-Content -Path $global:LogFile -Value $logLine
}

#endregion Helper Functions

#region Setup Log Location

# Use a fixed directory for logs.
$logDir = "C:\temp"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
# Set the global log file path with a timestamp.
$global:LogFile = Join-Path $logDir ("Office2016Uninstall_{0}.log" -f (Get-Date -Format "yyyyMMddHHmmss"))
Write-Log "Office 2016 complete uninstall script started."

#endregion Setup Log Location

#region Backup Registry

function Backup-Registry {
    Write-Log "Trying to back up"
    # Use a fixed backup folder under C:\temp\RegistryBackup
    $backupRoot = "C:\temp\RegistryBackup"
    if (-not (Test-Path $backupRoot)) {
        New-Item -ItemType Directory -Path $backupRoot | Out-Null
    }
    $timestamp = (Get-Date).ToString("yyyyMMddHHmmss")
    $backupPath = Join-Path $backupRoot $timestamp
    if (-not (Test-Path $backupPath)) {
        New-Item -ItemType Directory -Path $backupPath | Out-Null
    }
    Write-Log "Backing up registry hives to $backupPath"
    foreach ($hive in @("HKCR", "HKCU", "HKLM", "HKU", "HKCC")) {
        $regFile = Join-Path $backupPath "$hive.reg"
        Write-Log "$regFile"
        try {
            reg export $hive $regFile /y | Out-Null
            Write-Log "Backed up $hive to $regFile"
        } catch {
            Write-Log "ERROR backing up $hive - $_"
        }
    }
    Write-Log "Reg Backup Done"
}

#endregion Backup Registry

#region Process and Service Cleanup

function Stop-OfficeProcesses {
    Write-Log "Trying to kill processes"
    $processNames = @(
        "winword", "excel", "powerpnt", "onenote", "outlook", "mspub", "msaccess", "infopath",
        "groove", "lync", "officeclicktorun", "officeondemand", "officec2rclient",
        "appvshnotify", "firstrun", "setup", "integratedoffice", "integrator",
        "communicator", "msosync", "onenotem", "iexplore", "mavinject32", "werfault",
        "perfboost", "roamingoffice", "msiexec", "ose"
    )
    foreach ($name in $processNames) {\

        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Office processes terminated."
}

function Stop-OfficeService {
    Write-Log "Checking for Office Source Engine service"
    try {
        Stop-Service -Name "ose" -ErrorAction SilentlyContinue
        sc.exe config ose start= disabled | Out-Null
        Write-Log "Office Source Engine service stopped and disabled."
    } catch {
        Write-Log "ERROR stopping service 'ose' - $_"
    }
}

#endregion Process and Service Cleanup

#region Uninstall Office MSI Products

function Uninstall-OfficeMSI {
    Write-Log "Running Uninstall-OfficeMSI"
    Write-Log "Enumerating Office MSI products for removal..."
    try {
        $msi = New-Object -ComObject WindowsInstaller.Installer
    } catch {
        Write-Error "Failed to create WindowsInstaller.Installer object: $_"
        return
    }
    $officeProducts = @()
    foreach ($prod in $msi.Products) {
        try {
            $prodName = $msi.ProductInfo($prod, "ProductName")
            $prodVersion = $msi.ProductInfo($prod, "ProductVersion")
        } catch {
            continue
        }
        if ($prodName -match "Office 2016" -or ($prodVersion -like "16.*")) {
            $officeProducts += $prod
        }
    }
    if ($officeProducts.Count -eq 0) {
        Write-Log "No Office 2016 MSI products found."
    } else {
        Write-Log "Found $($officeProducts.Count) Office MSI product(s) to remove."
        foreach ($prod in $officeProducts) {
            $prodName = $msi.ProductInfo($prod, "ProductName")
            Write-Log "Uninstalling $prodName (ProductCode: $prod)"
            $uninstallCmd = "msiexec.exe /x{$prod} /qn /norestart"
            try {
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait -ErrorAction Stop
                Write-Log "Uninstall command completed for $prodName."
            } catch {
                Write-Log "ERROR uninstalling $prodName - $_"
            }
        }
    }
}

#endregion Uninstall Office MSI Products

#region Uninstall Office Standard 2016

function Uninstall-OfficeStandard {
    Write-Log "Running Uninstall-OfficeStandard"
    # Specifically check for Office Standard 2016 via its registry key.
    $regKeyPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Office16.STANDARD"
    if (Test-Path $regKeyPath) {
         try {
              $uninstallString = (Get-ItemProperty -Path $regKeyPath -Name "UninstallString").UninstallString
              if ($uninstallString) {
                    Write-Log "Found Office Standard uninstall string: $uninstallString"
                    if ($uninstallString -notmatch "/quiet") {
                        $uninstallString = "$uninstallString /quiet"
                    }
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString" -Wait -ErrorAction Stop
                    Write-Log "Office Standard uninstallation initiated."
              } else {
                    Write-Log "No UninstallString found in $regKeyPath."
              }
         } catch {
              Write-Log "ERROR uninstalling Office Standard - $_"
         }
    } else {
         Write-Log "Registry key $regKeyPath not found; Office Standard may not be installed."
    }
}

#endregion Uninstall Office Standard 2016

#region Uninstall Click-to-Run

function Uninstall-OfficeC2R {
    Write-Log "Running ninstall-OfficeC2R"
    $c2rPath = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $c2rPath) {
        Write-Log "Office Click-to-Run installation detected."
        $arguments = "/update user updatetoversion=0 /uninstall"
        try {
            Start-Process -FilePath $c2rPath -ArgumentList $arguments -Wait -ErrorAction Stop
            Write-Log "Click-to-Run uninstall initiated."
        } catch {
            Write-Log "ERROR during Click-to-Run uninstall - $_"
        }
    } else {
        Write-Log "No Office Click-to-Run installation found."
    }
}

#endregion Uninstall Click-to-Run

#region Remove Leftover Files and Folders

function Remove-OfficeFolders {
    Write-Log "Removing Office installation folders..."
    $folders = @(
        "$env:ProgramFiles\Microsoft Office 15",
        "$env:ProgramFiles(x86)\Microsoft Office 15",
        "$env:ProgramFiles\Microsoft Office\root",
        "$env:ProgramFiles(x86)\Microsoft Office\root",
        "$env:ProgramFiles\Microsoft Office",
        "$env:ProgramFiles(x86)\Microsoft Office",
        "$env:ProgramData\Microsoft\ClicToRun",
        "$env:COMMONPROGRAMFILES\Microsoft Shared\ClickToRun",
        "$env:COMMONPROGRAMFILES(x86)\Microsoft Shared\ClickToRun",
        (Join-Path $env:USERPROFILE "Microsoft Office"),
        (Join-Path $env:USERPROFILE "Microsoft Office 15"),
        (Join-Path $env:USERPROFILE "Microsoft Office 16")
    )
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            try {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Removed folder: $folder"
            } catch {
                Write-Log "ERROR removing folder $folder - $_"
            }
        }
    }
}

#endregion Remove Leftover Files and Folders

#region Remove Scheduled Tasks

function Remove-OfficeScheduledTasks {
    Write-Log "Removing Office-related scheduled tasks..."
    $tasks = @(
        "FF_INTEGRATEDstreamSchedule",
        "FF_INTEGRATEDUPDATEDETECTION",
        "C2RAppVLoggingStart",
        "Office 15 Subscription Heartbeat",
        "\Microsoft\Office\Office 15 Subscription Heartbeat",
        "Office 15 Sync Maintenance",
        "\Microsoft\Office\OfficeInventoryAgentFallBack",
        "\Microsoft\Office\OfficeTelemetryAgentFallBack",
        "\Microsoft\Office\OfficeInventoryAgentLogOn",
        "\Microsoft\Office\OfficeTelemetryAgentLogOn",
        "Office Background Streaming",
        "\Microsoft\Office\Office Automatic Updates",
        "\Microsoft\Office\Office ClickToRun Service Monitor",
        "Office Subscription Maintenance",
        "\Microsoft\Office\Office Subscription Maintenance"
    )
    foreach ($task in $tasks) {
        try {
            schtasks.exe /delete /tn $task /f | Out-Null
            Write-Log "Removed task: $task"
        } catch {
            Write-Log "ERROR removing task $task - $_"
        }
    }
}

#endregion Remove Scheduled Tasks

#region Remove Registry Entries

function Remove-OfficeRegistryKeys {
    Write-Log "Removing Office-related registry keys..."
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Office\15.0",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0",
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKCU:\Software\Microsoft\Office\15.0",
        "HKCU:\Software\Microsoft\Office\16.0",
        "HKCU:\Software\Microsoft\Office"
    )
    foreach ($key in $keys) {
        if (Test-Path $key) {
            try {
                Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Removed registry key: $key"
            } catch {
                Write-Log "ERROR removing registry key $key - $_"
            }
        }
    }
    $arpKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ENTERPRISE",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PROPLUS"
    )
    foreach ($key in $arpKeys) {
        if (Test-Path $key) {
            try {
                Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
                Write-Log "Removed ARP key: $key"
            } catch {
                Write-Log "ERROR removing ARP key $key - $_"
            }
        }
    }
}

#endregion Remove Registry Entries

#region Remove Shortcuts

function Remove-OfficeShortcuts {
    Write-Log "Removing Office shortcuts from Start Menu and Desktop..."
    $paths = @(
        "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016",
        "$env:USERPROFILE\Desktop"
    )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            try {
                Get-ChildItem -Path $path -Filter "*2016*.lnk" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Log "Removed shortcuts in: $path"
            } catch {
                Write-Log "ERROR removing shortcuts in $path - $_"
            }
        }
    }
}

#endregion Remove Shortcuts
#region Main Execution

Assert-Admin

#Backup-Registry

#Stop-OfficeProcesses

#Stop-OfficeService

Uninstall-OfficeMSI

Uninstall-OfficeC2R

Uninstall-OfficeStandard

Start-Sleep -Seconds 10

Remove-OfficeFolders

Remove-OfficeScheduledTasks

Remove-OfficeRegistryKeys

#Remove-OfficeShortcuts

Write-Log "Office 2016 complete removal process finished. A system reboot may be required."

#endregion Main Execution
