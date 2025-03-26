<#
.SYNOPSIS
    Completely removes Office 2016 MSI/Click-to-Run installations and cleans up leftover files, registry keys, scheduled tasks, and shortcuts.

.DESCRIPTION
    This script combines techniques from a CMD tool and a VBScript to:
      • Enumerate and uninstall Office 2016 MSI products via the Windows Installer COM object.
      • Uninstall Office Click-to-Run installations.
      • Terminate Office-related processes and stop the Office Source Engine service.
      • Remove leftover Office directories from Program Files, ProgramData, and user profiles.
      • Remove Office-related scheduled tasks.
      • Remove Office-related registry keys and ARP entries.
      • Remove Office-related shortcuts.
      
    Logs are saved to **C:\temp**.
    
.NOTES
    - Requires Administrator privileges.
    - Tested on Windows 10 (64-bit) with Office 2016.
    - Adjust paths and key lists as needed.
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

# Ensure that the log directory (C:\temp) exists.
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
    # Create a backup folder on the Desktop with a timestamp
    $timestamp = (Get-Date).ToString("yyyyMMddHHmmss")
    $backupPath = Join-Path $env:USERPROFILE "Desktop\RegistryBackup\$timestamp"
    if (-not (Test-Path $backupPath)) {
        New-Item -ItemType Directory -Path $backupPath | Out-Null
    }
    Write-Log "Backing up registry hives to: $backupPath"
    foreach ($hive in @("HKCR", "HKCU", "HKLM", "HKU", "HKCC")) {
        $regFile = Join-Path $backupPath "$hive.reg"
        try {
            reg export $hive $regFile /y | Out-Null
            Write-Log "Backed up $hive to $regFile"
        } catch {
            Write-Log "ERROR backing up $hive: $_"
        }
    }
}

#endregion Backup Registry

#region Process and Service Cleanup

function Stop-OfficeProcesses {
    # List of common Office processes (MSI-based and Click-to-Run)
    $processNames = @(
        "winword", "excel", "powerpnt", "onenote", "outlook", "mspub", "msaccess", "infopath",
        "groove", "lync", "officeclicktorun", "officeondemand", "officec2rclient",
        "appvshnotify", "firstrun", "setup", "integratedoffice", "integrator",
        "communicator", "msosync", "onenotem", "iexplore", "mavinject32", "werfault",
        "perfboost", "roamingoffice", "msiexec", "ose"
    )
    foreach ($name in $processNames) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Write-Log "Office processes terminated."
}

function Stop-OfficeService {
    # Stop and disable the Office Source Engine service ("ose")
    try {
        Stop-Service -Name "ose" -ErrorAction SilentlyContinue
        sc.exe config ose start= disabled | Out-Null
        Write-Log "Office Source Engine service stopped and disabled."
    } catch {
        Write-Log "ERROR stopping service 'ose': $_"
    }
}

#endregion Process and Service Cleanup

#region Uninstall Office MSI Products

function Uninstall-OfficeMSI {
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
            Write-Log "Uninstalling: $prodName (ProductCode: $prod)"
            $uninstallCmd = "msiexec.exe /x{$prod} /qn /norestart"
            try {
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait -ErrorAction Stop
                Write-Log "Uninstall command completed for $prodName."
            } catch {
                Write-Log "ERROR uninstalling $prodName: $_"
            }
        }
    }
}

#endregion Uninstall Office MSI Products

#region Uninstall Click-to-Run

function Uninstall-OfficeC2R {
    $c2rPath = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
    if (Test-Path $c2rPath) {
        Write-Log "Office Click-to-Run installation detected."
        $arguments = "/update user updatetoversion=0 /uninstall"
        try {
            Start-Process -FilePath $c2rPath -ArgumentList $arguments -Wait -ErrorAction Stop
            Write-Log "Click-to-Run uninstall initiated."
        } catch {
            Write-Log "ERROR during Click-to-Run uninstall: $_"
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
        "$env:ProgramFiles\Microsoft Office",  # Use with caution: may delete entire Office folder
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
                Write-Log "ERROR removing folder $folder: $_"
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
            Write-Log "Error removing task: $task"
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
                Write-Log "ERROR removing registry key $key: $_"
            }
        }
    }
    # Example: remove known ARP entries (customize as needed)
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
                Write-Log "ERROR removing ARP key $key: $_"
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
                Write-Log "ERROR removing shortcuts in $path: $_"
            }
        }
    }
}

#endregion Remove Shortcuts

#region Main Execution

# Ensure running as Administrator.
Assert-Admin

# Optional: Back up the registry before making changes.
Backup-Registry

# Terminate Office processes.
Stop-OfficeProcesses

# Stop the Office Source Engine service.
Stop-OfficeService

# Uninstall MSI-based Office 2016 products.
Uninstall-OfficeMSI

# Uninstall Click-to-Run installations.
Uninstall-OfficeC2R

# Allow time for uninstall processes to settle.
Start-Sleep -Seconds 10

# Remove leftover Office folders and files.
Remove-OfficeFolders

# Remove Office-related scheduled tasks.
Remove-OfficeScheduledTasks

# Remove known Office registry keys.
Remove-OfficeRegistryKeys

# Remove Office shortcuts.
Remove-OfficeShortcuts

Write-Log "Office 2016 complete removal process finished. A system reboot may be required."

#endregion Main Execution
