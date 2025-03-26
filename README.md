# Uninstall-Office-2016
Powershell script to try and uninstall Office 2016 

How It Works

* Logging to C:\temp:

The script ensures that the folder C:\temp exists and sets the global log file path to that directory.

* Backup, Process Cleanup, & Uninstallation:

It backs up major registry hives, stops Office processes and the “ose” service, then uses the Windows Installer COM object to enumerate and silently uninstall MSI-based Office products. It also calls OfficeC2RClient.exe if a Click‑to‑Run version is detected.

Use this script with caution and adjust the lists (of registry keys, folders, tasks, etc.) to suit your environment.

* Administrative Check & Logging:

The script begins by checking for administrator rights (via Assert-Admin) and setting up a log file.

* Registry Backup:

It backs up several major registry hives to a timestamped folder on the Desktop.

* Process & Service Cleanup:

The functions stop any running Office processes (using a list of known executables) and stop (and disable) the Office Source Engine service.

* Uninstall Office MSI Products:

Using the Windows Installer COM object, the script loops through all MSI products and checks if the product name or version suggests it is Office 2016. If so, it calls msiexec with the product code in quiet mode.

* Uninstall Click‑to‑Run:

If an Office Click‑to‑Run installation is detected (by checking for OfficeC2RClient.exe), it runs it with parameters to update to version 0 (which removes Office).

* Cleanup:

The script then deletes leftover Office folders (from both Program Files and user profile), removes Office scheduled tasks, deletes known registry keys and ARP entries, and finally removes Office-related shortcuts.

*  Final Message:

The log file records all actions, and the script advises that a system reboot may be required.
