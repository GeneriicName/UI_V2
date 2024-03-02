# Utility-menu-V2

A utility menu designed to streamline IT tasks and automate repetitive processes. This tool is especially helpful for Help Desks and IT professionals.

This is an updated version with a focus on UI improvements using PyQt6, licensed under GPL3.

![Utility Menu Screenshot](https://github.com/GeneriicName/UI_V2/assets/139624416/1c906f12-4288-4e49-a3d4-ff2692c5b135)

## Features

- **Clean Space from Remote Computers**: Configure directories to clean up via the config file, including options to delete the Windows search edb file.
- **Get Network Printers**: Retrieve network printers installed via print servers, TCP/IP, and WSD, including IP and print server information.
- **Delete OST File**: Remove OST files from remote computers.
- **Reset Print Spooler**: Quickly reset the print spooler service.
- **Sample Function**: Replace with your own function for customized tasks.
- **Fix Internet Explorer**: (May not work on all OS versions).
- **Fix Cockpit Printers**: Delete registry keys related to Jetro Cockpit printers.
- **Close Outlook and Lync/Skype**: Close Outlook and related processes.
- **Delete Zoom and Teams**: Uninstall Zoom (64/32 bit), Zoom Outlook plugin, and Microsoft Teams application.
- **Export**: Export network printers and mapped drives into .txt and .bat files for easy reinstallation.
- **Fix 3 Languages Bug**: Resolve a bug where the same language is displayed twice.
- **Delete User Folders**: Choose users to delete their folders, displaying their display names for clarity. Supports local multithreaded (UNC) or remote WMI (The remote computer will do the work) deletion methods.
- **Customizable Colors**: Easily customize the colors used in the script via the settings button, the default themes are light or dark.


Additionally, the script provides comprehensive information about the computer and user, including status, disk space, uptime, and more.

## Installation

**Requirements: Python 3.11-12, Windows 10**

```batch
git clone https://github.com/GeneriicName/UI_V2
cd UI_V2
pip install -r requirements.txt
```
## Configuration

This is an example of the config file which is included with the directory.

| Key | Value | Description |
| :--- | :--- | :--- |
| "log" | "\\\\path\\to\\logfile.log" | this is the path to the logfile if false, it wont log errors |
| "domain" | "DC=example,DC=domain,DC=com" | set your domain with ldap |
| "print_servers" | ["\\\\print_svr01", "\\\\print_svr02", "\\\\print_svr03"] | path to your print servers, list them with network path and double backslashes |
| "max_workers" | 8 | the max threads for the program to use when deleting files, notice that the program it self uses 2 threads so take it into account |
| "to_delete" | [["windows\\ccmcache", "Deleting ccsm cashe", "Deleted ccsm cashe"], ["temp"], ["Windows\\Temp", "Deleting windows temp files", "Deleted windows temp files"]] | paths to extra None user specific folders to delete their contents, and optional prompt, leave out the \\\\computername\\c$\\ |
| "user_specific_delete" | [] | paths to user specific folders to delete, and optional prompt, leave out the \\\\computername\\c$\\user, in the prompt you can use users_amount to insert the amount of users |
| "delete_user_temp" | true | delete temp files of each user? set true to if so |
| "delete_edb" | true | delete search.edb? set true if so |
| "do_not_delete" | ["public","default", "default user", "all users", "desktop.ini"] | set the usernames to exclude them from being deleted by the script |
| "start_with_exclude" | ["admin"] | add prefixes of usernames to exclude them, from being deleted |
| "users_txt" | "\\\\path\\to\\folder\\with\\user.txt files" | path of folder which contains computer names in usename.txt files |
| "assets" | "\\path\to\directory" | path to assets such as images |
| "title" | "hello world!" | give a title to your GUI window |

**To enable the script to work with both usernames and hostnames, you'll require a user.txt file containing the computer name from which the last user has logged on for each user. You can generate this file using a simple batch logon script/GPO/task. Ensure that the location where these files are saved is configured in the config.json file.**

***Example logon script***

```batch
@echo off
echo %computername% > "\\server\folder\%username%.txt"
```

## Additional Information

- **Privacy**: Certain features have been removed to protect sensitive information. However, the script includes valuable features such as network printer retrieval, user/computer information display, and GUI functionality.
- **Assets**: The "assets" folder contains images for the GUI application.
- **Compatibility**: Tested on Windows 10-11 and Python3.11-12; compatibility with earlier versions not guaranteed.
- **Logging**: Basic support for logging errors is included.
- **Modularity**: Script is not split into multiple modules due to time constraints.
- **Support**: I might be able to assist if you need help understanding or modifying the script. Feel free to reach out!
  
Feel free to contribute or fork this project for your own use!

