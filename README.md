# Utility-menu-V2

A utility menu I made while working as a help desk. It was made in between calls, so I'm aware it's far from perfect.

This is an updated version (mostly UI).

## Installation

```batch
git clone https://github.com/GeneriicName/UI_V2
cd UI_V2
pip install -r requirements.txt
```


## Configuration

This is an example of the config file which is included with the directory.

| Key | Value | Description |
| :---         | :---      | :---          |
| "log"   | "\\\\path\\to\\logfile.log"     | this is the path to the logfile if false, it wont log errors    |
| "domain"     | "DC=example,DC=domain,DC=com"       | set your domain with ldap      |
| "print_servers"   | ["\\\\print_svr01", "\\\\print_svr02", "\\\\print_svr03"]     | path to your print servers, list them with network path and double backslashes    |
| "max_workers"     | 8       | the max threads for the program to use when deleting files, notice that the program it self uses 2 threads so take it into account      |
| "to_delete"     | [["windows\\ccmcache", "Deleting ccsm cashe", "Deleted ccsm cashe"], ["temp"], ["Windows\\Temp", "Deleting windows temp files", "Deleted windows temp files"]]       | paths to extra None user specific folders to delete their contents, and optional prompt, leave out the \\\\computername\\c$\\      |
| "user_specific_delete"   | []     | paths to user specific folders to delete, and optional prompt, leave out the \\\\computername\\c$\\user, in the prompt you can use users_amount to insert the amount of users    |
| "delete_user_temp"     | true       | delete temp files of each user? set true to if so      |
| "delete_edb"   | true     | delete search.edb? set true if so    |
| "do_not_delete"     | ["public","default", "default user", "all users", "desktop.ini"]       | set the usernames to exclude them from being deleted by the script      |
| "start_with_exclude"   | ["admin"]     | add prefixes of usernames to exclude them, from being deleted    |
| "users_txt"     | "\\\\path\\to\\folder\\with\\user.txt files"     | path of folder which contains computer names in usename.txt files      |
| "assets"     | "\\path\to\directory"       | path to assets such as images      |
| "title"     | "hello world!"       | give a title to your GUI window      |




**To use this script with usernames as well as hostnames, you will need a user.txt file containing the computer name from which the last user has logged on for each user. You can easily achieve this with a simple batch logon script/GPO/task, and the location to which the files are dumped needs to be configured in the config.json file.**

***Example logon script***

```batch
@echo off
echo %computername% > "\\server\folder\%username%.txt"
```

## Features
**Clean space from remote computers:** You can configure which directories to clean up via the config file, as well as whether to delete the Windows search edb file.

**Get network printers from the computer:** This includes printers installed via print servers, TCP/IP, and WSD. It will retrieve the IP and tell you if it's found on any of the print servers.

**Delete the OST file from the remote computer.**

**Reset print spooler.**

**Fix Internet Explorer:** Depending on your OS version, this might not work.

**Fix cockpit printers:** This deletes the appropriate registry keys. Useful only if your organization uses Jetro Cockpit.

**Close outlook:** This will close outlook and lync (skype) related processes.

**Export**: Export the current user's network printes and mapped drives into both a .txt file and a .bat file, the .bat file will attempt to re-install all the server-installed network printers or TCP/IP-WSD that were found on the server, as well as remap the network drives.

**Fix 3 languages bug:** Fixes a bug when the same language is displayed twice.

**Delete users folders:**, Choose users to delete their folders in order to cleans up space. If the user is in your domain, it'll show their display name to make it clearer which users you are deleting, will exclude the current user of the remote computer, deletion methods are either local threads or WMI (remote computer will do the work), as well as additonal users that you can configure in the config.json file

**The script also displays a bunch of information on the computer and user:**
1. The script shows the user status (active, locked, disabled, expired, or password expired).
2. Similarly, it shows the computer status, including space in the C or D disk, uptime, current user, and whether the computer is online/offline. 
3. A bunch more quality of life features. 

## Additonal information

**I removed some Features:**
Because some features won't be relevant for anyone but the company i work at, or i can share them due to privacy.

**I'm fully aware that this isn't a completed program. I uploaded it with some features missing (only relevant for the company I'm working for), and I wouldn't want to release sensitive information such as AD groups, users, paths, etc. However, this script contains a few nice features such as retrieving network printers from a remote computer, displaying user and computer information, multithreaded file deletion, and a fairly simple GUI. You can take and modify these features for your own project.**

**The "assets" folder contains all the images for the GUI application. You can paste a username/computer name/IP address or network path of a printer in the entry box for translation to network their network path in the print servers or their IP address**

**This was tested on windows 10-11, i cannot guarantee it will work on earlier versions, and was built with python 3.9**

***This script features support for logging, although not fully. However, points of the script that are most likely prone to exceptions are covered.***


***I didn't split the script into multiple modules as I did not enoguh enough time on my hands to do it.***

##
I won't be updating this script any further, but if you need help understanding why I did something in my script or how it works or help modifying it for your own usage, feel free to ask.




![preview](https://github.com/GeneriicName/UI_V2/assets/139624416/42129d66-2d6f-41c0-a902-b8dadff058e0)
