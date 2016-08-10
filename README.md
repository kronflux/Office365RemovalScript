# Office365RemovalScript
This script attempts to completely and aggressively remove Microsoft Office 365 and related components.
It first attempts to kill any running tasks related to Office, then attempts a silent uninstall.
After the silent uninstall, it will run more aggressive options such as taking ownership of related files and straight up deleting them from the system.
It also attempts to remove related Registry entries and Scheduled Tasks.

Run this script at your own risk, it will completely break any Office 365(2013 or 2016) and possibly older Office installs.
