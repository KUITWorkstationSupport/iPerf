# Purpose
This runs iPerf (a network testing application), against an optionally specified iPerf server (defaults to iperf.ku.edu), for an optionally specified amount of time per test (defaults to 5 seconds), and finally exports results to CFS.

# Usage
iPerf-Test.ps1 `[[-iPerfServer] <String>] [[-iPerfTime] <Int32>] [[-ExportDir] <String>] [<CommonParameters>]`

# Data Points Collected
1. Time Stamp
2. Client TCP Send Rate (Mb/s)	
3. Client TCP Reverse Receive Rate (Mb/s)	
4. Client UDP Send Packet Loss (%)	
5. Client UDP Reverse Packet Loss (%)	
6. Client UDP Send Jitter (ms)
7. Link Speed (Mb/s)	
8. SSID	
9. Radio Type	
10. Channel	
11. Band	
12. Signal Strength	
13. BSSID	
14. MAC Address	
15. Source IP	
16. Adapter Name	
17. Driver Version	
18. Driver Provider	
19. Driver Date	
20. OS Name	
21. OS Version	
22. OS Last Boot	
23. OS Last Wake	
24. OS Last Wake Source	
25. Power Plan	
26. AC Power Status	
27. Wireless Adapter AC Power Setting	
28. Wireless Adapter DC Power Setting	
29. Username	
30. iPerf Server	
31. Model	
32. Chassis Type

# Setup
* iPerf-Test.ps1 assumes the iPerf executable is in the same directory the script is. 
* The iPerf executable itself expects cygwin1.dll in the same directory.
* Our deployment method has been to store the files on the sysvol, use a GPO to drop the files on the local computer, then setup a scheduled task to run it on some arbitrary interval.

1. Computer Configuration\Preferences\Windows Settings\Folders
   * [Setup the folder](Screenshots/Folder-1.PNG?raw=true)
   * [Remove when no longer applied](Screenshots/Folder-2.PNG?raw=true)
     * Recommended for easier cleanup.

1. Computer Configuration\Preferences\Windows Settings\Files
   * [Copy from source to the folder created by the GPO.](Screenshots/Files-1.PNG?raw=true)
     * Make sure you have `"*.*"` on the end of the source path to ensure all files are copied.
   * [Remove when no longer applied](Screenshots/Files-2.PNG?raw=true)
     * Recommended for easier cleanup.

1. Computer Configuration\Preferences\Windows Settings\Scheduled Tasks
   * [General Options](Screenshots/Task-1.PNG?raw=true)
     * Replace mode for easier cleanup
     * Run as "NT AUTHORITY\SYSTEM"
     * Run only when user is logged on
     * Run with highest privileges
     * Configure for Windows 7/2008R2
   * [Triggers > New Trigger](Screenshots/Task-2.PNG?raw=true)
     * You can run iPerf on whatever schedule you want. We decided on once an hour on a random delay, between 10am-4pm. To do that, set the options like this, then duplicate it for every top of the hour you want it to run at.
   * [Example Triggers](Screenshots/Task-3.PNG?raw=true)
   * [Actions > New Action](Screenshots/Task-4.PNG?raw=true)
     * Start a Program
     * powershell.exe
     * Args: -executionpolicy bypass -file "<path to your GPO's file destination dir>\iPerf-Test.ps1"
     * Start in: either blank or start it in your GPO's file destination and adjust the arg accordingly to ".\iPerf-Test.ps1"
   * [Conditions](Screenshots/Task-5.PNG?raw=true)
     * We wanted it to only run when it had any network connection.
     * Other interesting options here include waking the computer to run the test or only starting if on AC, both of which could be useful specifically for laptops.
   * [Settings](Screenshots/Task-6.PNG?raw=true)
     * Allow task to be run on demand. Also note that you have to launch Task Scheduler as admin to see/run it, and running as admin if invoking from command line.
     * Run task after scheduled start is missed. Optional, but pretty much a requirement for laptops as they might be moving around during their random time window.
     * Stop task if it runs longer. Arbitrarily decided as 4 hours, haven't seen any instances where it ran away.
     * If task is already running, do not start a new instance. I don't think the iPerf executable will handle multiple tests at the same time, haven't tested that. 
   * [Remove when no longer applied](Screenshots/Task-7.PNG?raw=true)
     * Recommended for easier cleanup.