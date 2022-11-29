# MTRoW
PowerShell scripts for bulk configuration of Teams MTR on Windows devices

BATCH CREATION OF XML CONFIGURATION FILES

The MTR_Create_XML.ps1 file creates a unique XML config file for each MTRoW device using an excel sheet to create variables for the resource account name and password. The XML in the script is based on Microsoft's XML template below. Note that Microsoft updates this regularly so you may want to check back and update the template in the script.

https://learn.microsoft.com/en-us/microsoftteams/rooms/xml-config-file

When using the XML file, keep a couple things in mind
1) Before going through the effort to create the XML config files, make sure you have a way to place the XML file in the correct directly on the MTR. Options include:
-Remote Powershell (must be enabled on the MTR - see https://learn.microsoft.com/en-us/microsoftteams/rooms/rooms-operations)
-InTune (it would be a pain to load a unique script for each MTR in InTune, but for bulk settings updates this could work. A good overview here       
        https://blog.chiffers.com/2021/09/14/managing-a-microsoft-teams-room-mtr-device-with-intune-part-3-configuration-profiles/)
-Config Manager if joined to the domain.
2) You do not have to use the entire template. As long as the syntax of the XML is correct, you can use only what you need.
3) You can customize the XML to be a full, initial configuration for the MTR - including the resource account name and password *or* use only certain parts of the XML to update one or more settings (e.g. Turn Cortana off across all MTRs).  If you are not using the resource account/password in the script you will not need the input file. 


To get started
1. Create the directory C:\Temp\mtr_xml on your machine (you can use any directory, but would need to modify the script)
2. Download the teams_rooms.xlsx input file and customize:
    -Column 1 is a friendly name for each account. You can use the display name if you like, or a nickname that makes it easy to recognize the accounts.  
    -Column 2 is the resource account UPN 
    -Column 3 is the resource account password
    -Column 4 is the MTR IP address or hostname. This column is used to name each XML file (e.g. 10.0.0.1.xml or hostname.xml)


