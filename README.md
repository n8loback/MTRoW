# MTRoW
PowerShell scripts for bulk configuration of Teams MTR on Windows devices

BATCH CREATION OF XML CONFIGURATION FILES

The MTR_Create_XML.ps1 file creates a unique XML config file for each MTRoW device, using an excel input file to create unique variables in each file for the MTR resource account and password.

The XML in the script is based on Microsoft's XML template below. Microsoft updates this regularly so you may want to check back and update the template in the script.

https://learn.microsoft.com/en-us/microsoftteams/rooms/xml-config-file

Before creating XML files in bulk, keep a few things in mind:
1) Make sure you have a plan for how to place the XML file in the correct directory on the MTR:
   C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState
      Options include
        -Remote Powershell (must be enabled on the MTR - see https://learn.microsoft.com/en-us/microsoftteams/rooms/rooms-operations)
        -InTune (it would be a pain to load a unique script for each MTR in InTune, but for bulk settings updates this could work. A good overview here       
         https://blog.chiffers.com/2021/09/14/managing-a-microsoft-teams-room-mtr-device-with-intune-part-3-configuration-profiles/)
         -Config Manager if joined to the domain.

2) You do not have to use the entire template. As long as the syntax of the XML is correct, you can slice/dice and use only what you need. For example, an initial  
   configuration file could contain a full XML, with account info. Subsequent XML updates could be used to toggle particular settings (e.g. turn third-party join on)


To get started
1. Create the directory C:\Temp\mtr_xml on your machine (you can use any directory, but would need to modify the script)
2. Download the teams_rooms.xlsx input file and customize:
    -Column 1 is a friendly name for each account. It is not used in the XML, or to name output files. It's just a reference.  
    -Column 2 is the resource account UPN. This is used for the <SkypeSignInAddress> and <ExchangeAddress> fields in the XML.
    -Column 3 is the resource account password. This is used for the <Password> field in the XML
    -Column 4 is the MTR IP address or hostname. This column is only used to name the output XML files. It is not used in the XML.
    -Column 5 is an NTP address. This is not used at all in the MTR_Create_XML.ps1 file, but is referenced below for the MTR config script. 


