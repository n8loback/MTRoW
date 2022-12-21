#Update Host file to trust all
Set-Item WSMan:\localhost\Client\TrustedHosts -Value *

$file = "C:\temp\mtr_xml\teams_rooms.xlsx"

#specify which tab in xlsx file to grab data from
$sheetName = "ACCTS"

#open xlsx file in background to grab data points
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false
$rowMax = ($sheet.UsedRange.Rows).count
$rowroom,$colroom = 1,1
$rowusername,$colusername = 1,2
$rowpassword,$colpassword = 1,3
$rowdeviceid,$coldeviceid = 1,4
$rowntp,$colntp = 1,5

#Set Variables
$USER = "Admin"
$PWD = ConvertTo-SecureString "sfb" -AsPlainText -Force
$NEWPWD = ConvertTo-SecureString "NEWPWD" -AsPlainText -Force
$CRED = New-Object System.Management.Automation.PSCredential -ArgumentList ($USER, $PWD)
$NEWCRED = New-Object System.Management.Automation.PSCredential -ArgumentList ($USER, $NEWPWD)
$TARGET2 = "C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState\SkypeSettings.xml"
$FILE3 = "c:\temp\managedroomsinstaller.msi"
$TARGET3 = "c:\rigel\managedroomsinstaller.msi"
$FILE5 = "C:\temp\mtr_xml\custom.png"
$TARGET5 = "C:\Users\Skype\AppData\Local\Packages\Microsoft.SkypeRoomSystem_8wekyb3d8bbwe\LocalState\custom.png"

for($i=1; $i -le $rowMax-1; $i++)

{
$mtr = $sheet.Cells.Item($rowroom+$i,$colroom).text
$username = $sheet.Cells.Item($rowusername+$i,$colusername).text
$password = $sheet.Cells.Item($rowpassword+$i,$colpassword).text
$deviceid = $sheet.Cells.Item($rowdeviceid+$i,$coldeviceid).text
$ntp = $sheet.Cells.Item($rowntp+$i,$colntp).text


#Establish PSSession with Remote MTR
$Session = New-PSSession -ComputerName $deviceid -Credential $CRED

#Set MTR NTP
Invoke-Command -ComputerName $deviceid -Credential $CRED -ScriptBlock{
    Set-Location HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DateTime\Servers
    Set-ItemProperty . 0 $Using:ntp
    Set-ItemProperty . "(Default)" "0"
    Set-Location HKLM:\SYSTEM\CurrentControlSet\services\W32Time\Parameters
    Set-ItemProperty . NtpServer ($Using:ntp + ",0x9")
    Pop-Location
    Stop-Service w32time
    Start-Service w32time}

#Push XML
Write-Host "Copying XML to $mtr"
Copy-Item "c:\temp\mtr_xml\xml_files\$deviceid.xml" -ToSession $Session -Destination $TARGET2

#Push Background Image
Write-Host "Copying Custom MTR Background to $mtr"
Copy-Item $FILE5 -ToSession $Session -Destination $TARGET5

#Push and Install Microsoft MMR Agent
write-host "Copying and INstalling Microsoft MMR agent to $mtr"
Copy-Item $FILE3 -tosession $SESSION -Destination $TARGET3 

Invoke-Command -Session $SESSION -ScriptBlock{
     c:\rigel\managedroomsinstaller.msi}

start-sleep 120

#Change MTR local admin PW
Invoke-Command -ComputerName $deviceid -Credential $CRED -ScriptBlock{
Set-localuser -name $Using:USER -Password $Using:NEWPWD -AccountNeverExpires}

#Trigger Reboot
Invoke-Command -ComputerName $deviceid -Credential $NEWCRED -ScriptBlock{
shutdown /r /t 60
}
}

$objExcel.quit()

#Clear and Check HostFile
Set-Item WSMan:\localhost\Client\TrustedHosts -Value ""