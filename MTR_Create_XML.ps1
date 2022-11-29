#Script to read xls file, grab cell data and save as XML file.

#xlsx file to use
$file = "C:\temp\mtr_xml\teams_rooms.xlsx"

#specify which tab in xlsx file to grab data from
$sheetName = "ACCTS"

#open xlsx file in background to grab data points
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false
$rowMax = ($sheet.UsedRange.Rows).count

$rowmtr,$colmtr = 1,1
$rowusername,$colusername = 1,2
$rowpassword,$colpassword = 1,3
$rowdeviceid,$coldeviceid = 1,4

for ($i=1; $i -le $rowMax-1; $i++)

{
$mtr = $sheet.Cells.Item($rowroom+$i,$colroom).text
$username = $sheet.Cells.Item($rowusername+$i,$colusername).text
$password = $sheet.Cells.Item($rowpassword+$i,$colpassword).text
$deviceid = $sheet.Cells.Item($rowipaddress+$i,$colipaddress).text

$String2 = @"
<SkypeSettings>
  <AutoScreenShare>1</AutoScreenShare>
  <HideMeetingName>1</HideMeetingName>
  <AutoExitMeetingEnabled>true</AutoExitMeetingEnabled>
  <AudioRenderDefaultDeviceVolume>70</AudioRenderDefaultDeviceVolume>
  <AudioRenderCommunicationDeviceVolume>30</AudioRenderCommunicationDeviceVolume>
  <UserAccount>
    <SkypeSignInAddress>$username</SkypeSignInAddress>
    <ExchangeAddress>$username</ExchangeAddress>
    <DomainUsername>domain\username</DomainUsername>
    <Password>$password</Password>
    <ConfigureDomain>domain1, domain2</ConfigureDomain>
    <ModernAuthEnabled>true</ModernAuthEnabled>
  </UserAccount>
  <TeamsMeetingsEnabled>true</TeamsMeetingsEnabled>
  <SfbMeetingEnabled>true</SfbMeetingEnabled>
  <IsTeamsDefaultClient>true</IsTeamsDefaultClient>
  <WebExMeetingsEnabled>true</WebExMeetingsEnabled>
  <ZoomMeetingsEnabled>true</ZoomMeetingsEnabled>
  <UseCustomInfoForThirdPartyMeetings>true</UseCustomInfoForThirdPartyMeetings>
  <CustomDisplayNameForThirdPartyMeetings>guestname</CustomDisplayNameForThirdPartyMeetings>
  <CustomDisplayEmailForThirdPartyMeetings>guest@microsoft.com</CustomDisplayEmailForThirdPartyMeetings>
  <BluetoothAdvertisementEnabled>false</BluetoothAdvertisementEnabled>
  <AutoAcceptProximateMeetingInvitations>true</AutoAcceptProximateMeetingInvitations>
  <CortanaWakewordEnabled>true</CortanaWakewordEnabled>
  <DualScreenMode>0</DualScreenMode>
  <DuplicateIngestDefault>true</DuplicateIngestDefault>
  <DisableTeamsAudioSharing>true</DisableTeamsAudioSharing>
  <FrontRowEnabled>true</FrontRowEnabled>
  <FrontRowVideoSize>medium</FrontRowVideoSize>
  <SingleFoRDefaultContentLayout>1</SingleFoRDefaultContentLayout>
  <DefaultFoRExperience>0</DefaultFoRExperience>
  <EnablePublicPreview>false</EnablePublicPreview>
  <NoiseSuppressionDefault>1</NoiseSuppressionDefault>
  <SendLogs>
    <EmailAddressForLogsAndFeedback>username@microsoft.com</EmailAddressForLogsAndFeedback>
    <SendLogsAndFeedback>True</SendLogsAndFeedback>
  </SendLogs>
  <Devices>
    <MicrophoneForCommunication>Device1</MicrophoneForCommunication>
    <SpeakerForCommunication>DeviceX</SpeakerForCommunication>
    <DefaultSpeaker>DeviceX</DefaultSpeaker>
    <ContentCameraId>Camera1</ContentCameraId>
    <ContentCameraEnhancement>true</ContentCameraEnhancement>
    <ContentCameraInverted>false</ContentCameraInverted>
  </Devices>
  <Theming>
       <ThemeName>Custom</ThemeName>
       <CustomThemeImageUrl>file name</CustomThemeImageUrl>
       <CustomThemeColor>
            <RedComponent>100</RedComponent>
            <GreenComponent>100</GreenComponent>
            <BlueComponent>100</BlueComponent>
       </CustomThemeColor>
  </Theming>
  <CoordinatedMeetings enabled="true">
    <TrustedAccounts>username1@microsoft.com,username2@contoso.com</TrustedAccounts>
    <Settings>
      <Audio default="true" enabled="true"/>
      <Video default="true" enabled="true"/>
      <Whiteboard default="true" enabled="true"/>
    </Settings>
  </CoordinatedMeetings>
  <EnableResolutionAndScalingSetting>true</EnableResolutionAndScalingSetting> 
  <MainFoRDisplay> 
      <MainFoRDisplayResolution>1920,1080</MainFoRDisplayResolution> 
      <MainFoRDisplayScaling>100</MainFoRDisplayScaling> 
  </MainFoRDisplay> 
  <ExtendedFoRDisplay> 
      <ExtendedFoRDisplayResolution>1920,1080</ExtendedFoRDisplayResolution> 
      <ExtendedFoRDisplayScaling>100</ExtendedFoRDisplayScaling> 
  </ExtendedFoRDisplay>  
  <EnableDeviceEndToEndEncryption>false</EnableDeviceEndToEndEncryption>
  <SplitVideoLayoutsDisabled>false</SplitVideoLayoutsDisabled>
</SkypeSettings> 
"@>c:\temp\mtr_xml\xml_files\$deviceid.xml
}
$objExcel.quit()

