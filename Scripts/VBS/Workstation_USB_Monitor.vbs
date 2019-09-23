'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 9/19/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'File Name: Workstation_USB_Monitor.vbs
'Version: v1.8, 9/19/2019

'This file was modified from https://github.com/zelon88/Workstation_USB_Monitor
'For use in the HR-AV application.

Option Explicit
dim strComputer, objWMIService, objNet, objFSO, colMonitoredEvents, objShell, wmiServices, wmiDiskDrives, wmiDiskDrive, _
 query, wmiDiskPartitions, wmiDiskPartition, wmiLogicalDisks, wmiLogicalDisk, return1, return2, objLatestEvent, param1, _
 param2, param3, param4, param5, usbOnly, silentOnly, arg, userName, hostName, mailFile, mFile, mailData, strComputerName, _
 resultCounter, strSafeDate, strSafeTime, strDateTime, strLogFilePath, strLogFileName, returnData, objLogFile, emailDisable, _
 logDisable, strSafeTimeRAW, strSafeTimeDIFF, strSafeTimeLAST, companyName, companyAbbr, fromEmail, toemail, _
 sendmailPath, logPath companyDomain, objVBSFile, pathToVBS, enableEmail

'Define variables for the session
' ----------
'The "appPath" is the full absolute path for the script directory, with trailing slash.
appPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"
'Perform a quick sanity check to be sure the value of "appPath" won't cause problems.
If appPath = NULL or appPath = FALSE Then
  appPath = ""
End If
sendmailPath =  Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Binaries\Sendmail\sendmail.exe "
logPath = "\\server\Logs" 
strComputer = "." 
resultCounter = 0
param1 = ""
param2 = ""
strSafeTimeRAW = 0
strSafeTimeDIFF = 0
strSafeTimeLAST = 0
usbOnly = false
silentOnly = false
emailDisable = false
logDisable = false
guiDisable = false
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colMonitoredEvents = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE Targetinstance " & _ 
 "ISA 'Win32_PNPEntity' and TargetInstance.DeviceId like '%USBStor%'") 
Set wmiServices = GetObject ("winmgmts:{impersonationLevel=Impersonate}!//" & strComputer)
Set arg = WScript.Arguments
Set objNet = CreateObject("Wscript.Network") 
Set objShell = WScript.CreateObject("WScript.Shell")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
userName = objNet.Username 
hostName = objNet.Computername
fromEmail = hostName & "@" & companyDomain
mailFile = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Temp\Workstation_USB_Monitor_Warning.mail"

'Execute the Service_Config.vbs configuration script containing required variables.
'https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Set objVBSFile = objFSO.OpenTextFile(Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Config\Service_Config.vbs", 1)
ExecuteGlobal objVBSFile.ReadAll
objVBSFile.Close
objVBSFile = NULL

'Retrieve the specified arguments.
If (arg.Count > 0) Then
   param1 = arg(0)
End If
If (arg.Count > 1) Then
   param2 = arg(1)
End If
If (arg.Count > 2) Then
   param3 = arg(2)
End If
If (arg.Count > 3) Then
   param4 = arg(3)
End If

'If the -u or --usb arguments are set we only retrieve data for USB storage devices.
If (param1 = "-u" Or param1 = "--usb") Then
  usbOnly = TRUE
End If
If (param2 = "-u" Or param2 = "--usb") Then
  usbOnly = TRUE
End If
If (param3 = "-u" Or param3 = "--usb") Then
  usbOnly = TRUE
End If
If (param4 = "-u" Or param4 = "--usb") Then
  usbOnly = TRUE
End If

'If the -e or --email arguments are set we disable the notification email.
If (param1 = "-e" Or param1 = "--email") Then
  enableEmail = FALSE
End If
If (param2 = "-e" Or param2 = "--email") Then
  enableEmail = FALSE
End If
If (param3 = "-e" Or param3 = "--email") Then
  enableEmail = FALSE
End If
If (param4 = "-e" Or param4 = "--email") Then
  enableEmail = FALSE
End If

'If the -l or --log arguments are set we disable the logfile.
If (param1 = "-l" Or param1 = "--log") Then
  logDisable = TRUE
End If
If (param2 = "-l" Or param2 = "--log") Then
  logDisable = TRUE
End If
If (param3 = "-l" Or param3 = "--log") Then
  logDisable = TRUE
End If
If (param4 = "-l" Or param4 = "--log") Then
  logDisable = TRUE
End If

'If the -s or --silent arguments are set we disable all echo's within the script.
If (param1 = "-s" Or param1 = "--silent") Then
  silentOnly = TRUE
End If
If (param2 = "-s" Or param2 = "--silent") Then
  silentOnly = TRUE
End If
If (param3 = "-s" Or param3 = "--silent") Then
  silentOnly = TRUE
End If
If (param4 = "-s" Or param4 = "--silent") Then
  silentOnly = TRUE
End If

'A funciton for running SendMail.
Function SendEmail() 
  objShell.run sendmailPath & " " & mailFile
End Function

'Perform the loop that checks for new devices.
Do While True
  Set objLatestEvent = colMonitoredEvents.NextEvent 
  'If USB only is set by the -u or --usb argument we run the top query. If -u or --usb is not set we run the bottom query.
  if (usbOnly = true) Then
    Set wmiDiskDrives = wmiServices.ExecQuery ( "SELECT Caption, DeviceID FROM Win32_DiskDrive WHERE InterfaceType = 'USB'")
  End If
  if (usbOnly = false) Then
    Set wmiDiskDrives = wmiServices.ExecQuery ( "SELECT Caption, DeviceID FROM Win32_DiskDrive")
  End If
  If (resultCounter = 0) Then
    'Use the disk drive device id to find associated information about the device.
    For Each wmiDiskDrive In wmiDiskDrives
      query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
      Set wmiDiskPartitions = wmiServices.ExecQuery(query)
      resultCounter = resultCounter + 1
      'Use partition device id to find logical disk.
      For Each wmiDiskPartition In wmiDiskPartitions
        Set wmiLogicalDisks = wmiServices.ExecQuery ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & wmiDiskPartition.DeviceID & _
         "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 
        return1 = ""
        'Build the return data
        For Each wmiLogicalDisk In wmiLogicalDisks
          return1 = "Device Type: " & wmiDiskDrive.Caption & ", " & _
           vbNewLine & "Device ID: " & wmiDiskPartition.DeviceID & ", " & _
           vbNewLine & "Logical Volume: " & wmiLogicalDisk.DeviceID & _
           vbNewLine & vbNewLine
          return2 = return1 & return2
        Next 
      Next
    Next
  End IF
  'Detection starts here and stops here when listening for more devices. (Be careful what goes near here).
  returnData = Notify()
  If (logDisable = false) Then 
    CreateLog returnData, logPath
  End If
Loop

'A function to format the notification email and notify the user.
function Notify()
  If (resultCounter > 0) Then
    resultCounter = resultCounter - 1
  End If
  If (resultCounter = 0) Then
    'Prepare the notification email and popup.
    Set mFile = objFSO.CreateTextFile(mailFile, true, false)  
    mFile.Write "To: " & toEmail & vbNewLine & "From: " & fromEmail & vbNewLine & "Subject: " & companyAbbr & " New Storage Device Connected!!!" & _
     vbNewLine & "This is an automatic email from the " & companyName & " Network to notify you that a new storage device was detected on a domain workstation." & _
     vbNewLine & vbNewLine & _
     "Please review the information below to verify that the connected device is not a threat." & _
     vbNewLine & vbNewLine & _
     "DEVICE DETAILS: " & _
     vbNewLine & vbNewLine & _
     "Workstation: " & hostName & ", " & _
     vbNewLine & "Username: " & userName & ", " & _
     vbNewLine & vbNewLine & "Detected Devices: " & _
     vbNewLine &vbNewLine & return2 & vbNewLine & _
     "This check was generated by " & strComputerName & " and is run in the background upon user logon." & _
     vbNewLine & vbNewLine & _
     "Script: """& companyAbbr & " Workstation_USB_Monitor.vbs""" 
    mFile.Close
    strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    strSafeTimeRAW = strSafeTime
    strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
    If (enableEmail = TRUE And strSafeTimeDIFF > 3) Then
      SendEmail
    End If
    'Display results if the silent argument is not set.
    If (silentOnly = FALSE And strSafeTimeDIFF > 3) Then
      mailData = "Devices Detected: " & vbNewLine & vbNewLine & return2
      MsgBox mailData, vbOKOnly, "Workstation USB Monitor"
    End If
    'Reset the outputs for the next iteration of the loop above. (MUST BE DONE!!! This was the source of a lot of debugging.)
    Notify = return2
    return2 = ""
    return1 = ""
  End If
End Function

'A function to create a log file.
Function CreateLog(strEventInfo, strLogFilePath)
  If Not (strEventInfo = "") Then
    'Logfile related variables are defined at log creation time for accurate time reporting.
    strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
    strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    strSafeTimeRAW = strSafeTime
    strSafeTimeDIFF = strSafeTime - strSafeTimeLAST
    'Some machines with lower performance may create multiple logfiles in rapid succession. This check ensures logs aren't duplicated.
    If (strSafeTimeDIFF > 3) Then
      strDateTime = strSafeDate & "-" & strSafeTime
      strLogFileName = strLogFilePath & "\" & userName & "-" & strDateTime & "-workstation_usb_monitor.txt"
      Set objLogFile = objFSO.CreateTextFile(strLogFileName, true, false)
      objLogFile.WriteLine(strEventInfo)
      objLogFile.Close
    End IF
    strSafeTimeLAST = strSafeTimeRAW
  End If
End Function
