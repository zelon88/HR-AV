'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 9/19/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'File Name: Accessibility_Defender.vbs
'Version: v1.0, 9/19/2019

'This file was modified from https://github.com/zelon88/Accessibility-Tools-utilmon-Defender
'For use in the HR-AV application.

Option Explicit

Dim oShell, oFSO, dangerousExes, exe, cmdHardCodedHash, cmdDynamicHash, strComputerName, strUserName, strLogFilePath, strSafeDate, _
 strSafeTime, strDateTime, strLogFileName, strEventInfo, objLogFile, cmdHashCache, objCmdHashCache, dangerHashCache, objVBSFile, _
 dangerHashData, mailFile, objDangerHashCache, oFile, toEmail, fromEmail, companyDomain, companyAbbr, companyName, appPath, pathToVBS, enableEmail

'The "appPath" is the full absolute path for the script directory, with trailing slash.
appPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"
'Perform a quick sanity check to be sure the value of "appPath" won't cause problems.
If appPath = NULL or appPath = FALSE Then
  appPath = ""
End If
Set oShell = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
dangerousExes = Array("Magnify.exe", "Narrator.exe", "osk.exe", "sapisvr.exe", "control.exe", "utilman.exe")
cmdHardCodedHash = "db 06 c3 53 49 64 e3 fc 79 d2 76 31 44 ba 53 74 2d 7f a2 50 ca 33 6f 4a 0f e7 24 b7 5a af f3 86"
cmdDynamicHash = ""
strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
strLogFilePath = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Logs\"
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
strLogFileName = strLogFilePath & "\" & strComputerName & "-" & strDateTime & "-Accessibility_Defender.txt"
cmdHashCache = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Cache\cmdHashCache.dat"
dangerHashCache = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Cache\dangerHashCache.dat"
mailFile = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Temp\Accessibility_Defender_Warning.mail"

'A function to execute VBS scripts in the context and scope of the running script. Works just like a PHP include().
'https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Sub Include(pathToVBS)
  Set objVBSFile = oFSO.OpenTextFile(pathToVBS, 1)
  ExecuteGlobal objVBSFile.ReadAll
  objVBSFile.Close
  Set objVBSFile = NULL
End Sub

'A function to clear the previous dangerCache and create a new one.
Function clearCache()
  If oFSO.FileExists(dangerHashCache) Then
    oFSO.DeleteFile(dangerHashCache)
  End If
  If Not oFSO.FileExists(dangerHashCache) Then
    oFSO.CreateTextFile(dangerHashCache)
  End If
End Function

'A function to create the CMD Hash Cache file.
Function getCmdHash()
  If oFSO.FileExists("C:\Windows\System32\cmd.exe") Then
    oShell.run "cmd /c CertUtil -hashfile ""C:\Windows\System32\cmd.exe"" SHA256 | find /i /v ""SHA256"" | find /i /v ""certutil"" > " & cmdHashCache, 0, TRUE
  End If
End Function

'A function to hash each of the hardcoded files and cache the value.
Function getDangerHash()
  For Each exe In dangerousExes
    If oFSO.FileExists("C:\Windows\System32\" & exe) Then
      oShell.run "cmd /c CertUtil -hashfile ""C:\Windows\System32\" & exe & """ SHA256 | find /i /v ""SHA256"" | find /i /v ""certutil"" >> " & dangerHashCache, 0, TRUE
    End If
  Next
End Function

'A function to read the CMD hash cache.
Function cmdHashData()
  If oFSO.FileExists(cmdHashCache) Then
    Set objCmdHashCache = oFSO.OpenTextFile(cmdHashCache)
    cmdHashData = objCmdHashCache.ReadAll()
    objCmdHashCache.close 
  End If
End Function

'A function to read the Danger hash cache and compare it to the CMD hash cache and hardcoded CMD hash.
Function hashMatch()
  hashMatch = FALSE
  If oFSO.FileExists(dangerHashCache) Then
    Set objDangerHashCache = oFSO.OpenTextFile(dangerHashCache)
    Do While Not objDangerHashCache.AtEndOfStream
      dangerHashData = objDangerHashCache.ReadLine()
      If dangerHashData = cmdHashData() Or dangerHashData = cmdHardCodedHash Then
        hashMatch = TRUE
      End If
    loop
    objDangerHashCache.close
  End If
End Function

'A function to create a log file.
Function createLog(strEventInfo)
  If Not (strEventInfo = "") Then
    Set objLogFile = oFSO.CreateTextFile(strLogFileName, True)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function

Function createEmail()
  If oFSO.FileExists(mailFile) Then
    oFSO.DeleteFile(mailFile)
  End If
  If Not oFSO.FileExists(mailFile) Then 
    oFSO.CreateTextFile(mailFile)
  End If
  Set oFile = oFSO.CreateTextFile(mailFile, True)
  oFile.Write "To: " & toEmail & vbNewLine & "From: " & strComputerName & "@" & companyDomain & vbNewLine & _
   "Subject: " & companyAbbr& " Accessibility Defender Warning!!!" & vbNewLine & _
      "This is an automatic email from the " & companyName & " Network to notify you that a workstation was defended from Accessibility Tools exploitation." & _
   vbNewLine & vbNewLine & "Please log-in and verify that the equipment listed below is secure." & vbNewLine & _
   vbNewLine & "USER NAME: " & strUserName & vbNewLine & "WORKSTATION: " & strComputerName & vbNewLine & _
   "This check was generated by " & strComputerName & "." & vbNewLine & vbNewLine & _
   "Script: ""Accessibility_Defender.vbs""" 
  oFile.close
End Function

'A function for running SendMail.
Function sendEmail() 
  oShell.run "cmd /c " & Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Binaries\Sendmail\sendmail.exe " & mailFile, 0, TRUE
End Function

'A function to display a warning message to the user and kill the machine after a specified time.
Function killWorkstation()
  oShell.Run "cmd /c C:\windows\system32\shutdown.exe", 0, false
End Function

Include(Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Config\Service_Config.vbs")

clearCache()
getCmdHash()
getDangerHash()
hashMatch()

If hashMatch Then
  createLog("The machine " & strComputerName & " just attempted to execute an Accessibility Tools exploitation!")
  If enableEmail Then
    createEmail()
    sendEmail()
  End If
  killWorkstation()
End If
