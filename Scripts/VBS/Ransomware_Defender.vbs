'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'File Name: Ransomware_Defender.vbs
'Version: v1.3.1, 8/23/2019

'This file was modified from https://github.com/zelon88/Ransomware_Defender
'For use in the HR-AV application.

Option Explicit

dim oShell, oShell2, oFSO, perimiterFile, perimiterFiles, perimiterCheck, perimiterFileHash, scriptName, tempFile, appPath, logPath, exe, cmdHardCodedHash, cmdDynamicHash, strComputerName, _
 strUserName, strSafeDate, strSafeTime, strDateTime, logFileName, strEventInfo, objLogFile, cmdHashCache, objCmdHashCache, dangerHashCache, tempDir, tempDir0, tempDir1, _
 dangerHashData, mailFile, objDangerHashCache, oFile, tempOutput, companyName, companyAbbr, companyDomain, toEmail, defaultPerimiterFile, tempData, _
 defaultPerimiterFileName, searchname1, folder, file, sourcefolder, targetFileName, enableEmail, pathToVBS, objVBSFile

  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The "scriptName" is the filename of this script.
  scriptName = "Ransomware_Defender.vbs"
  ' The "appPath" is the full absolute path for the script directory, with trailing slash.
  appPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"
  ' Perform a quick sanity check to be sure the value of "appPath" won't cause problems.
  If appPath = NULL or appPath = FALSE Then
    appPath = ""
  End If
  ' The "logPath" is the full absolute path for where network-wide logs are stored.
  logPath = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Logs\"
  ' The "mailFile" is the full absolute path to the location where a temporary email file will be generated.
  mailFile = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Temp\Ransomware_Defender_Warning.mail"
  ' The "defaultPerimiterFileName" is the master filename that all other perimiterfiles are copied from. It is located in the \Cache directory of the appPath.
  defaultPerimiterFileName = "Ransomware_Defender_Perimiter_File.dat"
  ' The "defaultPerimiterFile" is the master file that all other perimiter files are copied from. It is located in the \Cache directory of the appPath.
  defaultPerimiterFile = Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Cache\" & defaultPerimiterFileName
  ' You can change the values in the array below to add, remove, or rename perimiter files. 
  ' It's probably a good idea to randomize these values just in case ransomware authors build ransomware to avoid these defaults.
  perimiterFiles = Array("C:\Ransomware_Defender_Perimiter_File.dat", "C:\Program Files\Ransomware_Defender_Perimiter_File.dat", "C:\Users\Ransomware_Defender_Perimiter_File.dat", "C:\Windows\Ransomware_Defender_Perimiter_File.dat")
  ' The "perimiterFileHash" is a hard coded SHA256 hash that matches the "defaultPerimiterFile".
  perimiterFileHash = "cd 7e 60 a8 43 ca 66 50 6f 7e 48 10 3b 09 32 ec 6c 62 f1 81 1c 70 44 be ac 04 67 c6 8a d7 6e 18"
  ' ----------

'Set global variables for the session.
Set oShell = WScript.CreateObject("WScript.Shell")
Set oShell2 = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
tempDir0 = Left(appPath, InStrRev(appPath, "Scripts\VBS\"))
tempDir1 = tempDir0 & "\Cache"
tempDir = tempDir1 & "\" & strComputerName
tempFile = tempDir & "\" & strComputerName & "-Cache.dat"
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
logFileName = logPath & "\" & strComputerName & "-" & strDateTime & "-Ransomware_Defender.txt"

'A function to execute VBS scripts in the context and scope of the running script. Works just like a PHP include().
'https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Sub Include(pathToVBS)
  Set objVBSFile = oFSO.OpenTextFile(pathToVBS, 1)
  ExecuteGlobal objVBSFile.ReadAll
  objVBSFile.Close
  objVBSFile = NULL
End Sub

'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated.
'Returns FALSE if the application is not elevated.
Function isUserAdmin()
  On Error Resume Next
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If Err.number = 0 Then 
    isUserAdmin = TRUE
  Else
    isUserAdmin = FALSE
  End If
  Err.Clear
End Function

'A function to restart the script with admin priviledges if required.
Function restartAsAdmin()
    oShell2.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
End Function

'A function to verify the tempDir and clear the previous tempFile file and create a new one.
'Start by making C:\Program Files\Ransomware_Defender.
'Then make C:\Program Files\Ransomware_Defender\Cache.
'Then verify the cache files inside.
Function clearCache()
  If Not oFSO.FolderExists(tempDir0) Then
    oFSO.CreateFolder(tempDir0)
  End If
  If oFSO.FolderExists(tempDir0) Then
    If Not oFSO.FolderExists(tempDir1) Then
      oFSO.CreateFolder(tempDir1)
    End If
    If oFSO.FolderExists(tempDir1) Then
      If Not oFSO.FolderExists(tempDir) Then
        oFSO.CreateFolder(tempDir)
      End If
      If oFSO.FolderExists(tempDir) Then
        If oFSO.FileExists(tempFile) Then
          oFSO.DeleteFile(tempFile)
        End If
        If Not oFSO.FileExists(tempFile) Then
          oFSO.CreateTextFile(tempFile)
        End If
      End If
    End If
  End If
End Function

'A function to ensure a missing perimiter file hasn't been changed by malware.
'Returns TRUE when no matching files are found.
'Returns FALSE when a file with matching name is found.
Function searchForPerimiterFile(perimiterFile)
  searchForPerimiterFile = TRUE
  'Variable default is "Ransomware_Defender_Perimiter_File".
  searchname1 = Replace(defaultPerimiterFileName, ".dat", "")
  'Variable default is "Ransomware_Defender_Perimiter_File.dat".
  sourcefolder = Replace(perimiterFile, defaultPerimiterFileName, "")
  Set folder = oFSO.Getfolder(sourcefolder)  
  For Each file In folder.files
    targetFileName = oFSO.GetBasename(file)
    If InStr(lcase(targetFileName), lcase(searchname1)) > 0 Or InStr(lcase(searchname1), lcase(targetFileName)) > 0 Then
      searchForPerimiterFile = FALSE
    End If
  Next
End Function

'A function to detect each perimiter file on the system and ensure that it has not been altered.
'Returns TRUE when perimiter files exist and are valid.
Function verifyPerimiterFiles()
  perimiterCheck = TRUE
  verifyPerimiterFiles = TRUE
  For Each perimiterFile In perimiterFiles
    If Not oFSO.FileExists(perimiterFile) Then
      perimiterCheck = searchForPerimiterFile(perimiterFile)
      oFSO.Copyfile defaultPerimiterFile, perimiterFile
    Else
      oShell.run "c:\Windows\System32\cmd.exe /c CertUtil -hashfile """ & perimiterFile & """ SHA256 | find /i /v ""SHA256"" | find /i /v ""certutil"" > """ & tempFile & """", 0, TRUE
      Set tempOutput = oFSO.OpenTextFile(tempFile)
      If Not tempOutput.AtEndOfStream Then 
        tempData = tempOutput.ReadAll()
      End If
    End If
    If Trim(tempData) = Trim(perimiterFileHash) Or perimiterCheck = FALSE Then
      verifyPerimiterFiles = FALSE
      Exit For
    End If
  Next
End Function

'A function to create a log file.
Function createLog(strEventInfo)
  If Not strEventInfo = "" Then
    Set objLogFile = oFSO.CreateTextFile(logFileName, True)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function

'A function to create a Warning.mail file. Use to prepare an email before calling sendEmail().
Function createEmail()
  If oFSO.FileExists(mailFile) Then
    oFSO.DeleteFile(mailFile)
  End If
  If Not oFSO.FileExists(mailFile) Then
    oFSO.CreateTextFile(mailFile)
  End If
  Set oFile = oFSO.CreateTextFile(mailFile, True)
  oFile.Write "To: " & toEmail & vbNewLine & "From: " & strComputerName & "@" & companyDomain & vbNewLine & _
   "Subject: " & companyAbbr & " Ransomware Defender Warning!!!" & vbNewLine & _
   "This is an automatic email from the " & companyName & " Network to notify you that a workstation was disabled to prevent potential ransomware activity." & _
   vbNewLine & vbNewLine & "Please log-in and verify that the equipment listed below is secure." & vbNewLine & _
   vbNewLine & "USER NAME: " & strUserName & vbNewLine & "WORKSTATION: " & strComputerName & vbNewLine & _
   "This check was generated by " & strComputerName & "." & vbNewLine & vbNewLine & _
   "Script: """ & scriptName & """" 
  oFile.close
End Function

'A function for running SendMail to send a prepared Warning.mail email message.
Function sendEmail() 
  oShell.run "c:\Windows\System32\cmd.exe /c " & Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Binaries\Sendmail\sendmail.exe " & mailFile, 0, TRUE
End Function

'A function shut down the machine when triggered.
Function killWorkstation()
     oShell.Run "C:\Windows\System32\shutdown.exe /s /f /t 0", 0, false
End Function

'The main logic of the program which makes use of the code and functions above.
Include(Left(appPath, InStrRev(appPath, "Scripts\VBS\")) & "Config\Service_Config.vbs")

If isUserAdmin = TRUE Then
  clearCache()
  If verifyPerimiterFiles = FALSE Then
    createLog("The machine " & strComputerName & " has been disabled due to potential ransomware activity!")
    createEmail()
    sendEmail()
    killWorkstation()
  End If
Else
  restartAsAdmin()
End If
