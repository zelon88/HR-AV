'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/24/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'Portions of the UI-Core.vbs file are licensed under the Microsoft Limited Public License.
'Copies of all applicable software licenses can be found in the "Documentation" directory.

'This file contains popular functions that are likely to be used throught the entire operation.
'Note that some of the real-time protection files define these variables again for robustness.

'--------------------------------------------------
Option Explicit

Dim php73Directory, phpavEngineDirectory, whoamiOutput, strHRAVpassword, storedPassword, configFile, colAccounts, objUser, oRE1, _
 objUser2, objGroup, ouser, errorMessage, emailContent, emailSubject, strToClean, objRegExp, outputStr, _
 objUserFlags, objPasswordExpirationFlag
'--------------------------------------------------

'--------------------------------------------------
'Set variables for the session.
phpavEngineDirectory = scriptsDirectory & "\PHP\PHP-AV\"
php73Directory = "PHP\7.3.8\php.exe"
'--------------------------------------------------

'--------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated as admin user.
'Returns FALSE if the application is not elevated as admin user.
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
'--------------------------------------------------

'--------------------------------------------------
'A function to restart the script with admin priviledges if required.
Function restartAsAdmin()
    oShell2.ShellExecute "wscript.exe", Chr(34) & SanitizeFolder(fullScriptName) & Chr(34), "", "runas", 1
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated as HRAV user.
'Returns FALSE if the application is not elevated as HRAV user.
Function isUserHRAV()
  On Error Resume Next
  whoamiOutput = Sanitize(SystemBootstrap ("whoami", "", FALSE))
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If Err.number = 0 And whoamiOutput = strHRAVUserName Then 
    isUserHRAV = TRUE
  Else
    isUserHRAV = FALSE
  End If
  Err.Clear
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to restart the script with admin priviledges if required.
Function restartAsHRAV(strHRAVPassword)
    Bootstrap "PAExec\paexec.exe", "-u:" & Sanitize(strHRAVUserName) & " -p:" & Sanitize(strHRAVPassword) & " " & SanitizeFolder(fullScriptName), FALSE
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
Function createLog(strEventInfo)
  If Not strEventInfo = "" Then
    Set objLogFile = oFSO.CreateTextFile(logFileName, True)
    objLogFile.WriteLine(Sanitize(strEventInfo))
    objLogFile.Close
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a Warning.mail file. Use to prepare an email before calling sendEmail().
Function createEmail(emailSubject, emailContent)
  If oFSO.FileExists(mailFile) Then
    oFSO.DeleteFile(mailFile)
  End If
  If Not oFSO.FileExists(mailFile) Then
    oFSO.CreateTextFile(mailFile)
  End If
  Set oFile = oFSO.CreateTextFile(mailFile, True)
  oFile.Write "To: " & Sanitize(toEmail) & vbNewLine & "From: " & Sanitize(strComputerName) & "@" & Sanitize(companyDomain) & vbNewLine & _
   "Subject: " & Sanitize(companyAbbr & " " & appName) & " Warning!!! " & Sanitize(emailSubject) & vbNewLine & _
   "This is an automatic email from the " & Sanitize(companyName) & " Network to notify you that: " & Sanitize(emailContent) & _
   vbNewLine & vbNewLine & "Please log-in and verify that the equipment listed below is secure." & vbNewLine & _
   vbNewLine & "USER NAME: " & Sanitize(strUserName) & vbNewLine & "WORKSTATION: " & Sanitize(strComputerName) & vbNewLine & _
   "This check was generated by " & Sanitize(strComputerName) & "."
  oFile.close
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for running SendMail to send a prepared Warning.mail email message.
Function sendEmail() 
  oShell.run "c:\Windows\System32\cmd.exe /c sendmail.exe " & SanitizeFolder(mailFile), 0, TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to generate a random password for the HRAV user.
Function generatePassword()

End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to check and verify the HRAV user password.
Function verifyPassword()

End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to change the HRAV user password.
'Part of HRAV's security is to constantly rotate the HRAV user password.
'If the password is stale we can assume something is wrong.
Function changePassword()

End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to check if the HRAV user exists.
Function checkHRAVUser()
  On Error Resume Next
  Set ouser = GetObject("WinNT://" & Sanitize(strComputerName) & "/" & Sanitize(strHRAVUserName) & ",user")
  checkHRAVUser = FALSE
  If Err.number = 0 Then
    checkHRAVUser = TRUE
  End If
  Err.clear
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create an HRAV user and add it to the Administrators group. 
'By doing this we can run background services without asking the user for escalation all the time.
'Requires that the user confirm elevation the first time.
Function createHRAVUser()
  createHRAVUser = FALSE
  'Create the HRAV user on the local machine.
  Set colAccounts = GetObject("WinNT://" & Sanitize(strComputerName) & "")
  Set objGroup = GetObject("WinNT://" & Sanitize(strComputerName) & "/Administrators,group") 
  Set objUser2 = GetObject("WinNT://" & Sanitize(strComputerName) & "/" & Sanitize(strHRAVUserName) & ",user")
  Set objUser = colAccounts.Create("user", Sanitize(strHRAVUserName))  
  objUser.SetPassword generatePassword()
  objUser.SetInfo
  'Set the option "Password never expires."
  Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
  objUserFlags = objUser.Get("UserFlags")
  objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
  objUser.Put "userFlags", objPasswordExpirationFlag 
  objUser.SetInfo
  'Add the newly created HRAV user to the Administrators group.
  On Error Resume Next 
  objGroup.Add(objUser2.ADsPath) 
  Err.clear
  createHRAVUser = checkHRAVUser()
End Function

'--------------------------------------------------
Function verifyInstallation()
  verifyInstallation = FALSE
  createUserCheck = FALSE
  If checkHRAVUser = FALSE Then
    createUserCheck = createHRAVUser()
    MsgBox "There was a problem creating the HRAV user! This is usually because you do not have administrator permissions. Real-time protection", appName, vbCritical
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function shut down the machine when triggered.
Function killWorkstation()
     oShell.Run "C:\Windows\System32\shutdown.exe /s /f /t 0", 0, false
End Function
'--------------------------------------------------



