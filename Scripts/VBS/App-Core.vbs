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

Dim php73Directory, phpavEngineDirectory, whoamiOutput, strHRAVpassword, storedPassword, configFile, colAccounts, objUser, _
 objUser2, objGroup, ouser, errorMessage, emailContent, emailSubject,  objUserFlags, objPasswordExpirationFlag, _
 newKey1, newKey2, newKey3, newKey4, passwordFile, newPasswordFile, programFilesCheck, appdataFilesCheck, installationDirectory, _
 instHead, instMsg1, instMsg2, instMsg3, instMsg4, instMsg5, instMsg6, pfCopyResult, iW1Result, iW2Result, uCreated, instMsg7, _
 result0, key1, key2, key3, key4, uCheck, pfCheck, oLNK
'--------------------------------------------------

'--------------------------------------------------
'Set variables for the session.
phpavEngineDirectory = scriptsDirectory & "\PHP\PHP-AV\"
php73Directory = "PHP\7.3.8\php.exe"
passwordFile = cacheDirectory & appNAme & "_Keys.vbs"
InstallationDirectory = "C:\Program Files\" 
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
  objShell.Run SanitizeFolder(fullScriptName), 0, TRUE
  DieGracefully 0, "", TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated as HRAV user.
'Returns FALSE if the application is not elevated as HRAV user.
Function isUserHRAV()
  On Error Resume Next
  whoamiOutput = Sanitize(SystemBootstrap("whoami", "", FALSE))
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If Err.number = 0 And Trim(Replace(Replace(whoamiOutput, Chr(10), ""), Chr(13), "")) = strHRAVUserName Then 
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
  DieGracefully 1, "", TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a Warning.mail file and send it using sendmail via the Bootstrap to Sendmail at the bottom.
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
  Bootstrap "Sendmail\sendmail.exe", "", TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to generate a random password for the HRAV user.
'Generates a new random password and saves it to the \Config\ folder.
Function generatePassword()
  Randomize
  newKey1 = Rnd * 100000000000000000
  newKey2 = Rnd * 100000000000
  newKey3 = Rnd * 1000000 
  newKey4 = RandomString(4)
  generatePassword = Trim(newKey4 & Int((newKey1 - newKey2 + 1) * newKey3 + newKey2))
  If objFSO.FileExists(passwordFile) Then
    objFSO.DeleteFile(passwordFile)
  End If
  Set newPasswordFile = objFSO.CreateTextFile(passwordFile, TRUE)
  newPasswordFile.WriteLine("Option Explicit" & vbNewLine & "Dim key1, key2, key3, key4" & vbNewLine & _
   "key1 = " & newKey1 & vbNewLine & "key2 = " & newKey2 & vbNewLine & "key3 = " & vbNewLine & newKey3 & vbNewLine & "key4 = " & newKey4)
  newPasswordFile.Close
  If Not objFSO.FileExists(passwordFile) Then
    DieGracefully 2, "Could not generate a new password!"
  Else
    createLog(appName & "-" & sesID & ", Generated a new password on " & humanDateTime & "!")
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to retrieve and verify the HRAV user password.
'Returns the password as calculated from the keys in the \Config\ folder.
Function verifyPassword()
  verifyPassword = Trim(key4 & Int((key1 - key2 + 1) * key3 + key2))
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to change the HRAV user password.
'Part of HRAV's security is to constantly rotate the HRAV user password.
'If the password is stale we can assume something is wrong.
Function changePassword()
  Set objUser3 = GetObject("WinNT://" & Sanitize(strComputerName) & "/" & Sanitize(strHRAVUserName) & ", user")
  'YOU ARE HERE!!!'
  objUser3.SetPassword verifyPassword
  objUser3.SetInfo 
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
	On Error Resume Next
	Set objUser = GetObject("WinNT://" & strComputerName & "/" & strHRAVUserName & ",user")
	If Err.Number <> 0 Then
	  'If the user account does not exist, create it.
	  objShell.Run "NET USER "&strHRAVUserName&" PASSWORD /ADD " _
	  & "/ACTIVE:YES /COMMENT:""HR-AV Anti-Virus User Account"" /FULLNAME:" _
	  & strHRAVUserName &" /expires:never", 0, True
	 End If
	  On Error Resume Next 
	  Set objUser = GetObject("WinNT://" & strComputerName & "/" & strHRAVUserName & ",user")
	If Err.Number = 0 Then
	  'Connect newly created user to the Administrators group.
	  Set objGroup = GetObject("WinNT://" & strComputerName & "/Administrators")
	  'Add the user account to the group
	  On Error Resume Next
	  objGroup.Add(objUser.ADsPath)
	  WScript.sleep 600
	  objGroup.Add(objUser.ADsPath)
	End If
	'Set Account password to never expire
	'This is done externally due to NET USER limitations
	Const ufDONT_EXPIRE_PASSWD = &H10000
	objUserFlags = objUser.Get("UserFlags")
	If (objUserFlags And ufDONT_EXPIRE_PASSWD) = 0 Then
	  objUserFlags = objUserFlags Or ufDONT_EXPIRE_PASSWD
	  objUser.Put "UserFlags", objUserFlags
	  objUser.SetInfo
	End If
  Err.clear
  createHRAVUser = checkHRAVUser()
End Function
'--------------------------------------------------

'https://www.vistax64.com/threads/vbs-script-to-create-start-programs-menu-item.240145/
Function createStartMenuShortcut()
  Set oLNK = objShell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" & appName & ".lnk")
  With oLNK
  .TargetPath = InstallationDirectory & appName & "\" & appName & ".hta"
  .IconLocation = InstallationDirectory & appName & "\Resources\" & appName & ".ico"
  .Save
  End With
  oLNK = NULL
End Function
'--------------------------------------------------
'A function to detect if the application has been installed to \Program Files or not.
Function isInProgramFiles()
  isInProgramFiles = FALSE
  programFilesCheck = InStr(fullScriptName, "Program Files")
  appdataFilesCheck = InStr(fullScriptName, "AppData")
  If programFilesCheck > 0 Then
    isInProgramFiles = TRUE
  End If
  If appdataFilesCheck > 0 Then
    isInProgramFiles = FALSE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to verify the existing installation and install/correct a missing/broken installation.
Function verifyInstallation()
  verifyInstallation = FALSE
  pfCheck = isInProgramFiles
  uCheck = checkHRAVUser
  'Detect if running from Program Files. If not, fire "Installation Wizard 1."
  If pfCheck = FALSE Then
    iW1Result = installationWizard1()
    If iW1Result = 2 Then
      DieGracefully 210, "Operation cancelled by user!", FALSE
    End If
    DieGracefully 219, "Restart Required!", TRUE
  End If
  'Detect if running from Program Files but without an HRAV user. If so, fire "Installation Wizard 2."
  If pfCheck = TRUE And checkHRAVUser = FALSE Then
    iW2Result = installationWizard2()
    If iW2Result = 2 Then
      DieGracefully 220, "Operation cancelled by user!", FALSE
    End If
    DieGracefully 229, "Restart Required!", TRUE
  End If 
  'Detect if runing from Program Files and an HRAV user exists, signifying a valid installation environment.
  If pfCheck = TRUE And checkHRAVUser = TRUE Then
    verifyInstallation = TRUE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to copy all files to a dedicated Program Files directory.
Function copyToProgramFiles()
  If Not objFSO.FolderExists(InstallationDirectory) Then
    objFSO.CreateFolder(InstallationDirectory)
  End If
  If objFSO.FolderExists(InstallationDirectory) Then
    createLog("Created a folder at: " & InstallationDirectory)
    objFSO.GetFolder(currentDirectory).Copy InstallationDirectory
    If objFSO.FileExists(InstallationDirectory & appName & "\" & appName & ".hta") Then
      createLog("Copied files to: " & InstallationDirectory)
      copyToProgramFiles = TRUE
    Else
      DieGracefully 207, "Could not copy files to: " & InstallationDirectory, FALSE
    End If
  Else
    DieGracefully 206, "Could not create a folder at: " & InstallationDirectory, FALSE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to delete all original files from wherever they came from.
'In the case that this application is "compiled" with VBSEdit into a standalone .exe
'these files are extracted to .AppData. That's why the function is called "deleteFromAppData."
'If files are not found in .AppData this function returns FALSE and does nothing.
Function deleteFromAppData()
  'NEED TO MAKE THIS'
  deleteFromAppData = TRUE
End Function
'--------------------------------------------------

'--------------------------------------------------
Function installationWizard1()
  If Not isUserAdmin Then 
    restartAsAdmin()
  End If
  pfCopyResult = FALSE
  result0 = FALSE
  instHead = "Installation Wizard"
  instMsg1 = "Welcome to the " & appName & " Installation wizard!" & vbCRLF & vbCRLF & _
   "This wizard will guide you through the the installation process." & _
   "At any time you can click the cancel button to stop the installation process."
  instMsg2 = "Before we continue, we want you to know that this software is 100% free and open-source licensed to you under GNU GPLv3v (gnu.org/licenses/gpl-3.0.en.html)." & vbCRLF & vbCRLF & _
   "At HonestRepair, we beleive in the GNU definition of free software. Free software in this context doesn't mean  'Free beer.'" & vbCRLF & vbCRLF & _
   "It means 'Free' as in you have the 'Freedom' to modify, distribute, and understand the software you use." & vbCRLF & vbCRLF & _
   "To view or download the source code for this application, please visit our website (HonestRepair.net) or the official HR-AV Github repository (github.com/zelon88/HR-AV)."
  instMsg3 = "By clicking 'Ok' below, you agree that you understand your rights as a consumer of free software, and that any redistributed forms of this application must also be licensed under GNU GPLv3 to protect the rights of everyone."
  instMsg4 = "By clicking 'Ok' below, " & appName & " files will be installed to the following directory: " & vbCRLF & vbCRLF & installationDirectory
  instMsg5 = "Successfully copied " & appName & " files to " & installationDirectory & " on " & humanDateTime & "! The installation will now continue using the copied version of " & appName & "."
  instMsg6 = "Could not copy files to: " & installationDirectory & "!"
  instMsg7 = "Restarting from new installation directory."
  result0 = PrintGracefully(instHead, instMsg1, "vbOkCancel")
  result0 = PrintGracefully(instHead, instMsg2, "vbOkCancel")
  result0 = PrintGracefully(instHead, instMsg3, "vbOkCancel")
  result0 = PrintGracefully(instHead, instMsg4, "vbOkCancel")
  pfCopyResult = copyToProgramFiles()
  If dontContinue = FALSE And result0 <> 2 And result0 <> 3 And pfCopyResult = TRUE And objFSO.FileExists(InstallationDirectory & appName & "\" & appName & ".hta") Then
    PrintGracefully instHead, instMsg5, "vbOkOnly"
    objShell.Run """C:\Program Files\HR-AV\Scripts\VBS\Restart.vbs"""
    DieGracefully 0, instMsg7, TRUE 
  Else
    DieGracefully 201, instMsg6, FALSE 
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
Function installationWizard2()
  Dim inst2Head, inst2Msg1, inst2Msg2, inst2Msg3, inst2Msg4, instMsg5, instMsg6, userCreateResult, password
  pfCopyResult = FALSE
  inst2Head = "Installation Wizard (continued...)"
  inst2Msg1 = "You have successfully installed " & appName & " application files onto your computer!" & vbCRLF & vbCRLF & _
   "There's just a few more things to do to get your new AV solution ready. Specifically, we need to create an admin account for " & appName & " to use for system-wide malware fighting superpowers!"
  inst2Msg2 = "By clicking 'Ok' above, " & appName & " will create a new user named '" & strHRAVUserName & "' and add that user to the local Administrators group."
  inst2Msg3 = "Could not create an '" & strHRAVUserName & "' user!"
  inst2Msg4 = "Restarting using the newly created '" & strHRAVUserName & "' user."
  PrintGracefully inst2Head, inst2Msg1, "vbOkCancel" 
  PrintGracefully inst2Head, inst2Msg2, "vbOkCancel" 
  userCreateResult = createHRAVUser()
  createStartMenuShortcut
  password = verifyPassword()
  If userCreateResult = TRUE Then
    PrintGracefully instHead, instMsg5, "vbOkOnly" 
    Bootstrap "PAExec\paexec.exe", "-u:" & strHRAVUserName & " -p:" & password & " " & installationDirectory & "\HR-AV.hta", TRUE 
    DieGracefully 0, inst2Msg4, TRUE 
  Else
    DieGracefully 202, inst2Msg3, FALSE
  End If
  generatePassword = NULL
  password = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function shut down the machine when triggered.
Function killWorkstation()
  shell.Run "C:\Windows\System32\shutdown.exe /s /f /t 0", 0, false
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to execute VBS scripts in the context and scope of the running script. Works just like a PHP include().
'https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Sub Include(pathToVBS)
  Set objVBSFile = objFSO.OpenTextFile(pathToVBS, 1)
  ExecuteGlobal objVBSFile.ReadAll
  objVBSFile.Close
  Set objVBSFile = NULL
End Sub
'--------------------------------------------------
