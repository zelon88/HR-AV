'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 11/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'Large portions of code in this file were borrowed from the Microsoft TechNet website on 8/14/2019 
'in accordance with the Microsoft Limited Public License...
'https://gallery.technet.microsoft.com/scriptcenter/796bd584-0fdb-43bc-a5d2-aa5fc99a5e5d
'Copies of all applicable software licenses can be found in the "Documentation" directory.

'This file contains the popular functions and subroutines required for basic UI functionality.

'This file requires Config.vbs.

'--------------------------------------------------
Option Explicit

'Define global variables for the session.
Dim objFSO, strComputer, objWMIService, scriptsDirectory, binariesDirectory, humanDateTime, _
 colItems, objItem, intHorizontal, intVertical, nLeft, nTop, sItem, helpLocSetting, errorNumber, run, _
 version, currentDirectory, appName, developerName, developerURL, windowHeight, windowWidth, objSysInfo, _
 BinaryToRun, Command, tempDirectory, uiVersion, Async, error, requiredDir, requiredDirs, installationError, _
 dieOnInstallationError, cacheDirectory, pagesDirectory, realDirectory, vbsScriptsDirectory, dMenus, sMenuOpen, _
 hrefLocation, fullScriptName, arrFN, scriptName, oMatch, oMatches, shell, objWshNet, strNamespace, strHRAVUserName, _
 strHRAVGroupName, strCurrentUserName, oEL, oItem, objShell, objShellExec, tempFile, tempData, entry, strComputerName, file, _
 sBinaryToRun, sCommand, sAsync, stempFile, sasync1, srun, stempData, mediaPlayer, pathToMedia, mediaDirectory, realTimeCoreFile, _
 errorMessage, sCommLine, dProcess, quietly, windowNote, strEventInfo, logFilePath, objLogFile, humanDate, logDate, resourcesDirectory, _
 logDateTime, logTime, charArr, tmpChar, charArr2, tmpChar2, outputStr1, logsDirectory, sesID, rStr, rStrLen, i1, reportsDirectory, objVBSFile, _
 cantError, sProcName, oWMISrvc, Timesec, dontContinue, pathToVBS, typeMsg, humanTime, message, oRE, RTPCacheFile1, oRTPCacheFile1, RTPCacheFile2, _
 requiredCacheFile, requiredCacheFiles, oRTPCacheFile2, exceptionDirectory, excepptionFile, startup

'--------------------------------------------------
'UI Related Variables.
Const sMenuItems = "File,Settings,Help" 
Const sFile = "Exit" 
Const sSettings = "View Settings"
Const sHelp = "Help, About" 
Const sHTML = "&nbsp;&nbsp;&nbsp;#sItem#&nbsp;&nbsp;&nbsp;" 
Const Letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
startup = FALSE
'Frequently Used Objects.
Set objShell = CreateObject("WScript.Shell")
Set shell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("WinNTSystemInfo")
Set objWshNet = CreateObject("WScript.Network")
Set oWMISrvc = GetObject("winmgmts:")
'Time Related Variables.
humanDate = Trim(FormatDateTime(Now, vbShortDate)) 
logDate = Trim(Replace(humanDate, "/", "-"))
humanTime = Trim(FormatDateTime(Now, vbLongTime))
logTime = Trim(Replace(Replace(humanTime, ":", "-"), " ", ""))
humanDateTime = Trim(humanDate & " " & humanTime)
logDateTime = Trim(logDate & "_" & logTime)
'Directory Related Variables.
fullScriptName = Trim(Replace(HRAV.commandLine, Chr(34), ""))
currentDirectory = Trim(objFSO.GetAbsolutePathName("."))
If InStr(fullScriptName, "-startup") > 0 Then
  startup = TRUE
  fullScriptName = "C:\Program Files\HR-AV\HR-AV.hta"
End If
If InStr(fullScriptName, "Program Files") > 0 Then
  currentDirectory = Replace(fullScriptName, appName & ".hta", "")
  currentDirectory = Mid(currentDirectory, 1, len(currentDirectory) - 1)
End If
scriptsDirectory = currentDirectory & "\Scripts\"
vbsScriptsDirectory = scriptsDirectory & "VBS\"
binariesDirectory = currentDirectory & "\Binaries\"
cacheDirectory = currentDirectory & "\Cache\"
tempDirectory = currentDirectory & "\Temp\"
pagesDirectory = currentDirectory & "\Pages\"
mediaDirectory = currentDirectory & "\Media\"
logsDirectory = currentDirectory & "\Logs\"
reportsDirectory = currentDirectory & "\Reports\"
resourcesDirectory = currentDirectory & "\Resources\"
exceptionDirectory = currentDirectory & "\Exceptions\"
logFilePath = Trim(logsDirectory & appName & "-Log_" & logDate)
realTimeCoreFile = vbsScriptsDirectory & "Real-Time-Core.vbs"
tempFile = tempDirectory & "temp.txt"
stempFile = tempDirectory & "systemp.txt"
RTPCacheFile1 = cacheDirectory & "RTP-cache1.dat"
RTPCacheFile2 = cacheDirectory & "RTP-cache2.dat"
excepptionFile = exceptionDirectory & "Exception_List.csv"
requiredDirs = Array(scriptsDirectory, binariesDirectory, tempDirectory, cacheDirectory, mediaDirectory, logsDirectory, reportsDirectory, exceptionDirectory)
requiredCacheFiles = Array(RTPCacheFile1, RTPCacheFile2)
arrFN = Split(fullScriptName, "\")
scriptName = Trim(arrFN(UBound(arrFN)))
'Misc Variables.
sesID = Int(Rnd * 10000000)
strNamespace = "root\cimv2"
strCurrentUserName = Trim(objSysInfo.UserName)
strHRAVUserName = "HRAV"
strHRAVGroupName = "Administrators" 
strComputer = "."
strComputerName = Trim(objWshNet.ComputerName)
dontContinue = FALSE
'--------------------------------------------------

'--------------------------------------------------
'A function fo verify that all required directories exist and try to create them when they don't.
'If "dieOnInstallationError" is set to TRUE this application will die when required directories do not exist.
Function verifyDirectories()
  verifyDirectories = TRUE
  For Each requiredDir In requiredDirs
    If dieOnInstallationError = FALSE Then 
      On Error Resume Next
    End If
    If Not objFSO.FolderExists(requiredDir) Then
      objFSO.CreateFolder(requiredDir)
      If Not objFSO.FolderExists(requiredDir) Then
        verifyDirectories = FALSE
      End If
    End If
  Next
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to verify that all required cache files exist and try to create them when they don't
Function verifyCache()
  For Each requiredCacheFile in requiredCacheFiles
    If Not objFSO.FileExists(requiredCacheFile) Then
      objFSO.CreateTextFile requiredCacheFile, TRUE, TRUE
      If Not objFSO.FileExists(requiredCacheFile) Then
        DieGracefully 1122, "Cannot create required cache files!", FALSE
      End If
    End If
  Next
  Set oRTPCacheFile1 = objFSO.GetFile(RTPCacheFile1)
  Set oRTPCacheFile2 = objFSO.GetFile(RTPCacheFile2)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for sanitizing user input strings for strict use cases.
'Variables are redefined on every call incase they are compromised.
Function Sanitize(strToClean1)
  charArr = Array("/", "\", ":", "*", """", "<", ">", ",", "&", "#", "~", "%", "{", "}", "+")
  Sanitize = FALSE
  For Each tmpChar In charArr
    strToClean1 = Replace(strToClean1, tmpChar, "")
  Next
  Sanitize = strToClean1
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for sanitizing user input strings for use in directory paths.
'Variables are redefined on every call incase they are compromised.
Function SanitizeFolder(strToClean2)
  charArr2 = Array("*", """", "<", ">", "&", "#", "~", "{", "}", "+")
  SanitizeFolder = FALSE
  For Each tmpChar2 In charArr2
    strToClean2 = Replace(strToClean2, tmpChar2, "")
  Next
  SanitizeFolder = strToClean2
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to generate a random string of a specified length.
Function RandomString(rStrLen)
  Randomize
  For i1 = 1 to strLen
    rStr = rStr & Mid(Letters, Int(rStrLen * Rnd + 1))
  Next
  RandomString = rStr
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
'Appends to an existing "logFilePath" is one exists.
'Creates a new file if none exists.
Function createLog(strEventInfo)
  strEventInfo = Sanitize(Trim(strEventInfo)) & vbNewLine
  createLog = FALSE
  If Not strEventInfo = "" And strEventInfo <> FALSE And strEventInfo <> NULL Then
    Set objLogFile = oFSO.CreateTextFile(logFilePath, ForAppending, TRUE)
    objLogFile.WriteLine(Trim(SanitizeFolder(strEventInfo)))
    objLogFile.Close
    createLog = TRUE 
  End If
  If objFSO.FileExists(logFilePath) = FALSE Then
    createLog = FALSE
  End If
  strEventInfo = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to kill the script when a critical error occurs and display a useful message to the user.
'Also logs the output.
Function DieGracefully(errorNumber, errorMessage, quietly)
  cantError = FALSE
  If Not IsNumeric(errorNumber) Or TypeName(errorMessage) <> "String" Then
    cantError = TRUE
    MsgBox appName & "-" & sesID & " ERROR-" & errorNumber & " on " & humanDateTime & ", There was a critical error, but due to the severity of the error more information cannot be displayed. The application will now terminate."
  End If
  errorMessage = appName & "-" & sesID & " ERROR-" & errorNumber & " on " & humanDateTime & ", " & SanitizeFolder(errorMessage) & "!"
  createLog(errorMessage)
  If cantError <> TRUE Then
    If quietly <> TRUE Then
      MsgBox errorMessage, 16, "ERROR!!! - " & appName
    End If
    If IsNumeric(errorMessage) = FALSE Then
      errorNumber = 0
    End If
  End If
  Window.Close
  errorMessage = NULL
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to display a consistent message box to the user.
'Also logs the output.
Function PrintGracefully(windowNote, message, typeMsg)
  If typeMsg = "vbOkCancel" Then
    typeMsg = 1
  Else
    typeMsg = 0
  End If
  windowNote = SanitizeFolder(windowNote)
  message = SanitizeFolder(message)
  PrintGracefully = MsgBox(message, typeMsg, appName & " - " & windowNote)
  createLog(message)
  If PrintGracefully = 2 Or PrintGracefully = 3 Then
    dontContinue = TRUE
    DieGracefully 500, "Operation cancelled by user!", FALSE 
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to open a dialog box so the user can select files or folders.
Function BrowseForFile()
  Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "C:\")
  BrowseForFile = file.self.Path
  createLog("File browser selected file: " & BrowseForFile)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to play sounds and media with Windows Media Player objects.
Function playMedia(pathToMedia)
  Set mediaPlayer = CreateObject("WMPlayer.OCX")
  mediaPlayer.URL = SanitizeFolder(mediaDirectory & pathToMedia)
  mediaPlayer.controls.play 
  mediaPlayer.close
  createLog("Media player played file: " & pathToMedia)
End Function
'--------------------------------------------------

'--------------------------------------------------
'Bootstrap some other program or code in the Binaries folder.
'Example for bootstrapping a PHP script.
'  Bootstrap("PHP\php.exe", scriptsDirectory & "PHP\test.php", TRUE)
'The above function call uses the Bootstrap() function to call 
'Binaries\PHP\php.exe with an argument that evaluates to Scripts\PHP\test.php.
'The result will be that the PHP binary is used to execute a PHP script.
'If Async is set to TRUE, HTA-UI will wait for the command to finish before continuing.
Function Bootstrap(BinaryToRun, Command, Async)
  If Async = TRUE Then 
    async = TRUE
  Else 
    async = FALSE
  End If
  run = Trim("C:\Windows\System32\cmd.exe /c " & SanitizeFolder(binariesDirectory & BinaryToRun) & " " & Command & " > " & SanitizeFolder(tempFile))
  objShell.Run run, 0, async
  run = NULL
  If Not objFSO.FileExists(tempFile) Then
    objFSO.CreateTextFile tempFile, TRUE, TRUE 
  End If
  If Not objFSO.FileExists(stempFile) Then
    DieGracefully 1000, "Cannot create a temporary Bootstrap file at: '" & stempFile & "'!", FALSE 
  End If
  Set tempData = objFSO.OpenTextFile(tempFile, 1)
  If Not tempData.AtEndOfStream Then 
    Bootstrap = tempData.ReadAll()
  Else
    Bootstrap = FALSE
  End If
  tempData.Close
  createLog("Bootstrapper ran binary:" & BinaryToRun)
  'objFSO.DeleteFile(tempFile)
End Function
'--------------------------------------------------

'--------------------------------------------------
'SystemBootstrap some other program or code without preparing a folder for the binary.
'Example for system bootstrapping a the shutdown command with restart argument.
'  SystemBootstrap("shutdown.exe", "/r", TRUE)
'The above function call uses the SystemBootstrap() function to call 
'shutdown.exe with an argument that evaluates to /r.
'The result will be that the shutdown.exe binary is used with the /r argument to restart the computer.
'If sAsync is set to TRUE, HTA-UI will wait for the command to finish before continuing.
Function SystemBootstrap(sBinaryToRun, sCommand, sAsync)
  If sAsync = TRUE Then
    sasync1 = TRUE
  Else 
    sasync1 = FALSE
  End If
  srun = Trim("C:\Windows\System32\cmd.exe /c " & sCommand & " > " & SanitizeFolder(stempFile))
  objShell.Run srun, 0, sasync1
  srun = NULL
  If Not objFSO.FileExists(stempFile) Then
    objFSO.CreateTextFile stempFile, TRUE, TRUE 
  End If
  If Not objFSO.FileExists(stempFile) Then
    DieGracefully 1001, "Cannot create a temporary SystemBootstrap file at: '" & stempFile & "'!", FALSE 
  End If
  Set stempData = objFSO.OpenTextFile(stempFile, 1)
  If Not stempData.AtEndOfStream Then 
    SystemBootstrap = stempData.ReadAll()
  Else
    SystemBootstrap = FALSE
  End If
  stempData.Close
  createLog("System Bootstrapper ran binary:" & sBinaryToRun)
  'objFSO.DeleteFile(stempFile)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to sleep execution in HTA's since WScript.Sleep isn't available.
Sub Sleep(Timesec)
  objShell.Run "Timeout /T " & Timesec & " /nobreak", 0, TRUE
End Sub
'--------------------------------------------------

'--------------------------------------------------
'Load the main application window.
'Put a Bootstrap function in here to have it run as soon as the window has been displayed.
'Useful for longer running scripts and programs.
Sub Window_OnLoad 
  Set dMenus = createObject("Scripting.Dictionary") 
  For Each entry In Split(sMenuItems, ",") 
    menu.innerHTML = Trim(menu.innerHTML & "&nbsp;<span id=" & entry _ 
      & " style='padding-bottom:2px' onselectstart=cancelEvent>&nbsp;" _ 
      & entry & "&nbsp;</span>&nbsp;&nbsp;") 
    dMenus.Add entry, Split(eval("s" & entry), ",") 
  Next 
  sMenuOpen = "" 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Resize the application window.
Set objWMIService = GetObject("winmgmts:\\" & Sanitize(strComputer) & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
For Each objItem in colItems
  'Max screen width in pixels.
  intHorizontal = objItem.ScreenWidth
  'Max screen height in pixels.
  intVertical = objItem.ScreenHeight
Next
  window.resizeTo windowHeight,windowWidth
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes on mouse hover.
Sub menu_onmouseover 
  clearmenu 
  With window.event.srcElement 
    If .parentElement.ID = "menu" Then 
      .style.border = "thin outset" 
      .style.cursor = "arrow" 
    End if 
  End With 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse leaves hover.
Sub menu_onmouseout 
  With window.event.srcElement 
    .style.border = "none" 
    .style.cursor = "default" 
  End With 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse hovers over a dropdown menu item.
Sub dropmenu_onmouseover 
  With window.event 
    .srcElement.style.cursor = "arrow" 
    .cancelbubble = TRUE 
    .returnvalue = FALSE 
  End With 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when a user hovers over a dropdown menu selection.
Sub SubMenuOver 
  With window.event.srcElement 
    If .ID = "dropmenu" Then Exit Sub 
      .style.backgroundcolor = "darkblue" 
      .style.color = "white" 
      .style.cursor = "arrow" 
  End With 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when mouse leaves hover over a dropdown menu selection.
Sub SubMenuOut 
  With window.event.srcElement 
    .style.backgroundcolor = "lightgrey" 
    .style.color = "black" 
    .style.cursor = "default" 
  End With 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle UI changes when a user clicks on a menu item.
Sub menu_onclick  
  If sMenuOpen <> "" Then Exit Sub 
  With window.event.srcElement 
    If .ID <> "menu" Then 
      .style.border = "thin inset" 
      nLeft = Trim(.offsetLeft)
      ntop  = Trim(.offsetTop + Replace(menu.style.Height, "px", "") - 5)
      sMenuOpen = Trim(.innertext) 
      With dropmenu 
        With .style 
          .border = "thin outset" 
          .backgroundcolor = "lightgrey" 
          .position = "absolute" 
          .Left = nLeft 
          .top = nTop 
          .width = "100px" 
          .zIndex = "101"
        End With 
        For Each sItem In dMenus.Item(sMenuOpen) 
          set oEL = document.createElement("SPAN") 
          .appendChild(oEL) 
          With oEl 
            .ID = sItem 
            .style.height = "20px" 
            .style.width = dropmenu.style.width 
            .style.zIndex = "102"
            .innerHTML = Replace(sHTML, "#sItem#", Trim(sItem)) 
            Set .onmouseover = getRef("SubMenuOver") 
            Set .onmouseout = getRef("SubMenuOut") 
            Set .onclick = getRef("SubMenuClick") 
            Set .onselectstart = getRef("cancelEvent") 
          End With
          Set oEL = document.createElement("BR") 
          .appendChild(oEL) 
        Next 
      End With
    End If 
  End With
end sub 
'--------------------------------------------------

'--------------------------------------------------
'Handle when an event is cancelled.
Sub cancelEvent 
  window.event.returnValue = FALSE 
End Sub
'--------------------------------------------------

'--------------------------------------------------
'Handle when a user deselects a menu.
Sub clearmenu 
  dropmenu.innerHTML = "" 
  dropmenu.style.border = "none" 
  dropmenu.style.backgroundcolor = "transparent" 
  If sMenuOpen <> "" Then 
    document.getElementByID(sMenuOpen).style.border = "none" 
    sMenuOpen = "" 
  End If 
End Sub 
'--------------------------------------------------

'--------------------------------------------------
'Display a MsgBox window confirming to the user that they have saved their settings.
Function saveSettings()
  PrintGracefully "Settings", "All settings saved and applied!", "vbOkOnly"
End Function
'--------------------------------------------------

'--------------------------------------------------
'Handle when a user clicks on a submenu.
Sub SubMenuClick 
  Set oRE = New RegExp
  sItem = Trim(window.event.srcElement.innerText) 
  clearmenu   
  hrefLocation = "Pages/"
  oRE.Pattern = "Pages"
  oRE.Global = TRUE
  Set oMatches = oRE.Execute(document.location.href)
  For Each oMatch In oMatches
    hrefLocation = ""
  Next
  Select Case LCase(sItem) 
    Case "exit" 
      window.close  
    Case "view settings"
      document.location = hrefLocation & "settings.hta"
    Case "about" 
      PrintGracefully "About", version & ". " & vbCRLF & vbCRLF & "Developed by " & developerName & "." & vbCRLF & vbCRLF & developerURL, "vbOkOnly"
    Case Else 
      PrintGracefully "Help", "You can get support for '" & appName & "' by visiting: " & vbCRLF & vbCRLF & helpLocSetting & ".", "vbOkOnly"
  End Select 
End Sub 
'--------------------------------------------------
