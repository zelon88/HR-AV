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

Option Explicit
'Large portions of code in this file were borrowed from the Microsoft TechNet website on 8/14/2019 
'in accordance with the Microsoft Limited Public License...
'https://gallery.technet.microsoft.com/scriptcenter/796bd584-0fdb-43bc-a5d2-aa5fc99a5e5d

'--------------------------------------------------
'Define global variables for the session.
Dim objFSO, strComputer, objWMIService, scriptsDirectory, binariesDirectory, _
 colItems, objItem, intHorizontal, intVertical, nLeft, nTop, sItem, helpLocSetting, _
 version, currentDirectory, appName, developerName, developerURL, windowHeight, windowWidth, _
 BinaryToRun, Command, tempDirectory, uiVersion, Async, error, requiredDir, requiredDirs, installationError, _
 dieOnInstallationError, cacheDirectory, pagesDirectory, realDirectory, vbsScriptsDirectory, dMenus, sMenuOpen, _
 hrefLocation, fullScriptName, arrFN, scriptName, oRE, oMatch, oMatches, shell, file, objWshNet, strNamespace, strHRAVUserName, _
 strHRAVGroupName, strCurrentUserName, oEL, oItem, objShell, objShellExec, run, tempFile, tempData, entry, objSysInfo, strComputerName, _
 sBinaryToRun, sCommand, sAsync, stempFile, stempDirectory, sasync1, srun, stempData, mediaPlayer, pathToMedia, mediaDirectory, message, errorNumber, _
 errorMessage, sCommLine, dProcess, cProcessList, quietly

'--------------------------------------------------
'Application Related Variables
version = "v0.6.9"
uiVersion = "v1.2"
helpLocSetting = "https://github.com/zelon88/HR-AV"
appName = "HR-AV"
developerName = "Justin Grimes"
developerURL = "https://github.com/zelon88"
dieOnInstallationError = TRUE
windowHeight = 660
windowWidth = 600
'UI Related Variables.
Const sMenuItems = "File,Settings,Help" 
Const sFile = "Exit" 
Const sSettings = "View Settings"
Const sHelp = "Help, About" 
Const sHTML = "&nbsp;&nbsp;&nbsp;#sItem#&nbsp;&nbsp;&nbsp;" 
'Frequently used objects.
Set objShell = CreateObject("WScript.Shell")
Set shell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("WinNTSystemInfo")
Set objWshNet = CreateObject("WScript.Network")
'Directory related variables.
currentDirectory = objFSO.GetAbsolutePathName(".")
scriptsDirectory = currentDirectory & "\Scripts\"
vbsScriptsDirectory = scriptsDirectory & "\VBS\"
binariesDirectory = currentDirectory & "\Binaries\"
cacheDirectory = currentDirectory & "\Cache\"
tempDirectory = currentDirectory & "\Temp\"
pagesDirectory = currentDirectory & "\Pages\"
mediaDirectory = currentDirectory & "\Media\"
requiredDirs = array(scriptsDirectory, binariesDirectory, tempDirectory, cacheDirectory, mediaDirectory)
fullScriptName = Trim(Replace(HRAV.commandLine, Chr(34), ""))
arrFN = Split(fullScriptName, "\")
scriptName = Trim(arrFN(UBound(arrFN)))
'Misc variables.
strNamespace = "root\cimv2"
strCurrentUserName = Trim(objSysInfo.UserName)
strHRAVUserName = "HRAV"
strHRAVGroupName = "Administrators" 
strComputer = "."
strComputerName = Trim(objWshNet.ComputerName)
installationError = FALSE
'--------------------------------------------------

'--------------------------------------------------
'Verify that all required directories exist and try to create them when they don't.
'If "dieOnInstallationError" is set to TRUE this application will die when required directories do not exist.
For Each requiredDir In requiredDirs
  If dieOnInstallationError = TRUE Then 
    On Error Resume Next
  End If
  If Not objFSO.FolderExists(requiredDir) Then
    objFSO.CreateFolder(requiredDir)
    If Not objFSO.FolderExists(requiredDir) Then
      installationError = TRUE
    End If
  End If
Next
'--------------------------------------------------

'--------------------------------------------------
'A function for sanitizing user input strings for strict use cases.
Function Sanitize(strToClean)
  Set oRE1 = New RegExp
  oRE1.IgnoreCase = True
  oRE1.Global = True
  oRE1.Pattern = "[(?*"",\\<>&#~%{}+_.@:\/!;]+"
  outputStr = oRE1.Replace(strtoclean, "-")
  oRE1.Pattern = "\-+"
  Sanitize = oRE1.Replace(outputStr, "-")
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for sanitizing user input strings for use in directory paths.
Function SanitizeFolder(strToClean)
  Set oRE1 = New RegExp
  oRE1.IgnoreCase = True
  oRE1.Global = True
  oRE1.Pattern = "[(?*"",<>&#~%{}+_.@:!;]+"
  outputStr = oRE1.Replace(strtoclean, "-")
  oRE1.Pattern = "-+"
  SanitizeFolder = oRE1.Replace(outputStr, "-")
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to kill the script when a critical error occurs and display a useful message to the user.
Function DieGracefully(errorNumber, errorMessage, quietly)
  errorMessage = Sanitize(errorMessage)
  If quietly <> TRUE Then
    MsgBox appName & " ERROR!!! " & errorNumber & " " & errorMessage, "ERROR!!! - " & appName, 16
  End If
  If IsNumeric(errorMessage) = FALSE Then
    errorNumber = 0
  End If
  For Each dProcess in cProcessList
    sCommLine = Trim(LCase(dProcess.CommandLine))
    If InStr(sCommLine, fullScriptName) = 0 Then
      dProcess.Terminate()
    End If
  Next
  Window.Close
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to display a consistent message box to the user.
Function PrintGracefully(message)
  message = Sanitize(message)
  MsgBox message, appName, 0
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to open a dialog box so the user can select files or folders.
Function BrowseForFile()
  Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "C:\")
  BrowseForFile = file.self.Path
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to play sounds and media with Windows Media Player objects.
Function playMedia(pathToMedia)
  Set mediaPlayer = CreateObject("WMPlayer.OCX")
  mediaPlayer.URL = SanitizeFolder(mediaDirectory & pathToMedia)
  mediaPlayer.controls.play 
  mediaPlayer.close
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
  tempFile = tempDirectory & "temp.txt"
  If Async = TRUE Then 
    async = TRUE
  Else 
    async = ""
  End If
  run = Trim("C:\Windows\System32\cmd.exe /c " & SanitizeFolder(binariesDirectory & BinaryToRun) & " " & Command & " > " & SanitizeFolder(tempFile))
  objShell.Run run, 0, async
  Set tempData = objFSO.OpenTextFile(tempFile, 1)
  Bootstrap = tempData.ReadAll()
  tempData.Close
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
  stempFile = stempDirectory & "systemp.txt"
  If sAsync = TRUE Then
    sasync1 = TRUE
  Else 
    sasync1 = ""
  End If
  srun = Trim("C:\Windows\System32\cmd.exe /c " & sCommand & " > " & SanitizeFolder(stempFile))
  objShell.Run srun, 0, sasync1
  Set stempData = objFSO.OpenTextFile(stempFile, 1)
  SystemBootstrap = stempData.ReadAll()
  stempData.Close
  'objFSO.DeleteFile(stempFile)
End Function
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
  PrintGracefully("All settings saved and applied!")
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
      MsgBox version & ". " & vbCRLF & vbCRLF & "Developed by " & developerName & "."_ 
        & vbCRLF & vbCRLF & developerURL, _ 
        vbOKOnly + vbInformation, "About "& appName 
    Case Else 
      msgbox "You can get support for '" & appName & "' by visiting: " _ 
      & vbCRLF & vbCRLF & helpLocSetting, vbOKOnly + vbInformation, appName & " Help"
  End Select 
End Sub 
'--------------------------------------------------
