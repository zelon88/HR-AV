'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 10/28/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'Portions of the UI-Core.vbs file are licensed under the Microsoft Limited Public License.
'Copies of all applicable software licenses can be found in the "Documentation" directory.

'--------------------------------------------------
'Set global variables for the session.
Option Explicit

Dim objShell, objFSO, sesID, humanDate, logDate, humanTime, logTime, humanDateTime, logDateTime, currentDirectory, configFile, scriptsDirectory, vbsScriptsDirectory, _
 binariesDirectory, cacheDirectory, tempDirectory, pagesDirectory, mediaDirectory, logsDirectory, reportsDirectory, resourcesDirectory, logFilePath, objVBSFile, Timesec, _
 charArr, charArr2, tmpChar, tmpChar2, strToClean1, strToClean2, strEventInfo, objLogFile, logFilePath, whoamiOutput, strHRAVUserName, strHRAVPassword, fullScriptName, arr, _
 obj, x, i, tempArray, rpCounter, pcs, oWMISrvc, errorNumber, errorMessage, quietly, cantError, windowNote, message, typeMsg, dontContinue, sBinaryToRun, sCommand, sAsync, srun, _
 stempfile, sasync1, stempData, searchScripts, scriptsToSearch, procSearch, procsToSearch

'Commonly Used Objects.
Set objShell = CreateObject("WScript.Shell")
Set oWMISrvc = GetObject("winmgmts:")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Environment Related Variables.
sesID = Int(Rnd * 10000000)
'Time Related Variables.
humanDate = Trim(FormatDateTime(Now, vbShortDate)) 
logDate = Trim(Replace(humanDate, "/", "-"))
humanTime = Trim(FormatDateTime(Now, vbLongTime))
logTime = Trim(Replace(Replace(humanTime, ":", "-"), " ", ""))
humanDateTime = Trim(humanDate & " " & humanTime)
logDateTime = Trim(logDate & "_" & logTime)
'Directory Related Variables.
currentDirectory = Replace(Trim(objFSO.GetAbsolutePathName(".")), "\Scripts\VBS\", "")
currentDirectory = Mid(currentDirectory, 1, len(currentDirectory) - 11)
configFile = currentDirectory & "\Config\Config.vbs"
scriptsDirectory = currentDirectory & "\Scripts\"
vbsScriptsDirectory = scriptsDirectory & "\VBS\"
binariesDirectory = currentDirectory & "\Binaries\"
cacheDirectory = currentDirectory & "\Cache\"
tempDirectory = currentDirectory & "\Temp\"
pagesDirectory = currentDirectory & "\Pages\"
mediaDirectory = currentDirectory & "\Media\"
logsDirectory = currentDirectory & "\Logs\"
reportsDirectory = currentDirectory & "\Reports\"
resourcesDirectory = currentDirectory & "\Resources\"
logFilePath = Trim(logsDirectory & "RTP-Log_" & logDate)
'--------------------------------------------------

'--------------------------------------------------
'A function to execute VBS scripts in the context and scope of the running script. Works just like a PHP include().
'https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Sub Include(pathToVBS) 
  Set objVBSFile = objFSO.OpenTextFile(pathToVBS, 1)
  ExecuteGlobal objVBSFile.ReadAll
  objVBSFile.Close
  objVBSFile = NULL
End Sub
'--------------------------------------------------

'--------------------------------------------------
'A function to sleep execution in HTA's since WScript.Sleep isn't available.
Sub Sleep(Timesec)
  objShell.Run "Timeout /T " & Timesec & " /nobreak", 0, TRUE
End Sub
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
'A function to create a log file.
'Appends to an existing "logFilePath" is one exists.
'Creates a new file if none exists.
Function createLog(strEventInfo)
  strEventInfo = Sanitize(Trim(strEventInfo))
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
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated as HRAV user.
'Returns FALSE if the application is not elevated as HRAV user.
Function isUserHRAV()
  On Error Resume Next
  whoamiOutput = Sanitize(SystemBootstrap("whoami", "", FALSE))
  objShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
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
'A function to returns the index of string "obj" in array "arr". obj can be anything. 
'Returns TRUE if "obj" is in "arr". Returns FALSE if nothing was found.
'https://gist.github.com/sholsinger/943116/caf67a2504d6e45e4acc49597fac5f1bb6033ba2
Function InArray(arr, obj)
  On Error Resume Next
  x = FALSE
  If isArray(arr) Then
    For i = 0 To UBound(arr)
      If arr(i) = obj Then
        x = TRUE
        Exit For
      End If
    Next
  End If  
  Call err_report()
  InArray = x
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to enumerate running processes into an array. 
'Each array element contains a CSV containing a PID, process name, and executable path.
Function enumerateRunningProcesses() 
  enumerateRunningProcesses = Array()
  tempArray = Array() 
  rpCounter = 0
  For each pcs in oWMISrvc.InstancesOf("Win32_Process")
  Redim Preserve tempArray(rpCounter)
    tempArray(rpCounter) = pcs.ProcessID & "," & pcs.Name & "," & pcs.ExecutablePath
    rpCounter = rpCounter + 1
  Next
  enumerateRunningProcesses = tempArray
  tempArray = NULL
  rpCounter = NULL
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
  msgbox stempfile
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
'A function to kill all scripts running with CScript or WScript. 
'Assumes TRUE until an instance of wscript or cscript are found.
'Leaves HTA's running.
Function killAllScripts()
  killAllScripts = TRUE
  createLog("Attempting to kill running scripts.")
  objShell.Run "taskkill /im wscript.exe", , FALSE
  objShell.Run "taskkill /im cscript.exe", , FALSE
  searchScripts = enumerateRunningProcesses()
  For Each scriptsToSearch In searchScripts
    If InStr(LCase(scriptsToSearch), "wscript") > 0 Or InStr(LCase(scriptsToSearch), "cscript") > 0 Then
      killAllScripts = FALSE
      createLog("Could not kill script: """ & scriptsToSearch & """!")
    End If
  Next
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to check that the main mshta.exe process is running. 
'Assumes FALSE until an instance of mshta or appName are found.
'Leaves HTA's running.
Function checkForMSTHA() 
  checkForMSTHA = FALSE
  createLog("Checking for MSHTA.exe.")
  procSearch = enumerateRunningProcesses() 
  For Each procsToSearch In procSearch
    If InStr(LCase(scriptsToSearch), "mshta") > 0 And InStr(LCase(scriptsToSearch), appName) > 0 Then
      checkForMSTHA = TRUE
    End If
  Next
  procSearch = NULL
  procsToSearch = NULL
End Function 
'--------------------------------------------------