'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 11/7/2019
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
 stempfile, sasync1, stempData, searchScripts, scriptsToSearch, procSearch, procsToSearch, strComputer, objRAMService,  result, resultSet, availableRAMBytes, availableRAMKB, availableRAMMB, _
 availableRAMGB, commitLimitRAMBytes, commitLimitRAMKB, commitLimitRAMMB, commitLimitRAMGB, committedRAMBytes, committedRAMKB, committedRAMMB, committedRAMGB, objDrives, objDrive, edCounter

'Commonly Used Objects.
Set objShell = CreateObject("WScript.Shell")
Set oWMISrvc = GetObject("winmgmts:")
Set objRAMService = GetObject("winmgmts:\\.\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Environment Related Variables.
Const sesID = Int(Rnd * 10000000)
Const KB = 1024
Const MB = KB * 1024
Const GB = MB * 1024
'Time Related Variables.
Const humanDate = Trim(FormatDateTime(Now, vbShortDate)) 
Const logDate = Trim(Replace(humanDate, "/", "-"))
Const humanTime = Trim(FormatDateTime(Now, vbLongTime))
Const logTime = Trim(Replace(Replace(humanTime, ":", "-"), " ", ""))
Const humanDateTime = Trim(humanDate & " " & humanTime)
Const logDateTime = Trim(logDate & "_" & logTime)
'Directory Related Variables.
Const strComputer = "."
currentDirectory = Replace(Trim(objFSO.GetAbsolutePathName(strComputer)), "\Scripts\VBS\", "")
currentDirectory = Trim(Mid(currentDirectory, 1, len(currentDirectory) - 11))
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
  'Set an object to the target VBS code file.
  Set objVBSFile = objFSO.OpenTextFile(pathToVBS, 1)
  'Read the contents of the target VBS code file and execute it in a global scope.
  ExecuteGlobal objVBSFile.ReadAll
  'Close the object handle to the target VBS code file.
  objVBSFile.Close
  'Clean up unneeded memory. Destroy the handle to the target VBS code file so it cannot be reused.
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
  'Redefine charArr every time this function is called in case it becomes compromised.
  charArr = Array("/", "\", ":", "*", """", "<", ">", ",", "&", "#", "~", "%", "{", "}", "+")
  Sanitize = FALSE
  'Iterate through the charArr and remove each instance of matching character from the input string.
  For Each tmpChar In charArr
    strToClean1 = Replace(strToClean1, tmpChar, "")
  Next
  Sanitize = strToClean1
  'Clean up unneeded memory.
  tmpChar = NULL
  charArr = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for sanitizing user input strings for use in directory paths.
'Variables are redefined on every call incase they are compromised.
Function SanitizeFolder(strToClean2)
  'Redefine charArr every time this function is called in case it becomes compromised.
  charArr2 = Array("*", """", "<", ">", "&", "#", "~", "{", "}", "+")
  SanitizeFolder = FALSE
  'Iterate through the charArr and remove each instance of matching character from the input string.
  For Each tmpChar2 In charArr2
    strToClean2 = Replace(strToClean2, tmpChar2, "")
  Next
  SanitizeFolder = strToClean2
  'Clean up unneeded memory.
  tmpChar2 = NULL
  charArr2 = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
'Appends to an existing "logFilePath" is one exists.
'Creates a new file if none exists.
Function createLog(strEventInfo)
  'Sanitize the input variable.
  strEventInfo = Sanitize(Trim(strEventInfo))
  createLog = FALSE
  'Perform some sanity checks on the input variable.
  If Not strEventInfo = "" And strEventInfo <> FALSE And strEventInfo <> NULL Then
    'Create a handle object to the logFilePath for appending data to the end of it.
    Set objLogFile = oFSO.CreateTextFile(logFilePath, ForAppending, TRUE)
    'Sanitize the input string again & write it to the end of the logFilePath.
    objLogFile.WriteLine(Trim(SanitizeFolder(strEventInfo)))
    'Close the handle to the logFilePath.
    objLogFile.Close
    'This function will return TRUE if a logfile was successfully written.
    createLog = TRUE 
  End If
  'One last check to be sure that a log was actually written when it was supposed to be above.
  If objFSO.FileExists(logFilePath) = FALSE Then
    createLog = FALSE
  End If
  'Clean up unneeded memory.
  strEventInfo = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated as HRAV user.
'Returns FALSE if the application is not elevated as HRAV user.
Function isUserHRAV()
  On Error Resume Next
  'Determine who is executing the current script.
  whoamiOutput = Sanitize(SystemBootstrap("whoami", "", FALSE))
  'See if the current user has admin rights by trying to access an arbitrary registry key which requires admin rights.
  objShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  'If we are able to read the registry key above then either the user is an admin or the script is elevated.
  If Err.number = 0 And Trim(Replace(Replace(whoamiOutput, Chr(10), ""), Chr(13), "")) = strHRAVUserName Then 
    isUserHRAV = TRUE
  'If the attempt to read the registry key above fails we can assume that the current user does not have admin rights.
  'Could also mean that the script is not elevated.
  Else
    isUserHRAV = FALSE
  End If
  Err.Clear
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to restart the script with admin priviledges if required.
Function restartAsHRAV(strHRAVPassword)
  'Use PAExec (open-source version of PSExec) to re-run the current script as the HRAV user using HRAV login credentials.
  'These credentials will be visible via the task manager if the "Command Line" column is visible. 
  'That is why it is critical that HR-AV resets its own password each time it starts.
  'It is the only way to keep the password secret.
  Bootstrap "PAExec\paexec.exe", "-u:" & Sanitize(strHRAVUserName) & " -p:" & Sanitize(strHRAVPassword) & " " & SanitizeFolder(fullScriptName), FALSE
  'Kill the current script.
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
  'Clean up unneeded memory.
  i = NULL
  x = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to enumerate running processes into an array. 
'Each array element contains a CSV containing a PID, process name, and executable path.
Function enumerateRunningProcesses() 
  enumerateRunningProcesses = Array()
  tempArray = Array() 
  rpCounter = 0
  'Iterate through each process & build an array of CSV's containing information about each process.
  For each pcs in oWMISrvc.InstancesOf("Win32_Process")
    Redim Preserve tempArray(rpCounter)
    tempArray(rpCounter) = pcs.ProcessID & "," & pcs.Name & "," & pcs.ExecutablePath
    rpCounter = rpCounter + 1
  Next
  enumerateRunningProcesses = tempArray
  'Clean up unneeded memory.
  tempArray = NULL
  rpCounter = NULL
  pcs = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to kill the script when a critical error occurs and display a useful message to the user.
'Also logs the output.
Function DieGracefully(errorNumber, errorMessage, quietly)
  cantError = FALSE 
  'Perform some sanity checks on the input variables to be sure they are of the proper types.
  'errorNumber must be an integer.
  'errorMessage must be a string.
  'quietly must be a boolean value.
  If Not IsNumeric(errorNumber) Or TypeName(errorMessage) <> "String" Then
    cantError = TRUE
    MsgBox appName & "-" & sesID & " ERROR-" & errorNumber & " on " & humanDateTime & ", There was a critical error, but due to the severity of the error more information cannot be displayed. The application will now terminate."
  End If
  'If the sanity checks have passed we issue the error.
  errorMessage = appName & "-" & sesID & " ERROR-" & errorNumber & " on " & humanDateTime & ", " & SanitizeFolder(errorMessage) & "!"
  'Log the error message.
  createLog(errorMessage)
  If cantError <> TRUE Then
    'Unless specified, output the error to the user.
    If quietly <> TRUE Then
      MsgBox errorMessage, 16, "ERROR!!! - " & appName
    End If
    'If no specific error number is provided we assume error "0".
    If IsNumeric(errorMessage) = FALSE Then
      errorNumber = 0
    End If
  End If
  'Close the window.
  Window.Close
  'Clean up unneeded memory.
  errorMessage = NULL
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to display a consistent message box to the user.
'Also logs the output.
'"winowNote" is a string to be used for window Title and Task Bar text.
'"message" is the content of the MsgBox to be displayed to the user.
'If "typeMsg" is vbOkCancel script execution will be conditioned on the user selecting an "OK" prompt.
Function PrintGracefully(windowNote, message, typeMsg)
  'If typeMsg us vbOkCancel then user will be prompted for input. 
  'Execution will only continue if the user selects "OK".
  'Execution will halt if the user selects "Cancel".
  If typeMsg = "vbOkCancel" Then 
    typeMsg = 1
  Else
    typeMsg = 0
  End If
  'Sanitize input variables.
  windowNote = SanitizeFolder(windowNote)
  message = SanitizeFolder(message)
  'Display the prepared message to the user.
  PrintGracefully = MsgBox(message, typeMsg, appName & " - " & windowNote)
  'Create a log entry of the message.
  createLog(message)
  'Gather the input from the message box & kill the script if cancelled.
  If PrintGracefully = 2 Or PrintGracefully = 3 Then
    dontContinue = TRUE
    'Create a log & halt execution if the user clicks "Cancel".
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
  objFSO.DeleteFile(stempFile)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to kill all scripts running with CScript or WScript. 
'Assumes TRUE until an instance of wscript or cscript are found.
'Leaves HTA's running.
Function killAllScripts()
  killAllScripts = TRUE
  createLog("Attempting to kill running scripts.")
  'Before we scan for running processes try to kill anything we have running.
  objShell.Run "taskkill /im wscript.exe", , FALSE
  objShell.Run "taskkill /im cscript.exe", , FALSE
  'Gather running process list in array form.
  searchScripts = enumerateRunningProcesses()
  'Iterate through running processes and look for anything running with cscript.exe or wscript.exe.
  For Each scriptsToSearch In searchScripts
    If InStr(LCase(scriptsToSearch), "wscript") > 0 Or InStr(LCase(scriptsToSearch), "cscript") > 0 Then
      'If a program is still running with wscript.exe or cscript.exe this function returns FALSE & logs an warning.
      killAllScripts = FALSE
      createLog("Could not kill script: """ & scriptsToSearch & """!")
    End If
  Next
  'Clean up unneeded memory.
  scriptsToSearch = NULL
  searchScripts = NULL
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to check that the main mshta.exe process is running. 
'Assumes FALSE until an instance of mshta or appName are found.
'Leaves HTA's running.
Function checkForMSTHA() 
  checkForMSTHA = FALSE
  createLog("Checking for MSHTA.exe.")
  'Gather running process list in array form.
  procSearch = enumerateRunningProcesses() 
  'Iterate through running processes and look for anything running with mshta.exe.
  For Each procsToSearch In procSearch
    If InStr(LCase(scriptsToSearch), "mshta") > 0 And InStr(LCase(scriptsToSearch), appName) > 0 Then
      'If a program is running wih mshta.exe this function will return TRUE.
      checkForMSTHA = TRUE
    End If
  Next
  'Clean up unneeded memory.
  procSearch = NULL
  procsToSearch = NULL
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to get real-time RAM information. Used for thread management.
Function checkRAM()
  'Redefine query each time this function is called.
  Set resultSet = objRAMService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
  CreateLog("Checking system memory utilization.")
  For Each result in resultSet
    'Set variables for current Available RAM. In bytes, KB, MB, & GB.
    availableRAMBytes = Round(objItem.AvailableBytes,3)
    availableRAMKB = Round(availableRAMBytes / KB,3)
    availableRAMMB = Round(availableRAMBytes / MB,3)
    availableRAMGB = Round(availableRAMBytes / GB,3)
    'Set variables for current CommitLimit. In bytes, KB, MB, & GB.
    commitLimitRAMBytes = Round(objItem.CommitLimit,3)
    commitLimitRAMKB = Round(commitLimitRAMBytes / KB,3)
    commitLimitRAMMB = Round(commitLimitRAMBytes / MB,3)
    commitLimitRAMGB = Round(commitLimitRAMBytes / GB,3)
    'Set variables for current CommitLimit. In bytes, KB, MB, & GB.
    committedRAMBytes = Round(objItem.CommittedBytes,3)
    committedRAMKB = Round(committedRAMBytes / KB,3)
    committedRAMMB = Round(committedRAMBytes / MB,3)
    committedRAMGB = Round(committedRAMBytes / GB,3)
  Next
  'Clean up unneeded memory.
  resultSet = NULL
  result = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to enumerate the disk drives on the local system into an array.
'The objDrives object below must be re-declared each time this function is called because drive volumes can change.
Function enumerateDrives() 
  'Redefine variables each time this function is called.
  Set objDrives = objFSO.Drive
  enumerateDrives = Array()
  tempArray = Array()
  edCounter = 0
  'Iterate through each drive volume on the system and add it to the tempArray().
  For Each objDrive in objDrives
    tempArray(edCounter) = objDrive.DriveLetter
    edCounter = edCounter + 1
  Next
  'Copy the tempArray() to the enumerateDrives() array.
  enumerateDrives = tempArray
  'Clean up unneeded memory.
  objDrive = NULL
  objDrives = NULL
  tempArray = NULL
  edCounter = NULL
End Function
'--------------------------------------------------