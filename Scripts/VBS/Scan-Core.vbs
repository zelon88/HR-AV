'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 11/12/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'Portions of the UI-Core.vbs file are licensed under the Microsoft Limited Public License.
'Copies of all applicable software licenses can be found in the "Documentation" directory.

'--------------------------------------------------
'Set global variables for the session.
Option Explicit

Dim objShell, objFSO, sesID, humanDate, logDate, humanTime, logTime, humanDateTime, logDateTime, currentDirectory, _
 binariesDirectory, cacheDirectory, tempDirectory, pagesDirectory, mediaDirectory, logsDirectory, reportsDirectory, _ 
 resourcesDirectory, logFilePath, objVBSFile, Timesec, scriptsDirectory, vbsScriptsDirectory, arr, fullScriptName, _
 charArr, charArr2, tmpChar, tmpChar2, strToClean1, strToClean2, strEventInfo, objLogFile, logFilePath, whoamiOutput, _
 obj, x, i, tempArray, rpCounter, pcs, oWMISrvc, errorNumber, errorMessage, quietly, cantError, windowNote, message, _ 
 typeMsg, dontContinue, sBinaryToRun, sCommand, sAsync, srun, strHRAVUserName, strHRAVPassword, result, resultSet, _
 stempfile, sasync1, stempData, searchScripts, scriptsToSearch, procSearch, procsToSearch, strComputer, objRAMService, _
 availableRAMGB, commitLimitRAMBytes, commitLimitRAMKB, commitLimitRAMMB, commitLimitRAMGB, committedRAMBytes, _
 committedRAMKB, committedRAMMB, committedRAMGB, objDrives, objDrive, edCounter, availableRAMBytes, availableRAMKB, _
 eDelimiter, eString, eLimit, fgcPath, objFGCFile, exCounter, nexCounter, newInfection, infectionArray, exception, _
 exceptionFile, exceptionCSVData, type, workeType, targetType, memoryLimit, availableRAMMB, exceptionArray, _
 excepptionArray, priority, chunkCoef, priorityCoef, workerRAMLimit, availableRAM, workerChunkSize, workerLimit, _ 
 enumFolder, enumSubFolder, mValEl, mArray, mValue, tPath, checkExceptions, nexCounter, newInfection, exception,_
 exCounter, exceptionData, configFile, targets, wTarget, checkWorkerTimer

'Commonly Used Objects.
Set objShell = CreateObject("WScript.Shell")
Set oWMISrvc = GetObject("winmgmts:")
Set objRAMService = GetObject("winmgmts:\\.\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Environment Related Constants.
Const sesID = Int(Rnd * 10000000)
Const KB = 1024
Const MB = KB * 1024
Const GB = MB * 1024
Const priorityCoef = 1
Const chunkCoef = 3
Const checkWorkerTimer = 3
Const strComputer = "."
'Environment Related Variables.
workerLimit = 0
workerCount = 0
availableRAM = 0
workerRAMLimit = 0
workerChunkSize = 0
nexCounter = 0
exCounter = 0
exceptionData = ""
newInfection = ""
exception = ""
exceptionData = ""
exceptionArray = Array()
targetArray = Array()
checkExceptions = Array()
'Time Related Constants.
Const humanDate = Trim(FormatDateTime(Now, vbShortDate)) 
Const logDate = Trim(Replace(humanDate, "/", "-"))
Const humanTime = Trim(FormatDateTime(Now, vbLongTime))
Const logTime = Trim(Replace(Replace(humanTime, ":", "-"), " ", ""))
Const humanDateTime = Trim(humanDate & " " & humanTime)
Const logDateTime = Trim(logDate & "_" & logTime)
'Directory Related Variables.
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
exceptionDirectory = currentDirectory & "\Exceptions\"
excepptionFile = exceptionDirectory & "Exception_List.csv"
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
'http://www.vbforums.com/showthread.php?280636-VBScript-Array-Push
  Function push(ByRef mArray, ByVal mValue)
    If IsArray(mArray) Then
      If IsArray(mValue) Then
        For Each mValEl In mValue
          Redim Preserve mArray(UBound(mArray) + 1)
          mArray(UBound(mArray)) = mValEl
        Next
      Else
        Redim Preserve mArray(UBound(mArray) + 1)
        mArray(UBound(mArray)) = mValue
      End If
    Else
      If IsArray(mValue) Then
        mArray = mValue
      Else
        mArray = Array(mValue)
      End If
    End If
    push = UBound(mArray)
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
    MsgBox appName & "-" & sesID & " ERROR-" & errorNumber & " on " & humanDateTime & _ 
     ", There was a critical error, but due to the severity of the error more information cannot be displayed. " & _
     "The application will now terminate."
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
  createLog("Checking system memory utilization.")
  For Each result in resultSet
    'Set variables for current Available RAM. In bytes, KB, MB, & GB.
    availableRAMBytes = Round(objItem.AvailableBytes,3)
    checkRAM = availableRAMBytes
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
  createLog("Enumerating mounted disk drive volumes.")
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

'--------------------------------------------------
'A function to turn a CSV string variable into an array.
'Also works with other delimiters other than comma.
'https://phpvbs.verygoodtown.com/en/vbscript-explode-function/
Function explode(eDelimiter, eString, eLimit) 
  explode = FALSE
  If len(eDelimiter) = 0 Then Exit Function
  If len(eLimit) = 0 Then elimit = 0
  If eLimit > 0 Then
    explode = Split(eString, eDelimiter, eLimit)
  Else
    explode = Split(eString, eDelimiter)
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to read files into memory as a string like PHP's file_get_contents.
'Inspired by https://blog.ctglobalservices.com/scripting-development/jgs/include-other-files-in-vbscript/
Function fileGetContents(fgcPath) 
  createLog("Reading contents of file '" & SanitizeFolder(fgcPath) & "' into memory." )
  'Set a handle to the file to be opened.
  Set objFGCFile = objFSO.OpenTextFile(fgcPath, 1)
  'Read the contents of the file into a string.
  fileGetContents = objFGCFile.ReadAll
  'Close the handle to the file we opened earlier in the function.
  objFGCFile.Close
  'Clean up unneeded memory.
  objFGCFile = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to purge the infectionArray of exceptions.
'For performance, we check exceptions after all infections have been detected.
'We iterate throgugh all infections & check them against the exception list.
Function checkExceptions(infectionArray) 
  checkExceptions = Array()
  exCounter = 0
  nexCounter = 0
  newInfection = ""
  exception = ""
  exceptionArray = ""
  createLog("Checking exception list for matching infections.")
  'Detect if no exceptionFile exists & create one if needed.
  If Not objFSO.FileExists(exceptionFile) Then
    objFSO.CreateTextFile(exceptionFile, TRUE)
  End If
  'Load the exceptions.csv file and load it into an array.
  exceptionCSVData = fileGetContents(exceptionFile)
  exceptionArray - explode(",", exceptionCSVData, 0)
  'Iterate through the exception list & check if any of the detected infectinons are exempt.
  For Each exception In exceptionArray
    If InArray(infectionArray, exception) Then
      ReDim Preserve infectionArray(exCounter)
      infectionArray(exCounter) = ""
      createLog("Exception '" & SanitizeFolder(exception) & "' found.")
    End If
    exCounter = exCounter + 1
  Next
  'Rebuild the input array without the deleted elements found above.
  For Each newInfection In infectionArray
    If newInfection <> "" Then
      ReDim Preserve exceptionArray(nexCounter)
      checkExceptions(nexCounter) = newInfection
    End If
    nexCounter = nexCounter + 1
  Next
  'clean up unneeded memory.
  exCounter = NULL
  nexCounter = NULL
  newInfection = NULL
  exception = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to add a target file or registry key to the exception list.
Function addException(exception) 
  'Re-define variables incase this function has been called before.
  exceptionData = ""
  exceptionArray = ""
  createLog("Adding an exception for '" & SanitizeFolder(exception) & "'.")
  'See if a "exceptionFile" exists.
  If objFSO.FileExists(exceptionFile) Then
    'Read the exception file into memory.
    exceptionData = fileGetContents(exceptionFile)
    'Turn the comma separated list loaded into memory into an array.
    exceptionArray = explode(",", exceptionData, 0)
    'Delete the existing exception file. 
    'This triggers the additional code condition below that executes when an exception file is missing.
    objFSO.DeleteFile(exceptionFile)
  End If
  'If no "exceptionFile" exists. 
  'If no errors are encountered, this code should fire every time the function is called.
  If Not objFSO.FileExists(exceptionFile) Then
    'Create an empty exception file.
    Set objEFile = objFSO.CreateTextFile(exceptionFile, TRUE)
    'Re-define the length of the Exception array.
    ReDim Preserve exceptionArray(UBound(exceptionArray) + 1)
    'Add the new exception to the exception array.
    exceptionArray(UBound(exceptionArray)) = exception
    'Write the newly created exceptionArray to the newly created exception file.
    objEFile.WriteLine(Join(exceptionArray, ","))
    objEFile.Close
  End If
  'Clean up unneeded memory.
  exceptionData = NULL
  objEFile = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'https://stackoverflow.com/questions/1433785/vbscript-to-iterate-through-set-level-of-subfolders
Function enumerateSubdirs(enumFolder) 
  enumerateSubdirs = Array()
  'Iterate through each subfolder of the "enumFolder".
  For Each enumSubFolder in enumFolder.enumSubFolder
    enumSubFolderPath = enumSubFolder.Path 
    'Add the current path to the "targetArray".
    esdTemp = push(enumerateSubdirs, enumSubFolderPath)
    'We must use a temp variable here to avoid an out of memory error.
    enumerateSubdirs = esdTemp
    'Iterate deeper into the directory hierarchy.
    enumerateSubdirs enumSubFolderPath
  Next
  'Clean up unneeded memory.
  esdTemp = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to determine what type of target is selected.
'https://stackoverflow.com/questions/21035366/how-to-check-the-given-path-is-a-directory-or-file-in-vbscript
Function GetFSElementType(ByVal tPath)
  With CreateObject("Scripting.FileSystemObject") 
    tPath = .GetAbsolutePathName(tPath)
    Select Case TRUE
      Case .FileExists(tPath)   : GetFSElementType = 1
      Case .FolderExists(tPath) : GetFSElementType = 2
      Case Else                : GetFSElementType = 0
  End Select
  End With
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to determine if a target is a file.
'https://stackoverflow.com/questions/21035366/how-to-check-the-given-path-is-a-directory-or-file-in-vbscript
Function IsFile(tPath)
  IsFile = (GetFSElementType(tPath) = 1)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to determine if a target is a folder.
'https://stackoverflow.com/questions/21035366/how-to-check-the-given-path-is-a-directory-or-file-in-vbscript
Function IsFolder(tPath)
  IsFolder = (GetFSElementType(tPath) = 2)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to determine if a target exists.
'https://stackoverflow.com/questions/21035366/how-to-check-the-given-path-is-a-directory-or-file-in-vbscript
Function FSExists(tPath)
  FSExists = (GetFSElementType(tPath) <> 0)
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to prepare the scanner for operation.
Function prepareScanner(priority, target, targetType)
  createLog("Preparing scanner type '" & SanitizeFolder(targetType) & "' with a priority of " & SanitizeFolder(priority) & _
   " against target '" & SanitizeFolder(target) & "'.")
  'Redefine the target array variable incase this is not the first time the function is being called.
  targetArray = Array()
  'Check how much RAM is available, in bytes.
  availableRAM = checkRAM()
  'Validate the priority. It must be a number between 1 and 10.
  'If the priority is not valid, we assume it is 5.
  'The priority is multiplied by 10 and used to represent a percentage of available system RAM.
  If Not IsNumeric(priority) Or priority > 10 Or priority <= 0 Then 
    priority = 5
  End If
  'Here is where we multiply the priority by 10. Also rounds the result into a whole number.
  priorityCoef = Round(priority * 10) 
  'Here we find the percentage of ram that represents our priority percentage. 
  'A priority of 9 allows the Scan-Core to consume 90% of available RAM.
  workerRAMLimit = Round((availableRAM * priorityCoef) / 100) 
  workerChunkSize = Round(workerRAMLimit / chunkCoef)
  'Add the root target to the target array.
  targetArray = push(target, targetArray)
  'If the target is a file, we add it to the target array.
  If LCase(targetType) = "file" Then
    targetArray = enumSubFolder
    'The workerlimit is set as the number of elements in the array.
    'Default is one worker per subdirectory.
    'If the target is a lonely file, it only gets a single worker.
    workerLimit = UBound(targetArray) + 1
  End IF
  'If the file is a registry object, we validate it and add it to the target array.
  If LCase(targetType) = "registry" Then

  End If
  createLog("Allocating " & workerLimit & " workers, each with a memory limit of " & ((workerRamLimit / 1024) / 1024) & _
   "MB and a chunk size of " & ((workerRamLimit / 1024) / 1024) & "MB to be started sequentially utilizing a maximum of " & _
   priorityCoef & "% of available memory."
  prepareScanner = targetArray
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to start a worker.
'Workers perform scan & file operations on targets using resources.
'A worker is a single thread with a designated memory limit and a specific target object.
'workerType can be scanner or janitor.
'targetType can be either "registry" or "file".
'target can be specific registry keys or files specified by path.
'Priority must be an integer between 1 & 10.
Function startWorker(workerType, wTarget, targetType)
  createLog("Starting worker type '" & SanitizeFolder(workerType) & "' against target '" & SanitizeFolder(target) & _
   "' of type '" & SanitizeFolder(targetType) & "'.")
  If LCase(workerType) = "scanner" Then
    If LCase(targetType) = "file" Then 
      'Run the PHP\7.3.8\php.exe binary with cmd.exe, hide the window, don't wait for completion, & 
       'call Scripts\PHP\PHP-AV\scanCore.php script against the target with the specified RAM & chunk settings, without recursion.
      objShell.Run "C:\Windows\System32\cmd.exe /c " & binariesDirectory & "PHP\7.3.8\php.exe " & _
       scriptsDirectory & _ "PHP\scanCore.php " & wTarget & " -m " & workerRAMLimit & " -c " & workerChunkSize " -nr", 0, FALSE
    End If
    If LCase(targetType) = "registry" Then

    End If
  End If

  If LCase(workerType) = "janitor" Then
    If LCase(targetType) = "file" Then 
    
    End If
    If LCase(targetType) = "registry" Then

    End If
  End IF
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to scan the system for infections.
Function smartScan(priority, target, targetType)
  targetArray = prepareScanner(priority, target, targetType)
  createLog("Starting scanner type '" & SanitizeFolder(targetType) & "' with a priority of " & SanitizeFolder(priority) & _
   " against target '" & SanitizeFolder(target) & "'.")
  For targets In targetArray
    On Error Resume Next
    'The following loop sleeps the outer loop for as long as there is insufficient RAM to start a new worker.
    'The loop will check system RAM every <checkWorkerTimer> seconds.
    While checkRam() < workerRAMLimit
      Sleep(checkWorkerTimer)
    Next
    'If the worker limit has not been met then we continue scanning targets.
    If workerCount != workerLimit Then
      startWorker(workerType, targets, targetType)
      'Increment the worker counter.
      workerCount = workerCount + 1
    Else
      createLog("Worker allocation depleted.")
      Exit For
    End If
  Next
End Function
'--------------------------------------------------