'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 12/18/2019
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
 charArr, charArr2, tmpChar, tmpChar2, strToClean1, strToClean2, strEventInfo, objLogFile, whoamiOutput, _
 obj, x, i, tempArray, rpCounter, pcs, oWMISrvc, errorNumber, errorMessage, quietly, cantError, windowNote, message, _ 
 typeMsg, dontContinue, sBinaryToRun, sCommand, sAsync, srun, strHRAVUserName, strHRAVPassword, result, resultSet, _
 stempFile, sasync1, stempData, searchScripts, scriptsToSearch, procSearch, procsToSearch, strComputer, objRAMService, _
 availableRAMGB, commitLimitRAMBytes, commitLimitRAMKB, commitLimitRAMMB, commitLimitRAMGB, committedRAMBytes, _
 committedRAMKB, committedRAMMB, committedRAMGB, objDrives, objDrive, availableRAMBytes, availableRAMKB, _
 eDelimiter, eString, eLimit, fgcPath, objFGCFile, exCounter, nexCounter, newInfection, infectionArray, exception, _
 exceptionFile, exceptionCSVData, workeType, targetType, memoryLimit, availableRAMMB, exceptionArray, target, _
 excepptionArray, priority, chunkCoef, priorityCoef, workerRAMLimit, availableRAM, workerChunkSize, workerLimit, _ 
 enumFolder, enumSubFolder, mValEl, mArray, mValue, tPath, KB, MB, GB, workerCount, targetArray, workerType, _
 exceptionData, configFile, targets, wTarget, checkWorkerTimer, objFolder, objItem, objSh, strPath, execString, _
 pathInput, priorityInput, strComputerName, objNetwork, filePathInput, objFile, targetArray2, target2, wTargetWrappers, _
 targetTemp, targetSubDirs, oenumFolder, enumSubFolderPath, esdTemp

'Commonly Used Objects.
Set objShell = CreateObject("WScript.Shell")
Set objSh  = CreateObject("Shell.Application")
Set oWMISrvc = GetObject("winmgmts:")
Set objRAMService = GetObject("winmgmts:\\.\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ObjNetwork = CreateObject("WScript.Network")
'Environment Related Variables.
sesID = Int(Rnd * 10000000)
KB = 1024
MB = KB * 1024
GB = MB * 1024
priorityCoef = 1
chunkCoef = 3
checkWorkerTimer = 3
strComputer = "."
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
targetTemp = ""
targetSubdirs = ""
exceptionArray = Array()
targetArray = Array()
'Time Related Variables.
humanDate = Trim(FormatDateTime(Now, vbShortDate)) 
logDate = Trim(Replace(humanDate, "/", "-"))
humanTime = Trim(FormatDateTime(Now, vbLongTime))
logTime = Trim(Replace(Replace(humanTime, ":", "-"), " ", ""))
humanDateTime = Trim(humanDate & " " & humanTime)
logDateTime = Trim(logDate & "_" & logTime)
strComputerName = ObjNetwork.ComputerName
'Directory Related Variables.
currentDirectory = Replace(Replace(Trim(objFSO.GetAbsolutePathName(strComputer)), "\\", "\"), "HR-AV\Pages", "HR-AV")
configFile = Replace(currentDirectory & "\Config\Config.vbs", "\\", "\")
scriptsDirectory = Replace(currentDirectory & "\Scripts\", "\\", "\")
vbsScriptsDirectory = Replace(scriptsDirectory & "\VBS\", "\\", "\")
binariesDirectory = Replace(currentDirectory & "\Binaries\", "\\", "\")
cacheDirectory = Replace(currentDirectory & "\Cache\", "\\", "\")
tempDirectory = Replace(currentDirectory & "\Temp\", "\\", "\")
pagesDirectory = Replace(currentDirectory & "\Pages\", "\\", "\")
mediaDirectory = Replace(currentDirectory & "\Media\", "\\", "\")
logsDirectory = Replace(currentDirectory & "\Logs\", "\\", "\")
reportsDirectory = Replace(currentDirectory & "\Reports\", "\\", "\")
resourcesDirectory = Replace(currentDirectory & "\Resources\", "\\", "\")
logFilePath = Replace(Trim(logsDirectory & "RTP-Log_" & logDate), "\\", "\")
exceptionDirectory = Replace(currentDirectory & "\Exceptions\", "\\", "\")
excepptionFile = Replace(exceptionDirectory & "Exception_List.csv", "\\", "\")
tempFile = tempDirectory & "SC-STemp.txt"
stempFile = tempDirectory & "SC-Temp.txt"
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
  'Determine who is executing the current script.
  whoamiOutput = SystemBootstrap("whoami", "", FALSE)
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
  If Not objFSO.FileExists(tempFile) Then
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
'A function to push data onto an existing array. Similar to PHP's array_push().
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
    availableRAMBytes = Round(result.AvailableBytes,3)
    checkRAM = availableRAMBytes
    availableRAMKB = Round(availableRAMBytes / KB,3)
    availableRAMMB = Round(availableRAMBytes / MB,3)
    availableRAMGB = Round(availableRAMBytes / GB,3)
    'Set variables for current CommitLimit. In bytes, KB, MB, & GB.
    commitLimitRAMBytes = Round(result.CommitLimit,3)
    commitLimitRAMKB = Round(commitLimitRAMBytes / KB,3)
    commitLimitRAMMB = Round(commitLimitRAMBytes / MB,3)
    commitLimitRAMGB = Round(commitLimitRAMBytes / GB,3)
    'Set variables for current CommitLimit. In bytes, KB, MB, & GB.
    committedRAMBytes = Round(result.CommittedBytes,3)
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
  Set objDrives = objFSO.Drives
  enumerateDrives = Array()
  tempArray = Array()
  createLog("Enumerating mounted disk drive volumes.")
  'Iterate through each drive volume on the system and add it to the tempArray().
  For Each objDrive in objDrives
    push tempArray, objDrive.DriveLetter & ":\\"
  Next
  'Copy the tempArray() to the enumerateDrives() array.
  enumerateDrives = tempArray
  'Clean up unneeded memory.
  objDrive = NULL
  objDrives = NULL
  tempArray = NULL
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
    objFSO.CreateTextFile exceptionFile, TRUE
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
'A function to build an array of all subdirectories of a target folder.
'https://stackoverflow.com/questions/1433785/vbscript-to-iterate-through-set-level-of-subfolders
Function enumerateSubdirs(enumFolder) 
  Set oenumFolder = objFSO.GetFolder(enumFolder)
  enumerateSubdirs = Array()
  'Iterate through each subfolder of the "enumFolder".
  For Each enumSubFolder in oenumFolder.SubFolders
    enumSubFolderPath = enumSubFolder.Path
    'Add the current path to the "targetArray".
    esdTemp = push(enumerateSubdirs, enumSubFolderPath)
    'Iterate deeper into the directory hierarchy.
    enumerateSubdirs enumSubFolderPath
  Next
  enumerateSubdirs = esdTemp
  'Clean up unneeded memory.
  esdTemp = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to select a folder for scanning. Also used to scan the computer when myStartFolder is left blank.
'Opens a "Select Folder" dialog and will return the fully qualified path of the selected folder.
'https://www.robvanderwoude.com/vbstech_ui_selectfolder.php
Function selectFolder() 
  selectFolder = vbNull
  'Create a dialog object.
  Set objFolder = objSh.BrowseForFolder(0, "Select Folder", 0, "C:\")
  'Return the path of the selected folder.
  If IsObject(objfolder) Then selectFolder = objFolder.Self.Path
  'Clean up unneeded memory.
  Set objFolder = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to select a file for scanning.
'Opens a "Select File" dialog and will return the fully qualified path of the selected file.
'https://www.robvanderwoude.com/vbstech_ui_selectfolder.php
Function selectFile()   
  selectFile = vbNull
  'Create a dialog object.
  Set objFile = objSh.BrowseForFolder(0, "Select File", &H4000, "C:\")
  'Return the path of the selected folder.
  If IsObject(objfile) Then selectFile = objFile.Self.Path
  'Clean up unneeded memory.
  Set objFile = NULL
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
  targetArray2 = Array()
  'If the target is blank we set the target array to an array of disk volumes attached to the computer.
  If target = "" Then
    targetArray = enumerateDrives()
    For Each targetTemp in targetArray
      targetSubDirs = enumerateSubdirs(targetTemp)
      targetArray = push(targetSubDirs, targetArray)
      targetSubDirs = ""
    Next 
    targetTemp = ""
  End If
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
  'If the target is a file, we add it to the target array.
  If LCase(targetType) = "file" And target <> "" Then
    'Add the root target to the target array.
    targetArray = push(target, targetArray)
    targetArray2 = enumerateSubdirs(target)
    For Each target2 in targetArray2
      push targetArray, target2
    Next
    'Clean up unneeded memory.
    targetArray2 = NULL
  End IF
  'The workerlimit is set as the number of elements in the array.
  'Default is one worker per subdirectory.
  'If the target is a lonely file, it only gets a single worker.
  workerLimit = UBound(targetArray) + 1
  'If the file is a registry object, we validate it and add it to the target array.
  If LCase(targetType) = "registry" Then
  End If
  createLog("Allocating " & workerLimit & " workers, each with a memory limit of " & ((workerRamLimit / 1024) / 1024) & _
   "MB and a chunk size of " & ((workerRamLimit / 1024) / 1024) & "MB to be started sequentially utilizing a maximum of " & _
   priorityCoef & "% of available memory.")
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
  wTarget = Replace(Replace(wTarget, "\\", "\"), "\\", "\")
  createLog("Starting worker type '" & SanitizeFolder(workerType) & "' against target '" & SanitizeFolder(wTarget) & _
   "' of type '" & SanitizeFolder(targetType) & "'.")
  If LCase(workerType) = "scanner" Then
    If LCase(targetType) = "file" Then 
      'If the wTarget is a drive letter we don't need to use parenthesis to encapsulate it for PHP-AV.
      If Len(wTarget) < 4 Then
        wTargetWrappers = ""
      Else 
        wTargetWrappers = Chr(34)
      End If 
      'Run the PHP\7.3.8\php.exe binary with cmd.exe, hide the window, don't wait for completion, & 
       'call Scripts\PHP\PHP-AV\scanCore.php script against the target with the specified RAM & chunk settings, without recursion.
      execString = """" & scriptsDirectory & "PHP\PHP-AV\scanCore.php"" " & wTargetWrappers & wTarget & wTargetWrappers & _
       " -m " & workerRAMLimit & " -c " & workerChunkSize & " -nr"
      'MsgBox execString
      Bootstrap "PHP\7.3.8\php.exe", execString, FALSE
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
Function smartScan(priority, target, targetType, workerType)
  targetArray = prepareScanner(priority, target, targetType)
  createLog("Starting scanner engine type '" & SanitizeFolder(targetType) & "' with a priority of " & SanitizeFolder(priority) & _
   " against target '" & SanitizeFolder(target) & "'.")
  For Each targets In targetArray
    'The following loop sleeps the outer loop for as long as there is insufficient RAM to start a new worker.
    'The loop will check system RAM every <checkWorkerTimer> seconds.
    Do While checkRam() < workerRAMLimit
      Sleep(checkWorkerTimer)
    Loop
    'If the worker limit has not been met then we continue scanning targets.
    If workerCount <= workerLimit Then
      startWorker workerType, targets, targetType
      'Increment the worker counter.
      workerCount = workerCount + 1
    Else
      createLog("Worker allocation depleted.")
      Exit For
    End If
  Next
  createLog("Scan Complete.")
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to update the status bar HTML on the scan page and also log the status change.
'Changes the scanStatusDiv used in scan-related pages.
'Note that you can only set the newStatusMessage ONCE PER EXECUTION! 
'If you set it twice only the last one interpreted will be visible to the user.
Function updateStatusBar(newStatusMessage)
  newStatusMessage = SanitizeFolder(newStatusMessage)
  scanStatusDiv.innerHTML = newStatusMessage
  createLog(newStatusMessage)
  newStatusMessage = NULL
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to scan a selected folder.
'Fire this function AFTER the selectFolder() functon and priorityInput have been filled out via the UI.
'https://community.spiceworks.com/topic/456927-drop-down-menu-in-hta
Function scanFolder()
  pathInput = SanitizeFoldeR(folderPathInput.value)  
  priorityInput = SanitizeFolder(priorityInput.value)
  If Not pathInput = vbNull Then
    updateStatusBar("Scanning Folder: """ & pathInput & """") 
    smartScan priorityInput, pathInput, "file", "scanner"
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to scan the entire local computer.
'Fire this function AFTER the selectFile() functon and priorityInput have been filled out via the UI.
'https://community.spiceworks.com/topic/456927-drop-down-menu-in-hta
Function scanFile() 
  filePathInput = SanitizeFoldeR(filePathInput.value)
  priorityInput = SanitizeFolder(priorityInput.value)
  If Not pathInput = vbNull Then
    updateStatusBar("Scanning File: """ & filePathInput & """") 
    smartScan priorityInput, filePathInput, "file", "scanner"
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to scan the entire local computer.
'Fire this function AFTER the priorityInput has been filled out via the UI.
'https://community.spiceworks.com/topic/456927-drop-down-menu-in-hta
Function scanComputer()  
  priorityInput = SanitizeFolder(document.getElementById("priorityInput").Value)
  updateStatusBar("Scanning computer: """ & strComputerName & """")
  smartScan priorityInput, "", "file", "scanner"
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to scan the entire local registry.
'Fire this function AFTER the priorityInput has been filled out via the UI.
'https://community.spiceworks.com/topic/456927-drop-down-menu-in-hta
Function scanRegistry()  
  priorityInput = SanitizeFolder(priorityInput.value)
  updateStatusBar("Scanning registry: """ & strComputerName & """")
  smartScan priorityInput, "", "registry", "scanner"
End Function
'--------------------------------------------------