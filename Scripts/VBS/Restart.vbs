'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/29/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file is used during initial application setup to restart the application in an orderly fashion.
'This file can also be run to kill running HR-AV instances and start fresh ones.

Option Explicit

Dim objShell, cProcessList, oWMISrvc, dProcess, sCommLine, fullScriptName

Set objShell = CreateObject("WScript.Shell")
Set oWMISrvc = GetObject("winmgmts:")
Set cProcessList = oWMISrvc.ExecQuery("select * from win32_process where Name = 'HR-AV'")

'Scan for and kill any running instances of the script.
For Each dProcess in cProcessList
  sCommLine = Trim(LCase(dProcess.CommandLine))
  If InStr(sCommLine, "HR-AV") >= 0 Then
    dProcess.Terminate()
  End If
Next

'Communicate our intentions to the user, then wait 10 seconds before restarting the script in Program Files.
MsgBox "The application will restart in 10 seconds.", 0, "HR-AV- Installation Wizard"
Wscript.Sleep 10000
objShell.Run """C:\Program Files\HR-AV\HR-AV.hta""", 1, FALSE