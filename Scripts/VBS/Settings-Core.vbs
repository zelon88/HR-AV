'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file is for saving settings submitted by settings.hta to a specific configuration file. 

'--------------------------------------------------
'Define variables for the session.
Dim setting1, setting2, setting3, setting4, setting1Input, setting2Input, setting3Input, _
 setting4Input, readfile
'--------------------------------------------------

'--------------------------------------------------
'Get submitted setting values from the DOM.
Set setting1Input = document.getElementById("setting1")
Set setting2Input = document.getElementById("setting2")
Set setting3Input = document.getElementById("setting2")
Set setting4Input = document.getElementById("setting4")
'--------------------------------------------------

'--------------------------------------------------
'Set the directory/file locations for the settings.dat files.
setting1 = cacheDirectory & "setting1.dat"
setting2 = cacheDirectory & "setting2.dat"
setting3 = cacheDirectory & "setting3.dat"
setting4 = cacheDirectory & "setting4.dat"
'--------------------------------------------------

'--------------------------------------------------
'Save the new settings to the settings.dat files.
If objFSO.FileExists(setting1) Then
  Set readfile = objFSO.OpenTextFile(setting1, 1)
  If Not readfile.AtEndOfStream Then
    scriptLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(setting2) Then
  Set readfile = objFSO.OpenTextFile(setting2, 1)
  If Not readfile.AtEndOfStream Then
    helpLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(setting3) Then
  Set readfile = objFSO.OpenTextFile(setting3, 1)
  If Not readfile.AtEndOfStream Then
    maintLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(setting4) Then
  Set readfile = objFSO.OpenTextFile(setting4, 1)
  If Not readfile.AtEndOfStream Then
    logLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If
'--------------------------------------------------

'--------------------------------------------------
'Reset the DOM with the newest settings.
setting1Input.value = setting1
setting2Input.value = setting2
setting3Input.value = setting3
setting4Input.value = setting4
'--------------------------------------------------