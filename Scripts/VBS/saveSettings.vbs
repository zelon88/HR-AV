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
Dim scriptLocInput, helpLocInput, maintLocInput, logLocInput, scriptLocSettings, helpLocSettings, _
 logLocSettings, readfile
'--------------------------------------------------

'--------------------------------------------------
'Get submitted setting values from the DOM.
Set scriptLocInput = document.getElementById("setting1")
Set helpLocInput = document.getElementById("setting2")
Set maintLocInput = document.ge tElementById("setting2")
Set logLocInput = document.getElementById("setting4")
'--------------------------------------------------

'--------------------------------------------------
'Set the directory/file locations for the settings.dat files.
scriptLocSettings = cacheDirectory & "scriptLocSettings.dat"
helpLocSettings = cacheDirectory & "helpLocSettings.dat"
maintLocSettings = cacheDirectory & "maintLocSettings.dat"
logLocSettings = cacheDirectory & "logLocSettings.dat"
'--------------------------------------------------

'--------------------------------------------------
'Save the new settings to the settings.dat files.
If objFSO.FileExists(scriptLocSettings) Then
  Set readfile = objFSO.OpenTextFile(scriptLocSettings, 1)
  If Not readfile.AtEndOfStream Then
    scriptLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(helpLocSettings) Then
  Set readfile = objFSO.OpenTextFile(helpLocSettings, 1)
  If Not readfile.AtEndOfStream Then
    helpLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(maintLocSettings) Then
  Set readfile = objFSO.OpenTextFile(maintLocSettings, 1)
  If Not readfile.AtEndOfStream Then
    maintLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

If objFSO.FileExists(logLocSettings) Then
  Set readfile = objFSO.OpenTextFile(logLocSettings, 1)
  If Not readfile.AtEndOfStream Then
    logLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If
'--------------------------------------------------

'--------------------------------------------------
'Reset the DOM with the newest settings.
scriptLocInput.value = scriptLocSetting
helpLocInput.value = helpLocSetting
maintLocInput.value = maintLocSetting
logLocInput.value = logLocSetting
'--------------------------------------------------