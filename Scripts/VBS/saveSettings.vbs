'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file is for saving settings submitted by settings.hta to a specific configuration file. 

dim scriptLocInput, helpLocInput, maintLocInput, logLocInput, cameraLocInput, _
 scriptLocSettings, helpLocSettings, logLocSettings, cameraLocSettings, readfile

set scriptLocInput = document.getElementById("scriptLocInput")
set helpLocInput = document.getElementById("helpLocInput")
set maintLocInput = document.ge tElementById("maintLocInput")
set logLocInput = document.getElementById("logLocInput")
set cameraLocInput = document.getElementById("cameraLocInput")

scriptLocSettings = "Cache\scriptLocSettings.dat"
helpLocSettings = "Cache\helpLocSettings.dat"
maintLocSettings = "Cache\maintLocSettings.dat"
logLocSettings = "Cache\logLocSettings.dat"
cameraLocSettings = "Cache\cameraLocSettings.dat"

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

If objFSO.FileExists(cameraLocSettings) Then
  Set readfile = objFSO.OpenTextFile(cameraLocSettings, 1)
  If Not readfile.AtEndOfStream Then
    cameraLocSetting = readfile.ReadAll
  End If
  readfile.Close
End If

scriptLocInput.value = scriptLocSetting
helpLocInput.value = helpLocSetting
maintLocInput.value = maintLocSetting
logLocInput.value = logLocSetting
cameraLocInput.value = cameraLocSetting