'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 9/4/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file contains the master VBS engine configuration.
'If this file is lost or destroyed main application settings will be lost!

'--------------------------------------------------
'Define global variables for the session.
Option Explicit

Dim version, uiVersion, helpLocSetting, appName, developerName, developerURL, dieOnInstallationError, windowHeight, windowWidth, _
appDownloadURL, defDownloadURL, realTimeProtectionEnabled
'--------------------------------------------------

'--------------------------------------------------
'Application Related Variables
version = "v0.8.2" 
uiVersion = "v1.2"
helpLocSetting = "https://github.com/zelon88/HR-AV"
appDownloadURL = "https://github.com/zelon88/HR-AV/archive/master.zip"
defDownloadURL = "https://github.com/zelon88/HR-AV_Defs/archive/master.zip"
appName = "HR-AV"
developerName = "Justin Grimes"
developerURL = "https://github.com/zelon88"
dieOnInstallationError = TRUE
windowHeight = 660
windowWidth = 600
realTimeProtectionEnabled = TRUE
'--------------------------------------------------