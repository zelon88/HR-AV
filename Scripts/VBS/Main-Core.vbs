'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 9/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file contains the main logic that utilizes the functions and code specified in other Core files.
'This file requires Config.vbs, UI-Core.vbs and App-Core.vbs.

'--------------------------------------------------
Option Explicit

Dim realTimeCoreResults
'--------------------------------------------------

'--------------------------------------------------
'The main logic of the application. The functional entry point for execution.
'Requires functions and variables defined in Config.vbs, UI-Core.vbs, and App-Core.vbs.
'This script is to be run from an HTA which has already loaded the scripts listed above into memory.

'Verify the application is installed to the Program Files directory.
'Fire the installation wizard if not.
If Not isInProgramFiles() Then
  If verifyDirectories() Then
    If verifyInstallation() Then
      PrintGracefully appName & " - " & "Installation Wizard", "Installation Complete!", "vbOKOnly"
    End If
  End If
End If

'Check that the script is running as the HRAV admin user, or that the session is elevated.
If Not isUserHRAV() And Not isUserAdmin() Then 
  restartAsAdmin()
End If

'Verify that required cache files exist and set objects for them.
verifyCache()

'Check if the Real-Time Protection engine needs to be started and start it if needed.
If realTimeProtectionEnabled Then
  If DateDiff("n", oRTPCacheFile1.DateLastModified, Now) > 2 Then
    realTimeCoreResults = SystemBootstrap(realTimeCoreFile, "", FALSE)
  End If
End If
'--------------------------------------------------