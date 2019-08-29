'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/24/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file contains the main logic that utilizes the functions and code specified in other Core files.

'--------------------------------------------------
Option Explicit

If verifyDirectories() = TRUE Then
  If verifyInstallation() = TRUE Then
    msgbox "done"
  End If
End If