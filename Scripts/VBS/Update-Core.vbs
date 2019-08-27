'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/23/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file is for updating the HR-AV application and it's definition files.

Option Explicit

Dim xHttp, bStrm

Set xHttp = createobject("Microsoft.XMLHTTP")
Set bStrm = createobject("Adodb.Stream")

'--------------------------------------------------
Function downloadAppUpdate()
  xHttp.Open "GET", appDownloadURL, False
  xHttp.Send
  With bStrm
    .type = 1 'Binary.
    .open
    .write xHttp.responseBody
    .savetofile tempDirectory & "update.zip", 2 'Overwrite.
  End With
End Function
'--------------------------------------------------

'--------------------------------------------------
Function installAppUpdate()

End Function
'--------------------------------------------------

'--------------------------------------------------
Function cleanAppUpdate()

End Function
'--------------------------------------------------

'--------------------------------------------------
Function checkAppCompat()

End Function
'--------------------------------------------------

'--------------------------------------------------
Function downloadDefUpdate()
  xHttp.Open "GET", defDownloadURL, False
  xHttp.Send
  With bStrm
    .type = 1 'Binary.
    .open
    .write xHttp.responseBody
    .savetofile tempDirectory & "defs.zip", 2 'Overwrite.
  End With
End Function
'--------------------------------------------------

'--------------------------------------------------
Function installDefUpdate()

End Function
'--------------------------------------------------

'--------------------------------------------------
Function cleanDefUpdate()

End Function
'--------------------------------------------------

'--------------------------------------------------
Function checkDefCompat()

End Function
'--------------------------------------------------