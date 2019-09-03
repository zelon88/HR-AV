'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AVe 
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 8/24/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file contains the task manager for real-time protection.

'--------------------------------------------------
Option Explicit

Dim usbMonitorEnabled, registryMonitorEnabled, ransomwareDefenderEnabled, accessibilityDefenderEnabled, storageMonitorEnabled, _
 resourceMonitorEnabled, realTimeProtectionError, infrastructureCheckupEnabled, infrastructureHeartbeatEnabled, itRemindersEnabled, _
 itRemindersDue, infrastructureHeartbeatdue, infrastructureCheckupDue, resourceMonitorDue, storageMonitorDue, accessibilityDefenderDue, _
 ransomwareDefenderDue, registryDefenderDue, usbMonitorRunning

realTimeProtectionError = FALSE

If realTimeProtectionEnabled = TRUE Then 

  If usbMonitorEnabled = TRUE And usbMonitorRunning = TRUE Then

  End If
  If registryMonitorEnabled = TRUE And registryDefenderDue = TRUE Then

  End If
  If ransomwareDefenderEnabled = TRUE And ransomwareDefenderDue = TRUE Then

  End If
  If accessibilityDefenderEnabled = TRUE And accessibilityDefenderDue = TRUE Then

  End If
  If storageMonitorEnabled = TRUE And storageMonitorDue = TRUE Then

  End If
  If resourceMonitorEnabled = TRUE And resourceMonitorDue = TRUE Then

  End If
  If infrastructureCheckupEnabled = TRUE And infrastructureCheckupDue = TRUE Then

  End If
  If infrastructureHeartbeatEnabled = TRUE And infrastructureHeartbeatdue = TRUE Then

  End If
  If itRemindersEnabled = TRUE And itRemindersDue = TRUE Then

  End If
End If