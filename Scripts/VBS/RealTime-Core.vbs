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
 ransomwareDefenderDue, registryDefenderDue, usbMonitorRunning, realTimeSleep, servicesRunning, testServicesRunning, serviceRequired, _
 service, validService, serviceCheck, pcs, rpCounter, currentRunningProcs, runningServices, reqdServiceCount

validServices = Array("USB_Monitor")
realTimeProtectionError = FALSE
realTimeSleep = 10000 '10s

'A function to enumerate running processes into an array. 
'Each array element contains a CSV containing a PID, process name, and executable path.
Function enumerateRunningProcesses() 
  enumerateRunningProcesses = Array()
  rpCounter = 0
  For each pcs in oWMISrvc.InstancesOf("Win32_Process")
    enumerateRunningProcesses(rpCounter) = pcs.ProcessID & "," & pcs.Name & "," & pcs.ExecutablePath
    rpCounter = rpCounter + 1
  Next
  rpCounter = NULL
End Function

'A function to check that all HR-AV related background services are running.
Function servicesRunning() 
  serviceCheck = FALSE
  runningServices = 0
  reqdServiceCount = UBound(servicesEnabled) + 1
  currentRunningProcs = enumerateRunningProcesses()
  For Each serviceRequired In servicesEnabled
    serviceCheck = FALSE
    For Each validService In validServices
      serviceCheck = FALSE
      If serviceRequired = validService Then
        serviceCheck = TRUE
      End If
      If serviceCheck = TRUE Then
        For Each currentProc In currentRunningProcs
          If InStr(serviceRequired, currentProc) = 0 Then
            runningServices = runningServices + 1
          End If
        Next
      End If
    Next
  Next
  If Not runningServices = reqdServiceCount Then
    serviceCheck = FALSE
  Else
    serviceCheck = TRUE
  End If 
  servicesRunning = serviceCheck
End Function

Function startServices()

End Function

If realTimeProtectionEnabled Then 
  If Not servicesRunning() Then
    createLog("Attempting to start services.")
    testServicesRunning = startServices()
    If Not testServicesRunning Then
      createLog("Could not start services!")
    End If
  End If
  If usbMonitorEnabled And usbMonitorRunning Then

  End If
  If registryMonitorEnabled And registryDefenderDue Then

  End If
  If ransomwareDefenderEnabled And ransomwareDefenderDue Then

  End If
  If accessibilityDefenderEnabled And accessibilityDefenderDue Then

  End If
  If storageMonitorEnabled And storageMonitorDue Then

  End If
  If resourceMonitorEnabled And resourceMonitorDue Then

  End If
  If infrastructureCheckupEnabled And infrastructureCheckupDue Then

  End If
  If infrastructureHeartbeatEnabled And infrastructureHeartbeatdue Then

  End If
  If itRemindersEnabled And itRemindersDue Then

  End If
End If