'HR-AV Desktop Antivirus
'https://github.com/zelon88/HR-AV
'https://github.com/zelon88

'Author: Justin Grimes
'Date: 10/28/2019
'<3 Open-Source

'Unless Otherwise Noted, The Code Contained In This Repository Is Licensed Under GNU GPLv3
'https://www.gnu.org/licenses/gpl-3.0.html

'This file contains the master VBS engine configuration.
'If this file is lost or destroyed main application settings will be lost!


'--------------------------------------------------
'Define global variables for the session.
Option Explicit

Dim version, uiVersion, helpLocSetting, appName, developerName, developerURL, dieOnInstallationError, windowHeight, windowWidth, _
appDownloadURL, defDownloadURL, realTimeProtectionEnabled, runInBackground, registryMonitorInterval, ramsomwareDefenderInterval, _
accessibilityDefenderInterval, storageMonitorInterval, resourceMonitorInterval, infrastructureCheckupInterval, infrastructureHeartbeatInterval, _
workstationUSBMonitorEnabled, registryMonitorEnabled, ransomwareDefenderEnabled, infrastructureHeartbeatEnabled, infrastructureCheckupEnabled, _
accessibilityDefenderEnabled, storageMonitorEnabled, resourceMonitorEnabled, DEBUGMODE
'--------------------------------------------------


'--------------------------------------------------
'Application-Development Related Variables

  'These string values in this section should only be modified by your application distributor.
  version = "v0.8.9" 
  uiVersion = "v1.2"
  helpLocSetting = "https://github.com/zelon88/HR-AV"
  appDownloadURL = "https://github.com/zelon88/HR-AV/archive/master.zip"
  defDownloadURL = "https://github.com/zelon88/HR-AV_Defs/archive/master.zip"
  appName = "HR-AV"
  developerName = "Justin Grimes"
  developerURL = "https://github.com/zelon88"
  DEBUGMODE = FALSE
'--------------------------------------------------


'--------------------------------------------------
'Environment Related Variables

  'Set 'dieOnInstallationError' to 'TRUE' to kill the application instead of running the installer.
  'Useful for deploying via GPO to prevent larger installation mistakes.
  'Must be boolean, TRUE or FALSE.
  dieOnInstallationError = TRUE

  'The 'windowHeight' is the height of the main application window, in pixels.
  'In pixels. Must be an integer.
  windowHeight = 660

  'The 'windowWidth' is the width of the main application window, in pixels.
  'In pixels. Must be an integer.
  windowWidth = 600

  'Set 'realTimeProtectionEnabled' to 'TRUE' to enable the Real-Time-Protection engine (RealTime-Core.vbs) and all of it's services & shceduled tasks.
  'Set 'realTimeProtectionEnabled' to 'FALSE' to disable the Real-Time-Protection engine (RealTime-Core.vbs) and all of it's services & shceduled tasks.
  'Individual Real-Time-Protection services and tasks can still be run manually at any time.
  'Must be boolean, TRUE or FALSE.
  realTimeProtectionEnabled = TRUE

  'Set 'runInBackground' to 'TRUE' to allow the Real-Time-Protection engine (RealTime-Core.vbs) to run in the background, separately from the main application.
  'Set 'runInBackground' to 'FALSE' to prevent the Real-Time-Protection engine (RealTime-Core.vbs) from running when the main application is closed.
  'Must be boolean, TRUE or FALSE.
  runInBackground = TRUE
'--------------------------------------------------


'--------------------------------------------------
'Service-Specific Variables
  'Allow background services to run. 
  'Services that are enabled will be started and enforced automatically.

   'Workstation_USB_Monitor watches the USB ports on the local machine for new devices and which may indicate potentially malicious activity.
   'Cloned from 'https://github.com/zelon88/Workstation_USB_Monitor' and heavily modified for use in this application.
    'To enable "Workstation_USB_Monitor" to run in the background as managed by this application, set 'workstationUSBMonitorEnabled' to 'TRUE'.
    workstationUSBMonitorEnabled = TRUE
'--------------------------------------------------


'--------------------------------------------------
'Task-Specific Variables
  'Intervals (in minutes) for real-time-tasks to run. 
  'Minimum interval is 1 (one) minute.
  'Intervals set below 1m will be enforced every minute.

   'Registry Monitor watches the Windows Registry for changes which could indicate malicious activity. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Registry_Monitor' and heavily modified for use in this application.
    'To enable "Registry_Monitor.vbs" to run at scheduled intervals, set 'registryMonitorEnabled' to 'TRUE'.
    registryMonitorEnabled = TRUE
    'Set the 'registryMonitorInterval' interval, in minutes, that 'Registry_Monitor.vbs' will be triggered by the RealTime-Core.
    registryMonitorInterval = 10  
  
   'Ransomware Defender watches the local filesystem for suspicious changes which could indicate Ransomware activity. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Ransomware_Defender' and heavily modified for use in this application.
    'To enable "Ransomware_Defender.vbs" to run at scheduled intervals, set 'ransomwareDefenderEnabled' to 'TRUE'.
    ransomwareDefenderEnabled = TRUE
    'Set the 'ramsomwareDefenderInterval' interval, in minutes, that 'Ransomware_Defender.vbs' will be triggered by the RealTime-Core.
    ramsomwareDefenderInterval = 5
  
   'Accessibility Monitor watches the local accessibility tools for changes which could indicate backdoor attacker activity. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Accessibility_Tools_Utilimon_Defender' and heavily modified for use in this application.
    'To enable "Accessibility_Defender.vbs" to run at scheduled intervals, set 'accessibilityDefenderEnabled' to 'TRUE'.
    accessibilityDefenderEnabled = TRUE
    'Set the 'accessibilityDefenderInterval' interval, in minutes, that 'Accessibility_Defender.vbs' will be triggered by the RealTime-Core.
    accessibilityDefenderInterval = 60
  
   'Storage Monitor watches the local filesystems as a whole for changes which could indicate malicious activity. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Storage_Monitor' and heavily modified for use in this application.
    'To enable "Storage_Monitor.vbs" to run at scheduled intervals, set 'storageMonitorEnabled' to 'TRUE'.
    storageMonitorEnabled = TRUE
    'Set the 'storageMonitorInterval' interval, in minutes, that 'Storage_Monitor.vbs' will be triggered by the RealTime-Core.
    storageMonitorInterval = 10
  
   'Resource Monitor watches the local system as a whole for high resource consumption which could indicate malicious activity. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Resource_Monitor' and heavily modified for use in this application.
    'To enable "Resource_Monitor.vbs" to run at scheduled intervals, set 'resourceMonitorEnabled' to 'TRUE'.
    resourceMonitorEnabled = TRUE
    'Set the 'resourceMonitorInterval' interval, in minutes, that 'Resource_Monitor.vbs' will be triggered by the RealTime-Core.
    resourceMonitorInterval = 7

   'Infrastructure Heartbeat watches network endpoints for online status which could indicate action is required. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Infrastructure_Heartbeat' and heavily modified for use in this application.
    'To enable "Infrastructure_Heartbeat.vbs" to run at scheduled intervals, set 'infrastructureHeartbeatEnabled' to 'TRUE'.
    infrastructureHeartbeatEnabled = TRUE
    'Set the 'infrastructureHeartbeatInterval' interval, in minutes, that 'Infrastructure_Heartbeat.vbs' will be triggered by the RealTime-Core.
    infrastructureHeartbeatInterval = 15
  
   'Infrastructure Checkup performs periodic diagnostic checks on the local system, which could reveal potential security risks or othere indicators of compomise. 
   'When suspicious/malicious activity is detected a log file is created and an email is sent.
   'Cloned from 'https://github.com/zelon88/Infrastructure_Checkup' and heavily modified for use in this application.
    'To enable "Infrastructure_Checkup.vbs" to run at scheduled intervals, set 'infrastructureCheckupEnabled' to 'TRUE'.
    infrastructureCheckupEnabled = TRUE
    'Set the 'infrastructureCheckupInterval' interval, in minutes, that 'Infrastructure_Checkup.vbs' will be triggered by the RealTime-Core.
    infrastructureCheckupInterval = 10
'--------------------------------------------------

