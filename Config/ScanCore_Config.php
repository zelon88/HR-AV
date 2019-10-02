<?php

// This file contains the configuration data for the HR-AV Server application.
// Make sure to fill out the information below 100% accuratly BEFORE you attempt to run
// any HR-AV Server application scripts. Severe filesystem damage could result.

// BE SURE TO FILL OUT ALL INFORMATION ACCURATELY !!!
// PRESERVE ALL SYNTAX AND FORMATTING !!!
// SERIOUS FILESYSTEM DAMAGE COULD RESULT FROM INCORRECT DATABASE OR DIRECTORY INFO !!!

// htts://github.com/zelon88/HR-AV
// / ------------------------------


// / ------------------------------
// / License Information ...
  // / To continue, please accept the included GPLv3 license by changing the following 
  // / variable to '1'. By changing the '$Accept_GPLv3_OpenSource_License' variable to '1'
  // / you aknowledge that you have read and agree to the terms of the included LICENSE file.
$Accept_GPLv3_OpenSource_License = '1';
// / ------------------------------

// / ------------------------------  
// / Security Information ... 
  // / HR-AV can run on a local machine or on a network as a server to
  // / serve clients over http using standard web browsers.

  // / Secret Salts.
    // / Change these to something completely random and keep it a secret. Store your $Salts
    // / in hardcopy form or an encrypted drive in case of emergency.
    // / IF YOU LOSE YOUR SALTS YOU WILL BE UNABLE TO DECODE USER ID'S AFTER AN EMEREGENCY.
$Salts1 = 'som#@ethin5gSoRanDoewMdfgdfThatNobody_Will_evar+guess+itgefgjfdsjgdfgdgdfgfdsfgdasfdas';
$Salts2 = 'g@dfsgdfs3gsdfsomgdrwefwfgethingSoRanDoMThatNobody_Will_evar+guess+it';
$Salts3 = 'somethi4ngSoRanDoMfsdfsTh9atNobodygdfgsdfgfs3243234$^534_Will_evar+guess+it';
$Salts4 = 'somet#hingSoR2anDoMTherweatNobody;lk;jlfrdwas5l_Will_evar+guess+iwt';
$Salts5 = 'somethingSoRfsbm.il)(*&^%&#^GIFSKGFHGNggdfsig2423gh_Will_evar+guess+it';
$Salts6 = 'somethingSo1RanDoMThatNobodyawrsalafsadfsdfuaoe4th39ureijkf4u3iejrkmdsp:L>"?{":FSAFD+it';
// / ------------------------------

// / ------------------------------ 
// / Directory locations ...
  // / Install HR-AV to the following directory.
  // / DO NOT CHANGE THE DEFAULT INSTALL DIRECTORY!!! 
$InstLoc = str_replace(DIRECTORY_SEPARATOR.DIRECTORY_SEPARATOR, '', str_replace(DIRECTORY_SEPARATOR.'Scripts'.DIRECTORY_SEPARATOR.'PHP'.DIRECTORY_SEPARATOR.'PHP-AV', '', realpath(dirname(__FILE__))));
  // / The default location to scan if run with no input scan path argument. 
$ScanLoc = '';
  // / The absolute path where log files are stored.
$LogDir = $InstLoc.DIRECTORY_SEPARATOR.'Logs';
  // / The absolute path where report files are stored.
$ReportDir = $InstLoc.DIRECTORY_SEPARATOR.'Reports';
  // / The filename for the ScanCore virus definition file.
$DefsFileName = 'ScanCore_Virus.def';
  // / The filename for the ScanCore virus definition file.
$DefsDir = $InstLoc.DIRECTORY_SEPARATOR.'Definitions';
  // / The absolute path where virus definitions are found.
$DefsFile = $DefsDir.DIRECTORY_SEPARATOR.$DefsFileName;
// / ------------------------------ 

// / ------------------------------ 
// / General Information ...
  // / Number of bytes to store in each logfile before splitting to a new one.
$MaxLogSize = '1048576';
// / ------------------------------ 