<?php
// / -----------------------------------------------------------------------------------
// / APPLICATION INFORMATION ...
// / HR-AV, Copyright on 2/21/2016 by Justin Grimes, www.github.com/zelon88
// / 
// / LICENSE INFORMATION ...
// / This project is protected by the GNU GPLv3 Open-Source license.
// / 
// / APPLICATION DESCRIPTION ...
// / This application is designed to provide a web-interface for scanning files 
// / for viruses on a server for users of any web browser without authentication. 
// / 
// / HARDWARE REQUIREMENTS ... 
// / This application requires at least a Raspberry Pi Model B+ or greater.
// / This application will run on just about any x86 or x64 computer.
// / 
// / DEPENDENCY REQUIREMENTS ... 
// / This application requires Debian Linux (w/3rd Party audio license), 
// / Apache 2.4, PHP 7.0+, JScript, WordPress & mySql (optional) & ClamAV.
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// / The following code will load required HR-AV files.
if (!file_exists(str_replace(DIRECTORY_SEPARATOR.DIRECTORY_SEPARATOR, '', str_replace(DIRECTORY_SEPARATOR.'Scripts'.DIRECTORY_SEPARATOR.'PHP'.DIRECTORY_SEPARATOR.'PHP-AV'.DIRECTORY_SEPARATOR '', realpath(dirname(__FILE__))).DIRECTORY_SEPARATOR.'Definitions'.DIRECTORY_SEPARATOR.'ScanCore_Config.php'))); die ('ERROR!!! ScanCore0, Cannot process the HR-AV ScanCore Configuration file (config.php)!'.PHP_EOL); 
else require_once (str_replace(DIRECTORY_SEPARATOR.DIRECTORY_SEPARATOR, '', str_replace(DIRECTORY_SEPARATOR.'Scripts'.DIRECTORY_SEPARATOR.'PHP'.DIRECTORY_SEPARATOR.'PHP-AV'.DIRECTORY_SEPARATOR, '', realpath(dirname(__FILE__))).DIRECTORY_SEPARATOR.'Definitions'.DIRECTORY_SEPARATOR.'ScanCore_Config.php'));
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// / The following code sets the global variables for the session.
  // / Application related variables.
  $Versions = 'PHP-AV App v4.0 | Virus Definition v4.9, 4/10/2019';
  $encType = 'ripemd160';
  $defaultMemoryLimit = 4000000;
  $defaultChunkSize = 1000000; 
  $report = '';
  $filecount = $infected = $dircount = 0;
  $CONFIG = $CONFIG['extensions'] = Array();
  $abort = $CONFIG['debug'] = FALSE;
  // / Time related variables.
  $Date = date("m_d_y");
  $Time = date("F j, Y, g:i a"); 
  // / SesHash related variables for developing predictable paths.
  $RandomNumber = rand(10000, 1000000).rand(10000,1000000)
  $SesHash = substr(hash($encType, $Date.$Salts1.$Salts2.$Salts3.$Salts4.$Salts5.$Salts6), - 12);
  $SesHash2 = substr(hash($encType, $RandomNumber.$SesHash.$Date.$Time.$Salts1.$Salts2.$Salts3.$Salts4.$Salts5.$Salts6), - 12);
  $SesHash3 = $SesHash.DIRECTORY_SEPARATOR.$SesHash2;
  // / Directory related variables.
  $ReportSubSubDir = $ReportDir.DIRECTORY_SEPARATOR.$SesHash
  $ReportFile = $ReportDir.DIRECTORY_SEPARATOR.$SesHash3
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// / A function to create a logfile if one does not exist.
function verifyLogFile($LogFile, $MaxLogSize) { 
  $LogInc = 0;
  if (!is_dir($LogDir)) die('ERROR!!! ScanCore1, The specified $LogDir does not exist at '.$LogDir.' on '.$Time.'.');
  while (file_exists($LogFile) && round((filesize($LogFile) / $MaxLogSize), 2) > $MaxLogSize) { 
    $LogInc++; 
    $LogFile = $LogDir.DIRECTORY_SEPARATOR.'HR-AV_ScanCore_'.$LogInc.'.txt.'; 
    $MAKELogFile = file_put_contents($LogFile, 'OP-Act: Logfile created on '.$Time.'.'.PHP_EOL, FILE_APPEND); }
  if (!file_exists($LogFile)) $MAKELogFile = file_put_contents($LogFile, 'OP-Act: Logfile created on '.$Time.'.'.PHP_EOL, FILE_APPEND); }
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// / A function to add an entry to the logs.
function addLogEntry($entry, $error, $errorNumber) {
  if (!is_numeric($errorNumber)) $errorNumber = 0;
  if ($error === TRUE) $preText = 'ERROR!!! '.$Time.' ScanCore'.$errorNumber.', ';
  else $preText = 'OP-Act: ';
  return(file_put_contents($LogFile, $preText.$entry)); } 
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// / A function to parse supplied command-line arguments.
function parseArgs($argv) { 
  global $defaultMemoryLimit, $defaultChunkSize
  foreach $argv as $key=>$argv {
    $arg = htmlentities(str_replace(str_split('~#[](){};:$!#^&%@>*<"\''), '', $arg));
    if strpos(lcase($arg), '-memoryLimit') !== FALSE $memoryLimit = $argv[$key + 1];
    if strpos(lcase($arg), '-chunksize') !== FALSE $chunkSize = $argv[$key + 1]; }
  if (!file_exists($argv[1])) addLogEntry('The specified file was not found! The first argument must be a valid file or directory path!', TRUE, 200);
  if (!is_numeric($memoryLimit) or !is_numeric($chunkSize)) { 
    addLogEntry('Either the chunkSize argument or the memoryLimit argument is invalid. Substituting default values.', TRUE, 300); 
    $memoryLimit = $defaultMemoryLimit; 
    $chunkSize = $defaultChunkSize; }
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// Hunts files/folders recursively for scannable items.
function file_scan($folder, $defs, $ReportFile, $debug) {
  global $report, $memoryLimit;
  $dircount = 0;
  if ($d = @dir($folder)) {
    while (false !== ($entry = $d->read())) {
      $isdir = @is_dir($folder.'/'.$entry);
      if (!$isdir and $entry != '.' and $entry != '..') {      
        virus_check($folder.'/'.$entry, $defs, $debug, $defData); } 
      elseif ($isdir and $entry != '.' and $entry != '..') {
        $txt = 'OP-Act: Scanning folder '.$folder.' ... ';
        $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND);        
        $dircount++;
        file_scan($folder.'/'.$entry, $defs, $debug, $defData); } }
    $d->close(); } }
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// Reads tab-delimited defs file.
function load_defs($file, $debug) {
  global $ReportFile;
  $defs = file($file);
  $counter = 0;
  $counttop = sizeof($defs);
  while ($counter < $counttop) {
    $defs[$counter] = explode('  ', $defs[$counter]);
    $counter++; }
  $txt = 'OP-Act: Loaded '.sizeof($defs).' virus definitions.';
  $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND);
  return $defs; }
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// Check for >755 perms on virus defs.
function check_defs($file) {
  clearstatcache();
  $perms = substr(decoct(fileperms($file)), - 2);
  if ($perms > 55) return false;
  else return true; }
// / -----------------------------------------------------------------------------------

// / -----------------------------------------------------------------------------------
// Hashes and checks files/folders for viruses against static virus defs.
function virus_check($file, $defs, $defData, $ReportFile, $debug) {
  global $memoryLimit, $chunkSize, $report, $CONFIG;
  $infected = $filecount = 0;
  $filecount++;
  if ($file !== $DefsFileName) {
    if (file_exists($file)) { 
      $filesize = filesize($file);
      $data1 = hash_file('md5', $file);
      $data2 = hash_file('sha256', $file);
      // / Scan files larger than the memory limit by breaking them into chunks.
      if ($filesize >= $memoryLimit && file_exists($file)) { 
        $txt = 'OP-Act: Chunking file ... ';
        $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND);
        $handle = @fopen($file, "r");
        if ($handle) {
          while (($buffer = fgets($handle, $chunkSize)) !== false) {
            $data = $buffer; 
            if ($debug) { 
              $txt = 'OP-Act: Scanning chunk ... ';
              $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND); }
            foreach ($defs as $virus) {
              $virus = explode("\t", $virus[0]);
              if ($virus[1] !== '' && $virus[1] !== ' ') {
                if (strpos(strtolower($data), strtolower($virus[1])) !== FALSE or strpos(strtolower($file), strtolower($virus[1])) !== FALSE) { 
                  // File matches virus defs.
                  $txt = 'Infected: '.$file.' ('.$virus[0].', Data Match: '.$virus[1].')';
                  $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
                  $infected++; } } } }
          if (!feof($handle)) {
            $txt = 'ERROR!!! PHPAV160, Unable to open '.$file.' on '.$Time.'!';
            $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND); }
          fclose($handle); } 
          if ($virus[2] !== '' && $virus[2] !== ' ') {
            if (strpos(strtolower($data1), strtolower($virus[2])) !== FALSE) {
              // File matches virus defs.
              $txt = 'Infected: '.$file.' ('.$virus[0].', MD5 Hash Match: '.$virus[2].')';
              $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
              $infected++; } }
            if ($virus[3] !== '' && $virus[3] !== ' ') {
              if (strpos(strtolower($data2), strtolower($virus[3])) !== FALSE) {
                // File matches virus defs.
                $txt = 'Infected: '.$file.' ('.$virus[0].', SHA256 Hash Match: '.$virus[3].')';
                $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
                $infected++; } } 
            if ($virus[4] !== '' && $virus[4] !== ' ') {
              if (strpos(strtolower($data3), strtolower($virus[4])) !== FALSE) {
                // File matches virus defs.
                $txt = 'Infected: '.$file.' ('.$virus[0].', SHA1 Hash Match: '.$virus[4].')';
                $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND);
                $infected++; } } } } }
      // / Scan files smaller than the memory limit by fitting the entire file into memory.
      if ($filesize < $memoryLimit && file_exists($file)) {
        $data = file_get_contents($file); }
      if ($defData !== $data2) {
        foreach ($defs as $virus) {
          $virus = explode("\t", $virus[0]);
          if ($virus[1] !== '' && $virus[1] !== ' ') {
            if (strpos(strtolower($data), strtolower($virus[1])) !== FALSE or strpos(strtolower($file), strtolower($virus[1])) !== FALSE) {
             // File matches virus defs.
              $txt = 'Infected: '.$file.' ('.$virus[0].', Data Match: '.$virus[1].')';
              $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
              $infected++; } }
          if ($virus[2] !== '' && $virus[2] !== ' ') {
            if (strpos(strtolower($data1), strtolower($virus[2])) !== FALSE) {
                // File matches virus defs.
              $txt = 'Infected: '.$file.' ('.$virus[0].', MD5 Hash Match: '.$virus[2].')';
              $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
              $infected++; } }
            if ($virus[3] !== '' && $virus[3] !== ' ') {
              if (strpos(strtolower($data2), strtolower($virus[3])) !== FALSE) {
                // File matches virus defs.
                $txt = 'Infected: '.$file.' ('.$virus[0].', SHA256 Hash Match: '.$virus[3].')';
                $MAKELogFile = file_put_contents($ReportFile, 'OP-Act: '.$txt.PHP_EOL, FILE_APPEND);
                $infected++; } } 
            if ($virus[4] !== '' && $virus[4] !== ' ') {
              if (strpos(strtolower($data3), strtolower($virus[4])) !== FALSE) {
                // File matches virus defs.
                $txt = 'Infected: '.$file.' ('.$virus[0].', SHA1 Hash Match: '.$virus[4].')';
                $MAKELogFile = file_put_contents($ReportFile, $txt.PHP_EOL, FILE_APPEND);
                $infected++; } } } } 
  return $infected; }
// / -----------------------------------------------------------------------------------