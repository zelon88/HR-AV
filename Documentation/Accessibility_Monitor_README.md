# Accessibility-Tools-utilmon-Defender
A Windows 7-10 startup script for detecting and preventing "Ease Of Access" attacks.


This script was featured in the [how-to](https://www.honestrepair.net/index.php/category/howto/) blog post "[Windows Accessibility Tools… For Hackers Too???](https://www.honestrepair.net/index.php/2018/08/26/windows-accessibility-tools-for-hackers-too/)" on the [HonestRepair Blog](https://www.honestrepair.net/index.php/blog-posts/).

It is intended to be added to Group Policy Management on a domain or the Local Group Policy Editor on a standalone PC as a machine startup script. 

The script hashes cmd.exe (if it exists) and compares it against the hashes for each vulnerable tool in the Ease of Access center (utilmon.exe). A hard-coded hash exists as a default if cmd.exe was moved. 

*You must download "[Fake Sendmail For Windows](https://www.glob.com.au/sendmail/)" and extract all files to wherever you install the Accessibility_Defender.vbs script.*

If a compromise is detected the script will create a logfile of the incident and shut down the machine.
