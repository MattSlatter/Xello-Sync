# Xello-Sync

PowerShell Xello Automated Sync  
Version 1.0  
Author: Mathew Slatter (Matt.Slatter@marplehall.stockport.sch.uk)  

 A collection of scripts and programs to automate the process of exporting Student information (currently only from Sims MIS), processing it to include additional information not available in Sims, and then syncing this to Xello over FTP.

## Contents

**SIMS_XelloSync.ps1** - Powershell file containing Sync code  
**README** - This File  
**LICENCE** - GNU v3 Licence File  
**.\PSFTP\psftp.exe** - Copy of PSFTP.exe [https://www.chiark.greenend.org.uk/~sgtatham/putty/]  
**.\PSFTP\LICENCE.txt** - Licence file for PSFTP.exe  
**.\Scheduled Task - Example\Xello Sims Student Sync.xml** - Example Scheduled Task to run the Sync  
**.\Sims Report Definition** - Example\Xello Student Export.RptDef - Example Sims Report Definition to export the required data  
**\FTP Upload Files\\** - Directory that will contain exported / processed data  

## Prerequisites

	1. Sims must be installed on the devices that will be running the Sync. It can be a client or server (although I would recommend not the actual Sims Server)
	2. A Sims user account capable of running the export. It must have "Third Party Reporting" + "Scheduled Reports" Group Permissions + something like "Class Teacher" so it has permissions to the required student data.
	3. A School.csv file must have uploaded to Xello [https://help.xello.world/article/944-create-student-data-files]

## Installation

Extract the file structure from the Zip file to a chosen folder, keeping note of the root folder of this path.  

	1. Import the Report Definition file ".\Sims Report Definition - Example\Xello Student Export.RptDef" into your Sims installation.
	2. Import the scheduled task "Xello Sims Student Sync.xml" to Windows Task Scheduler on the device that will run the Sync. Edit as required, including specifying a user for it to run as.
	3. Edit the "Custom Values" section of the file SIMS_XelloSync.ps1. Ensure that all the variables are set to specific values either relating to your organisation or required by Xello. Note: $WorkingDirectory must be set to the value of the root installation directory path above and $SchoolCode must be the same value set in School.csv file you have already uploaded (see prerequisites above).

## FAQ

**Can I use my own Sims Report instead of the included example?**

As long as is exports the required information, yes. Ensure the column names match those required in Students.csv and then set the variable $SimsReport to the name of your report.

**If use a different MIS (e.g. CMIS) can I use this script?**

Currently the script only supports Sims however if there is a method for exporting data from your MIS then it should be fairly trivial to add compatibility for other MIS's. Depending on the availible output of this process you may need to add some error checking.  If you need any support with this then please get in touch.
If you only need to process and upload am existing CSV file the just comment out ("#") the line that invokes CommandReporter.exe. 

**I keep getting "Error: Accessed denied: Username or Password incorrect".**

If you're sure that you've entered the supplied details correctly, trying running PSFPT from the command line. Simply run "psftp.exe -P 22 -l USERNAME -pw PASSWORD ftp.xello.co.uk" to check you login credentials. If this works feel free to get in touch however if you get an "Access denied" error message you'll need to contact Xello and ask them to issue you with the correct details.
