####################################################################################
#                                                                                  #
# Xello Sims Export and Sync Script                                                #
# V1.0                                                                             #
# Author: Mathew Slatter (Matt.Slatter@marplehall.stockport.sch.uk)                #
#                                                                                  #
# Users CommandReporter.exe to export Sims Report and PSFTP.exe to FTP it to Xello #
# If you have a different MIS it should be pretty easy to adapt if you can export  #
# a CSV in the right format.                                                       #
####################################################################################


# Custom Values - Ensure you set these to specific values either relating to you organisation or required by Xello

$WorkingDirectory = "C:\Xello Sync"                      # This is root directory where these files are installed. 
$SimsDotDir = "C:\Program Files (x86)\SIMS\SIMS .net\"   # Sims.Net Install Dir (don't forget the space!)
$SimsUser = "Sims Username"                              # Sims Username Name - Must have "Third Party Reporting" + "Scheduled Reports" Group Permissions + something like "Class Teacher"
$SimsPassword = "Sims Password"                          # Sims Password
$SimsServer = "<hostname>\<SQL_instance_name>"           # SQL Sever Name (<hostname>\<SQL_instance_name>).  Can also be found in your $SimsDotDir\connect.ini file 
$SimsDatabase = "sims"                                   # Sims Database name (almost always "sims"). Can also be found in your $SimsDotDir\connect.ini file
$SimsReport  =  "Xello Student Export"                   # Sims Xello Export Report name.This won't need changinf if you've inported the included "Xello Student Export.RptDef" report definition
$OutFile = ".\FTP Upload Files\StudentExport.csv"        # Sims Export Filename
$ftpLocation = 'ftp.xello.co.uk'                         # Xello FTP Host Address (Probably won't change
$ftpPort = '22'                                          # Xello FTP Port, currently port 22, SSH
$ftpUser = 'Xello Username'                              # School Xello FTP Username, this is useually ther shool name  
$ftpPassword = 'Xellow Password'                         # School Xello FTP Password. This will have been provided by Xello.
$ftpGetCommandFile = '.\PSFTP\ftpGetCommands.txt'        # Temp file to store batch FTP command. There should be no need to change this but script required write permissions. 
$SimsExportFile = '.\FTP Upload Files\StudentExport.csv' # File that contains export data from MIS Report. 
$FtpUploadFile = ".\FTP Upload Files\Student.csv"        # Processed file for upload, should be called Student.csv. It inncludes required field not held in MIS (CurrentSchoolCode and PreRegSchoolCode) and corrects date format.
$PSFTP_Path = ".\PSFTP\psftp.exe"                        # PSFTP.exe file path
$LogFile = ".\LastRunLog.txt"                            # Log file location - This file will contain log of commands run by PSFTP.exe and any errors during the last run ONLY of this script. 
$SchoolCode = "SchoolCode"                               # School Code - Must match value contained in School.csv

# Note: This script does not include the optional columns StateProvNumber, Password or SSOStudentToken

# ---------------- Do not edit anything below this line ----------------

$CommandOutput=$ErrorOutput=$ErrorMessage = $null
$IsError = $false

try{
    
    Set-Location -Path $WorkingDirectory -ea Stop
    
    # Runs Sims Command Reporter to automate $SimsReport. Capture any errors to $ErrorMessage

    $ErrorMessage = & $SimsDotDir'CommandReporter.exe' /USER:$SimsUser /PASSWORD:$SimsPassword /SERVERNAME:$SimsServer /DATABASENAME:$SimsDatabase /REPORT:$SimsReport /OUTPUT:$OutFile | Select-String CommandReporterError
    
    #If there is no error message then continue
    if($ErrorMessage -eq $Null){

        # Write the batch FTP commands to temp file. This simplifies the running of PSFTP.exe below and ensure the values of $FtpUploadFile is respected on each run. 
        Set-Content -Path $ftpGetCommandFile -value ('put "' + $FtpUploadFile +'"')
        
        # Process $SimsExportFile to add the two required fields not held in MIS (CurrentSchoolCode and PreRegSchoolCode) and corrects date format. Outputs to $FtpUploadFile ready for FTP upload.
        Import-Csv $SimsExportFile | % {$_.DateOfBirth = ([datetime]($_.DateOfBirth)).ToString('dd-MM-yyyy');$_} | Select-Object "UPN ID","FirstName","LastName","Gender", "DateOfBirth","CurrentYear", @{Name="CurrentSchoolCode";Expression={"106138"}} , "PreRegSchoolCode","Username","Email" | Export-Csv $FtpUploadFile -NoTypeInformation
    
        #  Builds required PSFTP argument list
        $FTP_Command =  "-P $ftpPort -l $ftpUser -pw $ftpPassword $ftpLocation -b $ftpGetCommandFile -bc -batch"
    
        # Invokes PSFTP.exe with required arguments, capturing post the standard output and any errors (including some that are not errors!).
       $CommandOutput =  Invoke-Expression "$PSFTP_Path $FTP_Command" -ErrorVariable ErrorOutput 2>$null
   
       # We can't be sure if the command succeeded so we need to do some work to figure it out.

       # If there is no command output then there must be an error. Setting this here prevents an exception if we try and query it later.
       if($CommandOutput -eq $null){
            $IsError = $true
        }
    }else{
        # Else there was an error with Sims Export so skip processing and uploading the output and instead flag the error.
        $IsError = $true
    }
   
}catch [System.Management.Automation.ItemNotFoundException]{
    
    # Path not found when trying to Set-Location. Flag error and then set working path to directory containing this script so hopefully we can write $LogFile 
    
    $IsError = $true
    $ErrorMessage = $_.Exception.Message
    Set-Location -Path $PSScriptRoot

}catch{
    
    # Flag error and capture exception message to write in $LogFile

    $IsError = $true
    $ErrorMessage = $_.Exception.Message

}finally{
    
    # Output any PSFTP Commands to $LogFile. This will create file if it doesn't exist or otherwise overwrite it.
    # This might be useful even if an error has been flagged as it may contain info on the error.
    Set-Content -Path $LogFile -Value "Xello Sync run at: $((Get-Date).ToString())`r`n`r`n"
    Add-Content -Path $LogFile -Value $CommandOutput
    Add-Content -Path $LogFile -Value "`r`n----------------------------------------------------`r`n"

    #If we've not detected an error so far
    if(-not $IsError){
        #Check for Error opening Temp $ftpGetCommandFile. If present then flag Error and set $ErrorMessage 
        if(($CommandOutput | Select-String "Fatal: unable to open") -ne $null){
            $IsError = $true
            $ErrorMessage = "Cannot find path '" + $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ftpGetCommandFile) + "' because it does not exist." 
        }
    }

    #Check PSFTP $ErrorOutput for specific known error conditions. If any present then flag error and set appropriate $ErrorMessage
    if($ErrorOutput -ne $null){
        if(($ErrorOutput | Select-String "Access denied") -ne $null){
            $IsError = $true
            $ErrorMessage = "Accessed denied: Username or Password incorrect."
        }elseif(($ErrorOutput | Select-String "Host does not exist") -ne $null){
            $IsError = $true
            $ErrorMessage = "Host $ftpLocation does not exist or cannot be accessed."
        }elseif(($ErrorOutput | Select-String "Network error:") -ne $null){
            $IsError = $true
            $ErrorMessage = "Host $ftpLocation Connection timed out."
        }
    }
     
    # If we've found an error then append it to $LogFile
    if($IsError -eq $true){
        Add-Content -Path $LogFile -Value "Error: $ErrorMessage"
     }else{
        Add-Content -Path $LogFile -Value "Sync Successful"
     }
}

