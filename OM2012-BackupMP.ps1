#=====================================================================================================
# AUTHOR:	Dieter Wijckmans
# DATE:		05/01/2012
# Name:		Backup_mp2012.PS1
# Version:	1.0
# COMMENT:	Dump the Unsealed management packs to a backup location
# 
# Usage:	.\Backup_mp2012.ps1 
#
# http://gallery.technet.microsoft.com/SCOM2012-Backup-unsealed-fc6fb9b5
#=====================================================================================================

###Prepare environment for run###

####
# Start Ops Mgr snapin
###



##Read out the Management server name
$objCompSys = Get-WmiObject win32_computersystem
$inputScomMS = $objCompSys.name

#Initializing the Ops Mgr 2012 Powershell provider#
Import-Module -Name "OperationsManager" 
New-SCManagementGroupConnection -ComputerName $inputScomMS  


#Set Culture Info# In this case Dutch Belgium
$cultureInfo = [System.Globalization.CultureInfo]'EN-AU'

#Error handling setup
$error.clear()
$erroractionpreference = "SilentlyContinue"
$thisScript = $myInvocation.MyCommand.Path
$scriptRoot = Split-Path(Resolve-Path $thisScript)
$errorLogFile = Join-Path $scriptRoot "error.log"
if (Test-Path $errorLogFile) {Remove-Item $errorLogFile -Force}

#Define the backup location

#Get date
$Backupdatetemp = Get-Date
$Backupdatetemplocal = ($Backupdatetemp).tolocaltime()
$Backupdate = $Backupdatetemplocal.ToShortDateString()
$strBackupdate = $Backupdate.ToString()

#Define backup location
#$locationroot = "C:\backup\SCOM\unsealedMP\"
$locationroot = "$((Get-item $PSScriptRoot).PSDrive.Root)SCOMBackup\unsealedMP\"
if((test-path $locationroot) -eq $false) { mkdir $locationroot }
$locationfolder = $strbackupdate -Replace "/","-"
$location = $locationroot + $locationfolder
echo $location
new-item "$location" -type directory -force

#Delete backup location older than 15 days
#To make sure our disk will not be cluttered with old backups we'll keep 15 days of backup locations.
$Retentionperiod = "15"
$folders = dir $locationroot 
echo $folders
$now = [System.DateTime]::Now 
$old = $now.AddDays("-$Retentionperiod") 

foreach($folder in $folders) 
{ 
   if($folder.CreationTime -lt $old) { Remove-Item $folder.FullName -recurse } 
}


#Export all unsealed management packs
get-SCOMmanagementpack | where {$_.Sealed -eq $false} | Export-SCOMManagementPack -Path:$location 
 
 
#Error handling
$sMessageOK = 'SCOM: Unsealed management packs backup was successful'
$sMessageNOK = 'SCOM: Unsealed management packs backup was unsuccessful'
Write-Host $error
if ($error.count -eq "0")
{
	$oLog = New-Object System.Diagnostics.EventLog 
    $oLog.Set_Log("Operations Manager") 
    $oLog.Set_Source("Health Service Script")
    $oLog.WriteEntry($sMessageOK,"Information",910) 
	
	
}
else
{
	$Error | Out-File $errorLogFile
	$oLog = New-Object System.Diagnostics.EventLog 
    $oLog.Set_Log("Operations Manager") 
    $oLog.Set_Source("Health Service Script")
    $oLog.WriteEntry($sMessageNOK,"Error",911) 
}

#mailing
#$Sender = "fill in the sender here"


#OK
#$OKRecipient = "Fill in recipient for OK"
#$strOKSubject = "Backup succesfully finished"
#$strBody = "The backup taken of the  was succesfully taken"


#NOK
#$ErrRecipient = "Fill in recipient for NOK"
#$strErrSubject = "FAILED to backup"
#$strErrBody = "Error occurred when excuting backup of unsealed management packs."

#if ($error.count -eq "0")
#{
#    send-mailmessage -from "$Sender" -to "$OKRecipient" -subject "$strOKSubject" -body "$StrBody" -smtpServer Fill in your SMTP
#}
#else
#{
#    send-mailmessage -from "$Sender" -to "$ErrRecipient" -subject "$strErrSubject" -body "$StrErrBody" -smtpServer Fill in your SMTP
#}

###
# Remove the Ops Mgr PSSnapin
###

Remove-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client
