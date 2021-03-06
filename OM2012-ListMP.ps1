#=====================================================================================================
# AUTHOR:	Dieter Wijckmans
# DATE:		03/08/2012
# Name:		list_mp2012.PS1
# Version:	1.0
# COMMENT:	Dump the installed management packs to a html file for backup reasons.
# CREDITS:  This script is based on the script of Kristopher bash
#           The original blog post can be found here: http://operatingquadrant.com/2009/08/19/scom-automating-management-pack-documentation/
# Usage:	.\list_mp2012.ps1 
#
# http://gallery.technet.microsoft.com/SCOM-2012-Create-a-html-3950d5bc
#=====================================================================================================

###Prepare environment for run###

####
# Start Ops Mgr snapin, get management server
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
#$locationroot = "C:\backup\SCOM\Autodocumentation\"
$locationroot = "$((Get-item $PSScriptRoot).PSDrive.Root)SCOMBackup\AutoDocumentation\"
if((test-path $locationroot) -eq $false) { mkdir $locationroot }
$locationfolder = $strbackupdate -Replace "/","-"
$location = $locationroot + $locationfolder
new-item "$location" -type directory -force
$outFile=$location + "\MPContents_" + "$locationfolder" + ".html"
write-host "Outfile: " + $outFile

#Delete backup location older than 15 days
#To make sure our disk will not be cluttered with old backups we'll keep 15 days of backup locations.
$Retentionperiod = "15"
$folders = dir $locationroot 
$now = [System.DateTime]::Now 
$old = $now.AddDays("-$Retentionperiod") 

foreach($folder in $folders) 
{ 
   if($folder.CreationTime -lt $old) { Remove-Item $folder.FullName -recurse } 
}


#Export all management packs to html file
$mps = Get-SCManagementPack |Sort-Object DisplayName

#Get all the detail in the html file

"<html>" | out-file $outFile
"<head>" | out-file -append $outFile
"<style>body{color: black;font-family: Verdana; font-size: 9pt}table{text-align:left;clear:both;margin-bottom:15px;width:800px;}tr{vertical-align:top;}"| out-file -append $outFile
"td{font-size:9pt;text-align:left;}.titleblock{	font-size: 11pt;font-weight:bold;color:darkblue;border-bottom: solid 1px #454545;}" | out-file -append $outFile
".sectiontitle{color:darkblue;font-weight:bolder;border-bottom:solid 1px #e2e2e2;}.leftcol{width:125px;font-weight:bold;}.rightcol{width:675px;}" | out-file -append $outFile
".objecttitle{width:675px;color:blue;}" | out-file -append $outFile
"</style></head>" | out-file -append $outFile
"<body>"  | out-file -append $outFile

function f_MPProperties([string]$s_mpDispname)
{
	$sHeader="<table>"
	$sHeader=$sHeader +"<tr><td class=titleblock colspan=3>MANAGEMENT PACK:  "+ $s_MPDispName  + "</td></tr>"
	$sHeader=$sHeader + "<tr><td>Version: " + $mp.Version + "</td>"
	$sHeader=$sHeader + "<td>Created: " + $mp.TimeCreated  + "</td>"
	$sHeader=$sHeader + "<td>Modified: " + $mp.LastModified  + "</td></tr>"
	$sHeader=$sHeader + "<tr><td colspan=3>" + $mp.Description + "</td></tr></table>"
	$sHeader |out-file -append $outfile
}

function f_MPGroups([string]$s_mpDispname)
{
	$sGroups="<table><tr><td colspan=2 class=sectiontitle>GROUPS</td></tr>" 
	foreach($grp in $mp.GetClasses())
	{
	  $sGroups=$sGroups + "<tr><td class=leftcol>Group:</td><td class=rightcol>" + $grp.DisplayName + "</td></tr>"
	  $sGroups=$sGroups + "<tr><td colspan=2>" + $grp.Description + "</td></tr>"
	  $sGroups=$sGroups + "<tr><td colspan=2>&nbsp;</td></tr>"
	}
	$sGroups=$sGroups + "</table>"
	$sGroups |out-file -append $outfile
}


function f_MPMon([string]$s_mpDispname)
{
	$sMon="<table><tr><td colspan=2 class=sectiontitle>MONITORS</td></tr>"
	foreach($mon in $mp.GetMonitors())
	{
		$sMon=$sMon + "<tr><td class=leftcol>Monitor:</td><td class=objecttitle>" + $mon.DisplayName + "</td></tr>"
		$sMon=$sMon + "<tr><td class=leftcol>Type:</td><td class=rightcol>" + $mon.XMLTag + "</td></tr>"
		$sMon=$sMon + "<tr><td class=leftcol>Category:</td><td class=rightcol>" + $mon.Category + "</td></tr>"
		$sMon=$sMon + "<tr><td class=leftcol>Modified:</td><td class=rightcol>" + $mon.LastModified + "</td></tr>"
		$sMon=$sMon + "<tr><td colspan=2>" + $mon.Description + "</td></tr>"
	  	$sMon=$sMon + "<tr><td colspan=2>&nbsp;</td></tr>"
	}
	$sMon=$sMon + "</table>"
	$sMon |out-file -append $outfile
}

function f_MPRules([string]$s_mpDispname)
{
	$sRule="<table><tr><td colspan=2 class=sectiontitle>RULES</td></tr>" 
	foreach($rule in $mp.GetRules())
	{
		$sRule=$sRule + "<tr><td class=leftcol>Rule:</td><td class=objecttitle>" + $rule.DisplayName + "</td></tr>"
		$sRule=$sRule + "<tr><td class=leftcol>Category:</td><td class=rightcol>" + $rule.Category + "</td></tr>"
		$sRule=$sRule + "<tr><td class=leftcol>Write Action:</td><td class=rightcol>" + $rule.WriteActionCollection + "</td></tr>"
		$sRule=$sRule + "<tr><td class=leftcol>Modified:</td><td class=rightcol>" + $rule.LastModified +"</td></tr>"
		$sRule=$sRule + "<tr><td colspan=2>" + $rule.Description +"</td></tr>"
		$sRule=$sRule + "<tr><td colspan=2>&nbsp;</td></tr>"
	}
	$sRule=$sRule + "</table>"
	$sRule |out-file -append $outfile
}

function f_MPViews([string]$s_mpDispname)
{
	$sView="<table><tr><td colspan=2 class=sectiontitle>VIEWS</td></tr>" 
	foreach($view in $mp.GetViews())
	{
		$sview=$sview + "<tr><td class=leftcol>View:</td><td class=objecttiel>" + $view.DisplayName + "</td></tr>"
		$sview=$sview + "<tr><td colspan=2>" + $view.Description +"</td></tr>"
		$sView=$sView + "<tr><td colspan=2>&nbsp;</td></tr>"
	}
	$sview=$sview + "</table>"
	$sview |out-file -append $outfile
}

function f_GetMP([string]$mpDisplayName)
{
	f_MPProperties($mpDisplayName)
	f_MPGroups($mpDisplayName)
	f_MPMon($mpDisplayName)
	f_MPRules($mpDisplayName)
	f_MPViews($mpDisplayName)
	"<p style=`"page-break-before:always;`"/>" |out-file -append $outfile
}

foreach($mp in $mps)
{
	$mpDisplayName=$mp.DisplayName
	f_GetMP($mpDisplayName)
}

"</body></html>"|out-file -append $outfile

 
#Error handling
$sMessageOK = 'SCOM: Documentation of the install mps was successful'
$sMessageNOK = 'SCOM: Documentation of the install mps failed'
Write-Host $error
if ($error.count -eq "0")
{
	$oLog = New-Object System.Diagnostics.EventLog 
    $oLog.Set_Log("Operations Manager") 
    $oLog.Set_Source("Health Service Script")
    $oLog.WriteEntry($sMessageOK,"Information",915) 
	
	
}
else
{
	$Error | Out-File $errorLogFile
	$oLog = New-Object System.Diagnostics.EventLog 
    $oLog.Set_Log("Operations Manager") 
    $oLog.Set_Source("Health Service Script")
    $oLog.WriteEntry($sMessageNOK,"Error",916) 
}

#mailing
#$Sender = "Fill in email here"


#OK
#$OKRecipient = "fill in email here"
#$strOKSubject = "Backup succesfully finished"
#$strBody = "The backup taken of the  was succesfully taken"


#NOK
#$ErrRecipient = "fill in email here"
#$strErrSubject = "FAILED to backup"
#$strErrBody = "Error occurred when excuting backup of unsealed management packs."

#if ($error.count -eq "0")
#{
#    send-mailmessage -from "$Sender" -to "$OKRecipient" -subject "$strOKSubject" -body "$StrBody" -smtpServer fill in smtp server here 
#}
#else
#{
#    send-mailmessage -from "$Sender" -to "$ErrRecipient" -subject "$strErrSubject" -body "$StrErrBody" -smtpServer fill in smtp server here
#}

###
# Remove the Ops Mgr PSSnapin
###

Remove-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client
