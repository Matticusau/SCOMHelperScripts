<#
    .NOTES

      Author:  Matt Lavery
      Date:    10/03/2014

      Copyright (c) blog.matticus.net.  All rights reserved.
      
      Legal Stuff: These contents and code is provided “as-is”. The information, opinions and views expressed are those of the author and do not necessarily state or 
      reflect those of any other company with affiliation to the products discussed. This includes but is not limited to any Code, URLs, and Tools. The author does 
      not accept any responsibility from its use, and strongly recommends adequate evaluation against your own requirements to measure suitability.

      Change History
      Version  Who          When           What
      --------------------------------------------------------------------------------------------------
      1.0      MLavery      10-Mar-2014    Initial Coding
      1.1      MLavery      14-Oct-2014    Now uses $PSScriptRoot to reference the root drive

    .SYNOPSIS
      Backs up an unsealed Management Pack to Disk

    .DESCRIPTION
      Backs up an unsealed Management Pack to Disk

    .PARAMETER MP
      The Management Pack Object of the MP to backup.

    .PARAMETER MPName
      The Management Pack name to backup.

    .PARAMETER OutputPath
      Used to specify the output path to save the MP Backups to. Note a sub directory with the Date will be added to the path.
      If no value is supplied then the default is to the path \SCOMBackup\MPBackups\ at the root of the drive where the script is run from

    .EXAMPLE
      .\Backup-MP.ps1 -MPName Microsoft.SQLServer.2008

      Backs up the Management Pack Microsoft.SQLServer.2008

    .EXAMPLE
      .\Backup-MP.ps1 -MPName "ALL"

      Backs up ALL the unsealed Management Packs

    .EXAMPLE
      "Microsoft.SQLServer.2012.AlwaysOn.Monitoring","Microsoft.SQLServer.2008.Mirroring.Monitoring" | .\Backup-MP.ps1

      Backs up the two Management Packs passed along the pipeline

    .EXAMPLE
      Get-SCOMManagementPack -Name "*Windows*" | .\Backup-MP.ps1

      Backs up all of the management packs installed into the SCOM environment with the name like *Windows*

    .EXAMPLE
      Get-SCOMManagementPack -Name "*SQL*" | .\Backup-MP.ps1 -OutputPath C:\temp\MPBackups\}

      Backs up all of the management packs installed into the SCOM environment with the name like *SQL*
      Also overrides the default output path to be C:\temp\MPBackups
  
    .LINK
      http://blog.matticus.net

    .OUTPUTS
        CSV files to disk in the path \SCOMBackup\MPBackups\YYYY-MM-DD\ at the root of the drive where this script is run from (unless overridden by param)

#>
#Requires -Version 2.0
[CmdletBinding(DefaultParameterSetName = "MPName", SupportsShouldProcess = $false)]
param(
    [Parameter(ParameterSetName="MP", Mandatory=$true, ValueFromPipeline=$true, HelpMessage = "The Management Pack object to backup the Management Pack config from.")]
    [Alias("MPO")]
    [Microsoft.EnterpriseManagement.Configuration.ManagementPack]
    [ValidateNotNullOrEmpty()]
    $MP,
    
    [Parameter(ParameterSetName="MPName", Mandatory=$true, ValueFromPipeline=$true, HelpMessage = "The Management Pack name to backup the Management Pack config from.", Position=1)]
    [Alias("MPN")]
    [String]
    [ValidateNotNullOrEmpty()]
    $MPName,

    [Parameter(ParameterSetName="", Mandatory=$false, ValueFromPipeline=$false, HelpMessage = "Overrides the default output path")]
    [Alias("Path")]
    [String]
    $OutputPath
)

    
begin
{
    #Check if we were provided an output path
    if (!($OutputPath.Length -gt 0))
    {
        #$OutputPath = "$((Get-item (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path).PSDrive.Root)SCOMBackup\MPBackups\";
        #new variable PowerShell3+
        $OutputPath = "$((Get-item $PSScriptRoot).PSDrive.Root)SCOMBackup\MPBackups\";
    }
    
    #add the current date to the output folder
    $OutputPath = Join-Path -Path $OutputPath -ChildPath "$(Get-Date -Format {yyyy-MM-dd})\";

    #make sure the path exists
    if ((Test-Path $OutputPath) -eq $false)
    {
	    $NewDir = New-Item -Path $OutputPath -Type Directory -Force;
    }
    
    #Set the path we will use within the script
    $SaveToPath = $OutputPath;
	
}

process
{

    #Check which method to get the MP name
    if ($MPName.Length -gt 0)
    {
        #Check if we are backing up everything
        if ($MPName.ToUpper() -eq "ALL")
        {
            Write-Verbose "Obtaining ALL MPs";
            $MpObject = Get-SCOMManagementPack | Where-Object {$_.Sealed -eq $false};
        }
        else
        {
            Write-Verbose "Obtaining MP '$MPName'";
            $MpObject = Get-SCOMManagementPack -Name $MPName;
        }
    }
    else
    {
        Write-Verbose "Using Pipeline MP Object '$($MP.Name)'";
        $MpObject = $MP;
    }

    #backup the MP using the Export-SCOMManagementPack CmdLet
    Write-Verbose "Backing up the MP to disk"
    $MpObject | Export-SCOMManagementPack -Path "$($SaveToPath)"

}

end
{
    #Let the user know we are done
    Write-Host "Backup Completed" -ForegroundColor Green;
}

