<#
    .NOTES

      Author:  Matt Lavery
      Date:    20/02/2014

      Copyright (c) blog.matticus.net.  All rights reserved.
      
      Legal Stuff: These contents and code is provided “as-is”. The information, opinions and views expressed are those of the author and do not necessarily state or 
      reflect those of any other company with affiliation to the products discussed. This includes but is not limited to any Code, URLs, and Tools. The author does 
      not accept any responsibility from its use, and strongly recommends adequate evaluation against your own requirements to measure suitability.

      Change History
      Version  Who          When           What
      --------------------------------------------------------------------------------------------------
      1.0      MLavery      20-Feb-2014    Initial Coding
      1.1      MLavery      26-Mar-2014    Code Optimisation (param binding, help, etc)
      1.2      MLavery      14-Oct-2014    Now uses $PSScriptRoot to reference the root drive

    .SYNOPSIS
      Exports the Monitors and Rules from a Management Pack either from a SCOM Management Server or from Disk

    .DESCRIPTION
      Exports the Monitors and Rules from a Management Pack either from a SCOM Management Server or from Disk

    .PARAMETER MP
      The Management Pack object of the MP to export.

    .PARAMETER MPName
      The Management Pack name to export. Can be either an Management Pack name or the file name depending on the value of the FromDisk parameter.

    .PARAMETER FromDisk
      A switch to control if the Management Pack is to be retrieved from Disk. When not supplied then the Management Pack is retrieved from a Management Server.

    .PARAMETER OutputPath
      Used to specify the output path to save the CSV files to. Note a sub directory with the Date will be added to the path.
      If no value is supplied then the default is to the path \SCOMBackup\MPExports\ at the root of the drive where the script is run from

    .EXAMPLE
      .\Export-MPtoCSV.ps1 -MPName Microsoft.SQLServer.2008

      Exports the Monitors and Rules from the Management Pack Microsoft.SQLServer.2008

    .EXAMPLE
      .\Export-MPtoCSV.ps1 -MPName c:\MPDownloads\Microsoft.Exchange.Server.2007.Monitoring.Hub.mp -FromDisk

      Exports the Monitors and Rules from the Management Pack file Microsoft.Exchange.Server.2007.Monitoring.Hub.mp

    .EXAMPLE
      "Microsoft.SQLServer.2012.AlwaysOn.Monitoring","Microsoft.SQLServer.2008.Mirroring.Monitoring" | .\Export-MPtoCSV.ps1

      Exports the two Management Packs passed along the pipeline

    .EXAMPLE
      Get-SCOMManagementPack -Name "*Windows*" | .\Export-MPtoCSV.ps1

      Exports all of the management packs installed into the SCOM environment with the name like *Windows*

    .EXAMPLE
      Get-SCOMManagementPack -Name "*SQL*" | .\Export-MPtoCSV.ps1 -OutputPath C:\temp\MPExports

      Exports all of the management packs installed into the SCOM environment with the name like *SQL*
      Also overrides the default output path to be C:\temp\MPExports
  
    .EXAMPLE
      Get-ChildItem "D:\MPDownloads\*.mp" | .\Export-MPtoCSV -FromDisk -OutputPath D:\MPExports\
  
      Exports all of the management packs downloaded to the D:\MPDownloads directory to the directory D:\MPExports\

    .LINK
      http://blog.matticus.net

    .OUTPUTS
        CSV files to disk in the path \SCOMBackup\MPExports\YYYY-MM-DD\ at the root of the drive where this script is run from (unless overridden by param)

#>
#Requires -Version 2.0
[CmdletBinding(DefaultParameterSetName = "MPName", SupportsShouldProcess = $false)]
param(
    [Parameter(ParameterSetName="MP", Mandatory=$true, ValueFromPipeline=$true, HelpMessage = "The Management Pack object to export the Management Pack config from.")]
    [Alias("MPO")]
    [Microsoft.EnterpriseManagement.Configuration.ManagementPack]
    [ValidateNotNullOrEmpty()]
    $MP,
    
    [Parameter(ParameterSetName="MPName", Mandatory=$true, ValueFromPipeline=$true, HelpMessage = "The Management Pack name or File Name to export the Management Pack config from.")]
    [Alias("MPN")]
    [String]
    [ValidateNotNullOrEmpty()]
    $MPName,

    [Parameter(ParameterSetName="MPName", Mandatory=$false, ValueFromPipeline=$false, HelpMessage = "When supplied the Management Pack config will come from disk of the path supplied to MPName")]
    [Alias("FD")]
    [Switch]
    $FromDisk,

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
        #$OutputPath = "$((Get-item (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path).PSDrive.Root)SCOMBackup\MPExports\";
        #new variable PowerShell3+
        $OutputPath = "$((Get-item $PSScriptRoot).PSDrive.Root)SCOMBackup\MPExports\";
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

    #Check if we are to get the details from disk
    if ($FromDisk -eq $true)
    {
        Write-Verbose "Obtaining MP '$MPName' from disk";
        $MpObject = Get-SCOMManagementPack -ManagementPackFile $MPName;
    }
    else
    {
        #Check which method to get the MP name
        if ($MPName.Length -gt 0)
        {
            Write-Verbose "Obtaining MP '$MPName'";
            $MpObject = Get-SCOMManagementPack -Name $MPName;
        }
        else
        {
            Write-Verbose "Using Pipeline MP Object '$($MP.Name)'";
            $MpObject = $MP;
        }
    }

    #Export the Rules from the Management Pack by passing the MP object to the Get-SCOMRule cmdlet
    Write-Verbose "Exporting Rules"
    Get-SCOMRule -ManagementPack $MpObject | Select-Object * | Export-Csv "$($SaveToPath)$($MpObject.Name).Rules.csv";

    #Export the Monitors from the Management Pack by passing the MP object to the Get-SCOMMonitor cmdlet
    Write-Verbose "Exporting Monitors"
    Get-SCOMMonitor -ManagementPack $MpObject | Select-Object * | Export-Csv "$($SaveToPath)$($MpObject.Name).Monitors.csv";

}

end
{
    #Let the user know we are done
    Write-Host "Export Completed" -ForegroundColor Green;
}

