<#
    .NOTES

      Author:  Matt Lavery
      Date:    1/06/2014

      Copyright (c) blog.matticus.net.  All rights reserved.
      
      Legal Stuff: These contents and code is provided “as-is”. The information, opinions and views expressed are those of the author and do not necessarily state or 
      reflect those of any other company with affiliation to the products discussed. This includes but is not limited to any Code, URLs, and Tools. The author does 
      not accept any responsibility from its use, and strongly recommends adequate evaluation against your own requirements to measure suitability.

      Change History
      Version  Who          When           What
      --------------------------------------------------------------------------------------------------
      1.0      MLavery      01-Jun-2014    Initial Coding
      1.1      MLavery      25-Jun-2014    Added additional events to export
      
    .SYNOPSIS
      Exports and Evaluates the Event Log and produces a report

    .DESCRIPTION
      Exports and Evaluates the event log of a SCOM Management Server and produces a report of data

    .PARAMETER ManagementServer
      The Management server to evaluate

    .PARAMETER Path
      The file path of the report to create

    .EXAMPLE
      .\Export-OpsMgrEventLog.ps1

      Exports and Evaluates the Event Log of the local server and saves a report to the SCOMReports folder off the root drive of where the script is run from

    .EXAMPLE
      .\Export-OpsMgrEventLog.ps1 -ManagementServer bne-SCOM-01

      Exports and Evaluates the Event Log of the management server bne-SCOM-01 and saves a report to the SCOMReports folder off the root drive of where the script is run from

    .EXAMPLE
      .\Export-OpsMgrEventLog.ps1 -ManagementServer bne-SCOM-01 -Path c:\Reports

      Exports and Evaluates the Event Log of the management server bne-SCOM-01 and saves a report to c:\Reports folder

    .LINK
      http://blog.matticus.net

    .OUTPUTS
        CSV and HTML files to disk in either the provided path or the SCOMReports folder off the root drive of where the script is run from with a sub-folder \YYYY-MM-DD\

#>
#Requires -Version 2.0
[CmdletBinding(DefaultParameterSetName = "SCOM", SupportsShouldProcess = $false)]
param(
    [Parameter(ParameterSetName="SCOM", Mandatory=$false, ValueFromPipeline=$true, HelpMessage = "The Management Server to evaluate.")]
    [Alias("MS")]
    [String]
    [ValidateNotNullOrEmpty()]
    $ManagementServer,
    
    [Parameter(ParameterSetName="SCOM", Mandatory=$false, ValueFromPipeline=$false, HelpMessage = "The path to save the reports to")]
    [Alias("ReportPath")]
    [String]
    $Path
)

    
begin
{
    #Check if we were provided an output path
    if (!($Path.Length -gt 0))
    {
        $Path = "$((Get-item (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path).PSDrive.Root)SCOMReports\";
    }

    #add the current date to the output folder
    $Path = Join-Path -Path $Path -ChildPath "$(Get-Date -Format {yyyy-MM-dd})\";

    #make sure the path exists
    if ((Test-Path $Path) -eq $false)
    {
	    $NewDir = New-Item -Path $Path -Type Directory -Force;
    }
    
    #Set the path we will use within the script
    $SaveToPath = $Path;
	
}

process
{
    
    if ($ManagementServer.Length -gt 0) {$ServerName = $ManagementServer}
    else {$ServerName = $env:COMPUTERNAME} 

    #Check we can connect to the server
    if ((Test-Connection -ComputerName $ServerName -Count 1 -Quiet) -eq $false)
    {
        Throw "Unable to connect to server $ServerName";
        break;
    }


    #Event 2115
    Write-Verbose "Exporting Event 2115 data";
    $CSVFile = Join-Path -Path $SaveToPath -ChildPath "$($ServerName)_EventLog_Event2115_$(Get-Date -Format {yyyyMMddhhmmss}).csv";
    Get-EventLog -ComputerName $ServerName -LogName "Operations Manager" -Source "HealthService" | Where -Property EventId -EQ -Value 2115 | Export-CSV -Path $CSVFile;
    
    #Custom columns to filter contents of CSV file
    $ColYear = @{
        Name="Year"; 
        Expression={
            (Get-Date -Date $_.TimeGenerated).Year
        };
    };
    $ColMonth = @{
        Name="Month"; 
        Expression={
            (Get-Date -Date $_.TimeGenerated).Month
        };
    };
    $ColDay = @{
        Name="Day"; 
        Expression={
            (Get-Date -Date $_.TimeGenerated).Day
        };
    };
    $ColHour = @{
        Name="Hour"; 
        Expression={
            (Get-Date -Date $_.TimeGenerated).Hour
        };
    };
    $ColMinute = @{
        Name="Minute"; 
        Expression={
            (Get-Date -Date $_.TimeGenerated).Minute
        };
    };
    $ColManagementGroup = @{
        Name="ManagementGroup"; 
        Expression={
            (
                (
                    $_.Message -replace "^(A Bind Data Source in Management Group\s)",""
                ) -replace "(\shas posted items to the workflow, but has not received a response in\s).*(\s*seconds.\s*This indicates a performance or functional problem with the workflow.)*(\r|\n|\s|.)*",""
            )
        };
    };
    $ColSeconds = @{
        Name="Seconds"; 
        Expression={
            (
                (
                    $_.Message -replace "^(A Bind Data Source in Management Group\s).*(\shas posted items to the workflow, but has not received a response in\s)",""
                ) -replace "\s*seconds.\s*This indicates a performance or functional problem with the workflow.(\r|\n|\s|.)*",""
            )
        };
    };
    #Filter the contents of the Event 2115 results
    Write-Verbose "Filtering Event 2115 CSV data";
    Import-Csv $CSVFile | Select TimeGenerated,$ColYear,$ColMonth,$ColDay,$ColHour,$ColMinute,$ColManagementGroup,$ColSeconds | Export-Csv $CSVFile.Replace(".csv","_filtered.csv");

    #Create a summary of the Event 2115 results
    Write-Verbose "Summarizing Event 2115 CSV data";
    Import-Csv $CSVFile | Select-Object $ColYear,$ColMonth,$ColDay,$ColHour | Group-Object Year,Month,Day,Hour | `
    Select-Object @{Name="Year";Expression={$_.Values[0]};},@{Name="Month";Expression={$_.Values[1]};},@{Name="Day";Expression={$_.Values[2]};},@{Name="Hour";Expression={$_.Values[3]};},Count | `
    Export-CSV -Path $CSVFile.Replace(".csv","_summary.csv") -NoTypeInformation;

    #Event 31552
    Write-Verbose "Exporting Event 31552 data";
    $CSVFile = Join-Path -Path $SaveToPath -ChildPath "$($ServerName)_EventLog_Event31552_$(Get-Date -Format {yyyyMMddhhmmss}).csv";
    Get-EventLog -ComputerName $ServerName -LogName "Operations Manager" -Source "Health Service Modules" | Where -Property EventId -EQ -Value 31552 | Export-CSV -Path $CSVFile;

    #OpsMgr Management Configuration events
    Write-Verbose "Exporting OpsMgr Management Configuration events";
    $CSVFile = Join-Path -Path $SaveToPath -ChildPath "$($ServerName)_EventLog_OpsMgrMgmtConfig_$(Get-Date -Format {yyyyMMddhhmmss}).csv";
    Get-EventLog -ComputerName $ServerName -LogName "Operations Manager" -Source "OpsMgr Management Configuration" | Export-CSV -Path $CSVFile;

    #Error and Warning events in last 48 hours
    Write-Verbose "Exporting Error/Warning events in last 48 hours";
    $CSVFile = Join-Path -Path $SaveToPath -ChildPath "$($ServerName)_EventLog_ErrorWarning48Hours_$(Get-Date -Format {yyyyMMddhhmmss}).csv";
    Get-EventLog -ComputerName $ServerName -LogName "Operations Manager" -EntryType Error,Warning -After ((Get-Date).AddHours(-48)) | Export-CSV -Path $CSVFile;


}

end
{
    #Let the user know we are done
    Write-Host "Export Completed" -ForegroundColor Green;
}

