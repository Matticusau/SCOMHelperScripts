The following is a list of the tools and scripts contained within this project:

# PowerShell Scripts
**_Backup-MP.ps1_**
Backs up supplied management packs to the file system. Works along the PowerShell Pipeline.
See Get-Help .\Export-MPtoCSV.ps1

**_Download-AllSCOMMPs.ps1_**
This script will download all the Management Packs from the portal to a designated directory.
Based off the script by http://blogs.technet.com/b/stefan_stranger/archive/2013/03/13/finding-management-packs-from-microsoft-download-website-using-powershell.aspx 

**_Export-MPtoCSV.ps1_**
This script exports the supplied Management Pack to CSV. The Management Pack can be from the SCOM MS or the File System. Works along the PowerShell Pipeline.
See Get-Help .\Export-MPtoCSV.ps1

**_Export-OpsMgrEventLog.ps1_**
Used while onsite to export key events from the OpsMgr Event Log when performing a health check of a system.
See Get-Help .\Export-OpsMgrEventLog.ps1

**_OM2012-BackupMP.ps1_**
Backs up the unsealed MPs. Based on http://gallery.technet.microsoft.com/SCOM2012-Backup-unsealed-fc6fb9b5

**_OM2012-ExportOverrides.ps1_**
Exports all Overrides to CSV
Based on http://gallery.technet.microsoft.com/SCOM-2012-Export-Overrides-dbfe9e27

**_OM2012-ListMP.ps1_**
Creates a HTML report of all Management Packs within the SCOM Management Server,
Based on http://gallery.technet.microsoft.com/SCOM-2012-Create-a-html-3950d5bc

# SQL Server Scripts
**_DWDataSetAggregationSettings.sql_**
This file contains the TSQL queries used to manage the Data Warehouse aggregation retention settings. For more information see http://blogs.technet.com/b/kevinholman/archive/2010/01/05/understanding-and-modifying-data-warehouse-retention-and-grooming.aspx
