# SAS-PowerShell examples

These PowerShell examples include:

* ReadSasDataset.ps1, ReadSasDataTables.ps1, ReadSasDataColumns.ps1 -- these use the SAS OLE DB Local Provider to open/read SAS data set (sas7bdat) files directly.  No SAS installation is needed.  The SAS OLE DB Local Provider is part of the [SAS Providers for OLE DB package](https://support.sas.com/downloads/browse.htm?fil=1&cat=64), available as a free download from support.sas.com.
* SASMetadataGetColumns.ps1 -- connects to a SAS Metadata environment and gathers details about registered tables and columns.
* SASWorkspaceExample.ps1 and SASWorkspaceDownloadFile.ps1 -- connects to a SAS Workspace to run a SAS program and download a file from the server to the local file system.
