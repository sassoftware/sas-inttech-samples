# -------------------------------------------------------------------
# ReadSasDataset.ps1 
# Read a SAS data set file (SAS7BDAT)
# and create a report with all records
#
# Example usages:
#  
#   ReadSasDataset.ps1 c:\Data\sample.sas7bdat 
#      - puts all output to the console
#
#   ReadSasDataset.ps1 c:\Data\sample.sas7bdat | Out-GridView
#      - Opens a new output grid view with all data records displayed
#
#   ReadSasDataset.ps1 c:\Data\sample.sas7bdat | Export-CSV -Path C:\Report\sample.csv -NoTypeInformation
#      - Creates a CSV file, ready for use in Excel, with all data records
# -------------------------------------------------------------------
# check for an input file
if ($args.Count -eq 1) {
    $fileToProcess = $args[0] 
}
else {
    Write-Host "EXAMPLE Usage: ReadSasDataset.ps1 path-and-name-of-SAS-dataset"
    Exit -1
}

# check that the input file exists
if (-not (Test-Path $fileToProcess)) {
    Write-Host "$fileToProcess does not exist."
    Exit -1
}

# check that the SAS Local Data Provider is present
if (-not (Test-Path "HKLM:\SOFTWARE\Classes\sas.LocalProvider")) {
    Write-Host "SAS OLE DB Local Data Provider is not installed."
    Exit -1
}

# split the file into Path and root name
$fileItem = Get-Item $fileToProcess
$filePath = Split-Path $fileItem.FullName -Parent
$filename = [System.IO.Path]::GetFileNameWithoutExtension($fileItem.FullName)

# constants for cursor behavior
$adOpenDynamic = 2
$adLockOptimistic = 3
$adCmdTableDirect = 512

$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset

try {
    $objConnection.Open("Provider=SAS.LocalProvider;Data Source=`"$filePath`";")
    $objRecordset.ActiveConnection = $objConnection
    $objRecordset.Properties.Item("SAS Formats").Value = "_ALL_"

    # open the data set
    # IMPORTANT: passing in a "missing" value for the connection
    # because the connection is already on the RecordSet object
    $objRecordset.Open($filename, [Type]::Missing,
        $adOpenDynamic, `
            $adLockOptimistic, `
            $adCmdTableDirect) 
	   
    $objRecordset.MoveFirst()

    # read all of the records within the SAS data file
    do {
        # build up a new object with the field values
        $objectRecord = New-Object psobject
		
        for ($i = 0; $i -lt $objRecordset.Fields.Count; $i++) {
            # add static properties for each record
            $objectRecord | add-member noteproperty `
                -name $objRecordset.Fields.Item($i).Name `
                -value  $objRecordset.Fields.Item($i).Value;
			 
        }
        # emit the object as output from this script
        $objectRecord
		
        # move on to the next record
        $objRecordset.MoveNext()
    } 
    until ($objRecordset.EOF -eq $True)

    # close all of the connections
    $objRecordset.Close()
    $objConnection.Close()
}

catch {
    Write-Host "Unable to process " $fileToProcess
    if (($null -ne $objConnection) -and ($objConnection.Errors.Count -gt 0)) {
        foreach ($adoError in $objConnection.Errors) {
            Write-Host $adoError.Description
        }
    }
    else {
        Write-Host $_.Exception.ToString()
    }
}
finally {
    if (($null -ne $objRecordset) -and ($objRecordset.State -ne 0)) {
        $objRecordset.Close()
    }
    if (($null -ne $objConnection) -and ($objConnection.State -ne 0)) {
        $objConnection.Close()
    }
}
