# -------------------------------------------------------------------
# ReadSasDataTables.ps1 
# Scan a folder and all subfolders for SAS7BDAT files
# and report on the SAS data set attributes
# Example usages:
#  
#   ReadSasDataTables.ps1 c:\Data 
#      - puts all output to the console
#
#   ReadSasDataTables.ps1 c:\Data | Out-GridView
#      - Opens a new output grid view with all data attributes displayed
#
#   ReadSasDataTables.ps1 c:\Data | Export-CSV -Path C:\Report\tables.csv -NoTypeInformation
#      - Creates a CSV file, ready for use in Excel, with all of the data attribute information
# -------------------------------------------------------------------
# check for an input file
if ($args.Count -eq 1) {
    $folderToProcess = $args[0] 
}
else {
    Write-Host "EXAMPLE Usage: ReadSasDataTables.ps1 path"
    Exit -1
}

# check that the input file exists
if (-not (Test-Path $folderToProcess)) {
    Write-Host "`"$folderToProcess`" does not exist."
    Exit -1
}

# check that the SAS Local Data Provider is present
if (-not (Test-Path "HKLM:\SOFTWARE\Classes\sas.LocalProvider")) {
    Write-Host "SAS OLE DB Local Data Provider is not installed.  Download from http://support.sas.com!"
    Exit -1
}

# Get all of the candidate SAS files
foreach ($dataset in Get-ChildItem $folderToProcess -Recurse -Filter "*.sas7bdat") {
    # build up a new object with the schema values
    $objectRecord = New-Object psobject

    $filePath = Split-Path $dataset.Fullname
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($dataset.FullName)  
    $criteria = @(0) * 4
    $criteria[2] = $filename
    
    $adSchemaTables = 20

    $objConnection = New-Object -comobject ADODB.Connection
    $objRecordset = New-Object -comobject ADODB.Recordset 
    
    try {
        $objConnection.Open("Provider=SAS.LocalProvider;Data Source=`"$filePath`";")
        $objRecordset = $objConnection.OpenSchema($adSchemaTables, $criteria)
        if ($objRecordset.EOF) {
            Write-Host "Cannot open " $dataset.Fullname
        }
	  
        # add file system properties for each record
        $objectRecord | add-member noteproperty `
            -name "FileName" `
            -value $dataset.Name;
		
        $objectRecord | add-member noteproperty `
            -name "Path" `
            -value $filePath;
		  
        $objectRecord | add-member noteproperty `
            -name "FileTime" `
            -value $dataset.LastWriteTime;
		 
        $objectRecord | add-member noteproperty `
            -name "FileSize" `
            -value $dataset.Length;    
			
        # Now read properties from "schema" internal to data set    
		
        # map property enumerations to friendly labels
        $propertyEnums = @(2, 5, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17)
        $propertyNames = @("TableName", "Label", "Created", "Modified", `
                "LogicalRecords", "PhysicalRecords", "RecordLength", `
                "Compressed", "Indexed", "Type", "Encoding", `
                "WindowsCodepage")
						   
        for ($i = 0; $i -lt $propertyEnums.Count; $i++) {
            # add static properties for each record
            $objectRecord | add-member noteproperty `
                -name $propertyNames[$i] `
                -value  $objRecordset.Fields.Item($propertyEnums[$i]).Value;
        }
		
        # emit the complete record as output
        $objectRecord
		
        $objRecordset.Close()
        $objConnection.Close()
    }
    catch [system.exception] {
        Write-Host "Unable to open " $dataset.Fullname " to read table properties"
    }
}

