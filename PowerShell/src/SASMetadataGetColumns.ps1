# SasMetadataGetColumns.ps1
# Example usage:
# For table grid display, use:
#    .\SasMetadataGetColumns.ps1 | Out-Gridview
# For export to CSV
#    .\SasMetadataGetColumns.ps1 | Export-Csv -Path "c:\output\cols.csv" -NoTypeInformation
# -------------------------------------------------------------------
# create the Integration Technologies objects
$objFactory = New-Object -ComObject SASObjectManager.ObjectFactoryMulti2
$objServerDef = New-Object -ComObject SASObjectManager.ServerDef

# assign the attributes of your metadata server
$objServerDef.MachineDNSName = "yournode.company.com"
$objServerDef.Port = 8561  # metadata server port
$objServerDef.Protocol = 2     # 2 = IOM protocol
# Class Identifier for SAS Metadata Server
$objServerDef.ClassIdentifier = "0217E202-B560-11DB-AD91-001083FF6836"

# connect to the server
# we'll get back an OMI handle (Open Metadata Interface)
try {
    $objOMI = $objFactory.CreateObjectByServer(
        "", 
        $true, 
        $objServerDef, 
        "sasdemo", # metadata user ID
        "Password1" # password
    )

    Write-Host "Connected to " $objServerDef.MachineDNSName 
}
catch [system.exception] {
    Write-Host "Could not connect to SAS metadata server: " $_.Exception
    exit -1
}

# get list of repositories
$reps = "" # this is an "out" param we need to define
$rc = $objOMI.GetRepositories([ref]$reps, 0, "")

# parse the results as XML
[xml]$result = $reps

# filter down to "Foundation" repository
$foundationNode = $result.Repositories.Repository | ? { $_.Name -match "Foundation" } 
$foundationId = $foundationNode.Id

Write-Host  "Foundation ID is $foundationId"  

$libTemplate = 
"<Templates>" +
"<PhysicalTable/>" +
"<Column SASColumnName=`"`" SASColumnType=`"`" SASColumnLength=`"`"/>" +
"<SASLibrary Name=`"`" Engine=`"`" Libref=`"`"/>" +
"</Templates>"
    
$libs = ""

# Use GetMetadataObjects method
# Usage is similar to PROC METADATA, so you
# can look at PROC METADATA doc to get examples
# of templates and queries

# 2309 flag plus template gets table name, column name,  
# engine, libref, and object IDs. The template specifies 
# attributes of the nested objects.                      

$rc = $objOMI.GetMetadataObjects(
    $foundationId, 
    "PhysicalTable", 
    [ref]$libs, 
    "SAS",  
    2309,
    $libTemplate
)
  
# parse the results as XML
[xml]$libXml = $libs

Write-Host "Total tables discovered: " $libXml.Objects.PhysicalTable.Count

# Create output, which you can pipe to another cmdlet
# such as Out-Gridview or Export-CSV

# for each column in each table, create an output object
# (named $objCol here)
for ($i = 0; $i -lt $libXml.Objects.PhysicalTable.Count; $i++) {
    $table = $libXml.Objects.PhysicalTable[$i]
    for ($j = 0; $j -lt $table.Columns.Column.Count ; $j++) {
        $column = $table.Columns.Column[$j]
        $objCol = New-Object psobject
        $objCol | add-member noteproperty -name "Libref" -value $table.TablePackage.SASLibrary.Libref
        $objCol | add-member noteproperty -name "Table" -value $table.SASTableName
        $objCol | add-member noteproperty -name "Column" -value $column.SASColumnName
        $objCol | add-member noteproperty -name "Type" -value $column.SASColumnType
        $objCol | add-member noteproperty -name "Length" -value $column.SASColumnLength
    
        # emit the object to stdout or other cmdlet
        $objCol
    }
}