# Example of how to use PowerShell to script the
# SAS Integration Technologies client
# You can connect to a remote SAS Workspace
# and run a program, retrieve the SAS log and download a file

# To use: change this script to reference your SAS Workspace
# node name, port (if different), and user credentials

# create the Integration Technologies objects
$objFactory = New-Object -ComObject SASObjectManager.ObjectFactoryMulti2
$objServerDef = New-Object -ComObject SASObjectManager.ServerDef
$objServerDef.MachineDNSName = "yournode.company.com" # SAS Workspace node
$objServerDef.Port = 8591  # workspace server port
$objServerDef.Protocol = 2     # 2 = IOM protocol
# Class Identifier for SAS Workspace
$objServerDef.ClassIdentifier = "440196d4-90f0-11d0-9f41-00a024bb830c"

try {
    # create and connect to the SAS session 
    $objSAS = $objFactory.CreateObjectByServer(
        "SASApp", # server name
        $true, 
        $objServerDef, # built server definition
        "sasdemo", # user ID
        "Password1"    # password
    )
}
catch [system.exception] {
    Write-Host "Could not connect to SAS session: " $_.Exception.Message
}

# change these to your own SAS-session-based
# file path and file name
# Note that $destImg can't be > 7 chars
$destPath = "c:\DataSources"
$destImg = "hist"

# local directory for downloaded file
$localPath = "c:\temp"

# program to run
# could be read from external file
$program =  
"ods graphics / imagename='$destImg';
        ods listing gpath='$destPath' style=plateau;
        proc sgplot data=sashelp.cars;
        histogram msrp;
        density msrp;
        run;"

# run the program
$objSAS.LanguageService.Submit($program);

# flush the log - could redirect to external file
Write-Output "SAS LOG:"
$log = ""
do {
    $log = $objSAS.LanguageService.FlushLog(1000)
    Write-Output $log
} while ($log.Length -gt 0)

# now download the image file
$fileref = ""

# assign a Fileref so we can use FileService from IOM
$objFile = $objSAS.FileService.AssignFileref(
    "img", "DISK", "$destPath/$destImg.png", 
    "", [ref] $fileref);

$StreamOpenModeForReading = 1
$objStream = $objFile.OpenBinaryStream($StreamOpenModeForReading)

# define an array of bytes
[Byte[]] $bytes = 0x0

$endOfFile = $false
$byteCount = 0
$outStream = [System.IO.StreamWriter] "$localPath\$destImg.png"
do {
    # read bytes from source file, 1K at a time
    $objStream.Read(1024, [ref]$bytes)
  
    # write bytes to destination file
    $outStream.Write($bytes)
    # if less than requested bytes, we're at EOF
    $endOfFile = $bytes.Length -lt 1024
  
    # add to byte count for tally
    $byteCount = $byteCount + $bytes.Length
  
} while (-not $endOfFile)

# close input and output files
$objStream.Close()
$outStream.Close()

# free the SAS fileref
$objSAS.FileService.DeassignFileref($objFile.FilerefName)

Write-Output "Downloaded $localPath\$destImg.png: SIZE = $byteCount bytes" 

$objSAS.Close()