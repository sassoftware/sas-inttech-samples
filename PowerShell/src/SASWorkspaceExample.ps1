# Example of how to use PowerShell to script the
# SAS Integration Technologies client
# You can connect to a remote SAS Workspace
# and run a program, retrieve the SAS log and listing

# To use: change this script to reference your SAS Workspace
# node name, port (if different), and user credentials

# create the Integration Technologies objects
$objFactory = New-Object -ComObject SASObjectManager.ObjectFactoryMulti2
$objServerDef = New-Object -ComObject SASObjectManager.ServerDef
$objServerDef.MachineDNSName = "sasnode.mycompany.com" # SAS Workspace node
$objServerDef.Port = 8591  # workspace server port
$objServerDef.Protocol = 2     # 2 = IOM protocol
# Class Identifier for SAS Workspace
$objServerDef.ClassIdentifier = "440196d4-90f0-11d0-9f41-00a024bb830c"

# create and connect to the SAS session 
$objSAS = $objFactory.CreateObjectByServer(
    "SASApp", # server name
    $true, 
    $objServerDef, # built server definition
    "sasdemo", # user ID
    "secretPassword"    # password
)

# program to run
# could be read from external file
$program = "options formchar='|----|+|---+=|-/\<>*';"  
$program += "ods listing; proc means data=sashelp.cars mean mode min max; run;"

# run the program
$objSAS.LanguageService.Submit($program);

# flush the output - could redirect to external file
Write-Output "Output:"
$list = ""
do {
    $list = $objSAS.LanguageService.FlushList(1000)
    Write-Output $list
} while ($list.Length -gt 0)


# flush the log - could redirect to external file
Write-Output "LOG:"
$log = ""
do {
    $log = $objSAS.LanguageService.FlushLog(1000)
    Write-Output $log
} while ($log.Length -gt 0)

# end the SAS session
$objSAS.Close()