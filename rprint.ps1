#gO hEllHoUnDs

$MAX_ATTEMPTS=3

#clean connections
#net stop workstation /y
#net start workstation
#must be on WREN to use script
$wifi=netsh wlan show interfaces | Select-String '\sSSID'
if(([string]$wifi).IndexOf("WREN") -eq -1){
Write-Host "//WARNING//" -ForegroundColor RED
Write-Host "You are not on the WREN. You must be connected to the WREN to connect to company printers." -ForegroundColor Yellow
$response=Read-Host -Prompt "Continue? (y/n)"
while($response.ToLower() -ne 'n' -and $response.ToLower() -ne 'y'){
    $response=Read-Host -Prompt "Continue? (y/n)"
}
if($response.ToLower() -eq 'n'){
    Exit
}
}



Clear-Host
#user credentials
$user=Read-Host -Prompt "Type in your Office 365 username e.g. john.doe@web.edu"
$pass=Read-Host -Prompt "Type the password for your Office 365 account" -AsSecureString

#prepopulate oldprinters list
#get list of all printers
$currentPrinters=(Get-Printer | Format-List -Property Name | Out-String) -split '\r?\n'
$oldPrinters=@()
foreach($line in $currentPrinters){
    #if printer name contains old identifiers add to discard pile
    $identifiers=False #truncated for security concerns
    if($identifiers){
        $oldprinters += $line.Substring($line.IndexOf("\"))
    }
}

#ask for deletion or nah. Only ask if they exist.
if($oldPrinters.Length -gt 0){
    #Clear-Host
    Write-Host "You currently have"$oldPrinters.Length"old DREN/WPPSS6 printers on your computer that probably won't work"
    $response=Read-Host -Prompt "Delete old printers and declutter? (y/n)"
    while($response.ToLower() -ne 'n' -and $response.ToLower() -ne 'y'){
        $response=Read-Host -Prompt "Continue? (y/n)"
    }
    if($response.ToLower() -eq 'y')
{
    #delete old printers
    foreach($printer in $oldPrinters){
        Write-Host "Deleting"$printer
        Remove-Printer -Name $printer
    }
    Write-Host
    Write-Host "Done!"
    Start-Sleep -s 2
}
}
#ooga ooga


#connect to print server

$pass2=[System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))

#save credentials
cmdkey /add:print.westpoint.edu /user:$user /pass:$pass2 | Out-Null

Clear-Host
#let user know it's not broken
Write-Host "Attempting to connect to print server. This might take a second."


$output="empty"
$printserveraddr='' #removed for security concerns, follows \\server\share format
$printers=''

#create pretty progress bar to look at
#connect in background task
for ($attempt=0; $attempt -lt $MAX_ATTEMPTS; $attempt++){
    $printServerConnect=Start-Job -ScriptBlock{
        $output=net use  /persistent:yes | Out-String
        Write-Output $output
        }
    #progress bar while task incomplete
    $i=0
    while($printServerConnect.State -ne "Completed")
    {
        $status='connect attempt '+($attempt+1)+'...'
        Write-Progress -Id 1 -Activity 'Connecting to print server..' -Status $status -PercentComplete ($i % 100)
        $i++
        Start-Sleep 1
    }
    #if not failure, break
    $output=Receive-Job -Job $printServerConnect | Out-String
    #test if connected
    $printers=net view \\print.westpoint.edu\ | Out-String
    if($printers -ne ""){
        Break
    }
    else{
        $status='connection failed ('+($attempt+1)+'/'+$MAX_ATTEMPTS+')'
        Write-Progress -Id 1 -Activity 'Connecting to print server..' -Status $status -PercentComplete (0)
    }
    #wait between attempts 
    Start-Sleep 2
}
Write-Progress -Id 1 -Activity 'Connected to print server' -Status $i -Completed

#if task still failed (determine by looking at $printers) don't go further
if($printers -eq ""){
    Write-Host "Regrettably unable to connect to the server. Check if connected to WREN, or print server may be down." -ForegroundColor Red
    Start-Sleep -s 5
    Exit
}



$printers = net view $printserveraddr | Out-String
#only grab actual values
$printersArray = ($printers -split '\r?\n').Trim()
#powershell is weird and I don't like it
$end=$printersArray.Count-4
$printersArray=$printersArray[7..$end]

#define out of scope so can be used as break condition
$selection='none'
while($selection.ToLower() -ne 'q'){
    $itr=1
    Clear-Host
    Write-Host ============================ SELECT A PRINTER TO CONNECT TO ====================== -ForegroundColor Yellow 
    Write-Host
    $printers=@("base")
    foreach($printer in $printersArray){
        $prtstring=[string]$printer
        Write-Host $itr") " -ForegroundColor Yellow -NoNewline
        $start=$prtstring.IndexOf("Print")-1
        $deviceName=([string]$prtstring).Substring(0,$start).Trim()
        $start=$prtstring.IndexOf("Print")+"Print".Length
        $description=([string]$prtstring).Substring($start).Trim()
        Write-Host $deviceName -NoNewline -ForegroundColor Green
        $diff=" "*(20-$deviceName.Length)
        Write-Host $diff $description
        #assign to num
        $printers+=$deviceName
        $itr++
    }
    Write-Host
    Write-Host "============================ SELECT A PRINTER TO CONNECT TO ======================" -ForegroundColor Yellow
    Write-Host "*** hint, type the number of the printer you want and hit ENTER, e.g. 29" -ForegroundColor Yellow

    $selection=Read-Host "Selection (type q or CTRL+C to exit) "
    #loop to get proper input, don't accept if not q or not in array range
    while($selection.ToLower() -ne 'q' -and ($selection -notin (1..$printersArray.Length))){
        $selection=Read-Host "Selection (type q or CTRL+C to exit) "
    }
    #if exit then exit
    if($selection.ToLower() -eq 'q'){
        Exit
    }
    else{
        #otherwise attempt to add printer

        $selection=([string]$printers[$selection]).Trim()

        Write-Host "Connecting to" $selection....
        #add printer
        $output=Add-Printer -ConnectionName ('\\print.westpoint.edu\'+$selection) | Out-String
        if($output.IndexOf('error') -ne -1){
        Write-Host "Failed to connect to"$selection
        Write-Host "Check to make sure you're on the WREN, otherwise the print server may be down."
        }
        else
        {
        Write-Host "Successfully connected to"$selection
        }
        Start-Sleep 2
        $selection=''
    }
}

#bktiel 2019
