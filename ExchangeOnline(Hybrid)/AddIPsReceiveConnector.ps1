


function LoadInputFile  ($InitialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv|TXT (*.txt) | *.txt"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

function RemoteConnectEAC () {
    #Checks for Existing Session 
    $FoundSession = Get-PSSession | Where-Object {$_.ComputerName.ToLower().Contains("ex11")}
    #Starts session with ex11 Exchange management
    if ($null -eq ($FoundSession))
    {
        $DACredential = Get-Credential -message "Please Provide ADM Account Credentials"
        $Netex11Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ex11.vvsd.org/PowerShell/ -Authentication Kerberos -Credential $DACredential
        Import-PSSession $Netex11Session -AllowClobber
    }
    elseif ($FoundSession) {
        Import-PSSession $FoundSession

    }
    else {
        Write-Host "No session information found, or error in script..." -ForegroundColor Red
        Write-Host "Check Script and try again, exiting now"
        Exit 0
    }
    }

#Inital Script Introduction and info
Write-Host "Now Entering Script to authorize IPs to relay Mail Out to Web..."
Write-Host "Beginning Remote Connection too EX11 for commands" -ForegroundColor Blue
Start-Sleep -Seconds 5

#Run Functionn to Connect to EX11 before continuing Logic
RemoteConnectEAC

#Pause and explain the format needed for the CSV of IPs
Write-Host "Before continuing, please ensure that your list of IPs youd like to add to the Send Connector are in a CSV file (NO HEADER)..." -ForegroundColor Yellow
Read-Host "Press any key to continue to File Upload"


#Set Variables
$filepath = LoadInputFile -InitialDirectory "C:\"
$receiveconnector = "Anonymous SMTP Relay"
    
# Import IP addresses from CSV file
$IPs = Import-Csv $filepath

# Get receive connector
$RCon = Get-ReceiveConnector $receiveconnector

# Get receive connector remote IP addresses
$RemoteIPRanges = $RCon.RemoteIPRanges

# Loop through each IP address
foreach ($IP in $IPs) {
    $IPEx = $IP.Expression

    # Check if IP addres already exist
    if ($RemoteIPRanges -contains $IPEx) {
        Write-Host "IP address $($IPEx) already exist in receive connector $($receiveconnector)" -ForegroundColor Red
    }
    
    # If IP address not exist than add IP address
    else {
        $RemoteIPRanges += $IPEx

        # Remove the -WhatIf parameter after you tested and are sure to add the remote IP addresses
        Set-ReceiveConnector $receiveconnector -RemoteIPRanges $RemoteIPRanges #-WhatIf
        Write-Host "IP address $($IPEx) added to receive connector $($receiveconnector)" -ForegroundColor Green
    }
    Start-Sleep -Milliseconds 500
}

Get-PSSession | Remove-PSSession
