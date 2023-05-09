<#
.SYNOPSIS
    Gathers Information from servers in a Domain and from a TXT doc of server names (FQDN preferred)
.DESCRIPTION
    THis script will cycle through a list of machines in a text doc and attempt to check for life via a Ping, then gather relevent information about
    the device, such as CPU,MEM,HDD,Type (VM/Physical), vendor, and IPs associated and export to C:\temp on machine
.NOTES
    This may need to be adapated to your env, as this uses WINRM and PSSessions to accomplish the job.
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>




#define domain suffix here for easier 
$global:DomainSuffix = 


function Start-PingTest ($ServerInput) {

    #Initalize Global Arrays
$Global:OnlineServers = @()
$Global:OfflineServers = @()

#loop thriough list of users and Ping for Connectivity
foreach($server in $ServerInput){
    $PingServer = $server + $global:DomainSuffix

    $PingTest = Test-Connection -ComputerName $PingServer -Quiet -Count 2 -ErrorAction SilentlyContinue 
    Start-Sleep -Milliseconds 250

        If($PingTest){
            Write-Host($PingServer + " is Online!") -ForegroundColor DarkGreen -BackgroundColor Green

           $Global:OnlineServers += $PingServer
            #Write-Output $OnlineServers
        }
        else{
            Write-Host($PingServer + " is Offline :'( ") -ForegroundColor DarkRed -BackgroundColor Red

            $Global:OfflineServers += $PingServer
            $Global:OfflineServers | Out-File -FilePath "C:\export\UnableToPing.txt" -Append
        }
}
}


function CheckADMembership ($ArrayInput) {

    #Intialize Global Arrays
    $Global:DomainJoinedFalse = @()
    $Global:DomainJoinedTrue = @()

    $Devices = $ArrayInput

foreach ($device in $Devices) {
    $hostname = $device -replace $global:DomainSuffix,''
try {
    Get-ADComputer -Identity $hostname -Properties CanonicalName -ErrorAction Stop | Select-Object CanonicalName
    Write-Host "$($hostname) is Domain Joined!" -BackgroundColor DarkGreen -ForegroundColor Green 
    $Global:DomainJoinedTrue += $hostname
}
catch {
    Write-Host "$($hostname) is NOT Domain Joined. Different Credentials Needed for PSSession" -BackgroundColor DarkRed -ForegroundColor Red
    $Global:DomainJoinedFalse += $hostname
}
}
}

function GatherComputerInfo ($Servers,$Credential,$localCredentialneeded) {
    
#Need Domain Admin rights to access all Servers
#$credential = Get-Credential -Message "Please Provide ADM Credentials for PSSessions to be Created"

foreach ($server in $Servers) {
    $servers = $server #+ $global:DomainSuffix

    if ($localCredentialneeded -eq $true) {
        $credential = Get-Credential -Message "Please provide Local Credential for $($server)" -UserName "$server\"
    }

    #Initate Remote connection to each server and run Following Script
    Invoke-Command -ComputerName $servers -credential $credential -scriptblock {
    $currenthostname = $env:computername
    #Get all installed Roles and Features in Windows server
    $installedFeatures = Get-WindowsFeature | Where-Object {$_.installstate -eq "installed"} | Select-Object -ExpandProperty DisplayName
        $installedFeatures2=[system.String]::Join("|", $installedFeatures)
    #Gather Basic Computer info
    $Computerinfo = Get-ComputerInfo -property OsName,OsBuildNumber,OsVersion,BiosSeralNumber,CsManufacturer,CsModel,CsProcessors,CsNumberOfLogicalProcessors,CsNumberOfProcessors,CsPhyicallyInstalledMemory
    $MemoryInfo = [math]::round($Computerinfo.CsPhyicallyInstalledMemory/1MB, 2)
    <# $DriveInfo = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object DriveType -eq 3 | Select-Object DriveID,VolumeName,Size,FreeSpace
    $driveMaxSize = [math]::round($DriveInfo.Size/1GB, 2)
    $DriveUsedSize = [math]::round($DriveInfo.FreeSpace/1GB, 2)
    $DriveFreeSpace = [math]::round($DriveUsedSize/$DriveMaxSize*100, 2) #>
    #Get IP addresses 
    $gethostIP = Resolve-DnsName -Name $currenthostname | Where-Object -property Type -eq "A" | Select-Object Name,IPAddress

        if($ComputerInfo.CsModel -eq "Virtual Machine"){
    $hyperVhost = (get-item "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters").GetValue("HostName")
        }
        else{
            $hypervhost = "Is Physical Machine"
        }

        $data = [PSCustomObject]@{
            'ServerName' = $currenthostname
            'Installed Roles' = $installedFeatures2
            'Operating System' = $Computerinfo.OsName
            'Operating System Version' = $Computerinfo.OsVersion
            'Build Number' = $Computerinfo.OsBuildNumber
            'Serial Number' = $ComputerInfo.BiosSeralNumber
            'Manufacturer' = $Computerinfo.CsManufacturer
            'Model' = $Computerinfo.CsModel
            'vNIC (VM/Host IPs)' = $gethostIP.IPAddress
            'Physical Host for VM' = $hyperVhost
            'Processor Type' = $Computerinfo.CsProcessors
            'Physical Processors' = $Computerinfo.CsNumberOfProcessors
            'Logical Processors' = $Computerinfo.CsNumberOfLogicalProcessors
            'Memory(GB)' = $MemoryInfo
           <#  'DriveLetter' = $DriveInfo.DriveID
            'VolumeName' = $DriveInfo.VolumeName
            'Free Space Left' = $DriveFreeSpace + " %" #>
        }

        Return $data
        Return $installedFeatures

    } -AsJob
}
}

#Import List of All Server, hostnames only
$serverlist = Get-Content "C:\temp\CurrentServers.txt"

#Performs PingTest for Online Servers and sets to array
Start-PingTest -ServerInput $serverlist

#Verification of servers to gather information on
Write-Host "List of Servers to run Information Gathering programs on:" -ForegroundColor Blue
Write-Output $Global:OnlineServers
Write-Host "Waiting to continue Script" -ForegroundColor Yellow
Pause

#Check if computers are domain joined against AD Computers 
CheckADMembership -ArrayInput $Global:OnlineServers

#Logic for Gathering information for domain Joined Computers
Write-Host "Preparing to Gather information from Domain Joined Computers... please provide Credentials" -BackgroundColor DarkYellow -ForegroundColor DarkBlue
$Credential1 = Get-Credential -Message "Please Provide ADMSU Credentials for PSSessions to be Created"
GatherComputerInfo -Servers $Global:DomainJoinedTrue -Credential $Credential1 -localCredentialneeded $false

#Logic for Gather information from Non-Domain joined Computers Local admin
<# Write-Host "Preparing to Gather Information from Non-Domain Joined Computers... please provide Credentials for each server" -BackgroundColor DarkYellow -ForegroundColor Blue
$Credential2 = Get-Credential -Message "Please Provide Local Credentials for $($Device)"
GatherComputerInfo -Servers $Global:DomainJoinedFalse -localCredentialneeded $true #>

#Gathering of Data from Jobs into Excel
while (Get-Job -State 'Running') {
    $jobs = Get-Job -State 'Completed'
    $jobs | Receive-Job | Export-Csv 'C:\export\NewCurrentServers.csv' -Append -NoTypeInformation
    $jobs | Remove-Job
    Start-Sleep -Milliseconds 200
}

$jobs = Get-Job -State 'Completed'
$jobs | Receive-Job | Export-Csv 'C:\export\NewCurrentServers.csv' -Append -NoTypeInformation
$jobs | Remove-Job
    Get-Job
    Get-Job | Remove-Job

