
function New-RandomPassword {
    param(
        [Parameter()]
        [int]$MinimumPasswordLength = 12,
        [Parameter()]
        [int]$MaximumPasswordLength = 24,
        [Parameter()]
        [int]$NumberOfAlphaNumericCharacters = 8,
        [Parameter()]
        [switch]$ConvertToSecureString
    )
    
    Add-Type -AssemblyName 'System.Web'
    $length = Get-Random -Minimum $MinimumPasswordLength -Maximum $MaximumPasswordLength
    $password = [System.Web.Security.Membership]::GeneratePassword($length,$NumberOfAlphaNumericCharacters)
    if ($ConvertToSecureString.IsPresent) {
        ConvertTo-SecureString -String $password -AsPlainText -Force
    } else {
        $password
    }
}

function RemoteADSync ($computerName) {
    $credentials = Get-Credential -Message "Supply ADMSU Account to access $($computerName)"

    $remotesession = New-PSSession -ComputerName $computerName -Credential $credentials
    $job = Invoke-Command -Session $remoteSession -ScriptBlock {Start-AdSyncSyncCycle -policyType Delta} -AsJob
    $job
    Start-Sleep -Seconds 1
    Get-Job $job.id
    Receive-Job $job.id

    Start-Sleep -Seconds 5

    Disconnect-PSSession $remotesession
    Start-Sleep -Seconds 5
    Remove-PSSession $remotesession
    
}


function RecoverInactiveMailbox ($inactiveMailbox, $newUserMailbox) {
    
    #check For Exchange Online Powershell and install if not
    Write-Host "Checking for Exchange Online Module..." -ForegroundColor Yellow
        $ExchangeModuleInfo = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" | Format-Table -Property Name, Version, ModuleType

    
    if ($ExchangeModuleInfo -eq '') {
        Write-Host "Exchange Online Module not found... module required to continue"
        $InstallModule = Read-Host "Would you like to install Exchnage Powershell Module now, under the current user (y/N)?"
            if ($installModule -eq "y") {
                Write-Host "Installing Powershell Module as Current User..."
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser 
                Connect-ExchangeOnline
            }
            else {
                Write-Host "EXO module needed to continue, closing script..."
                Exit 1
            }
    }
    else {
        Write-Host "Exchnage Online Module Found!" -ForegroundColor Green
        Write-Output $($ExchangeModuleInfo.Version)
        Connect-ExchangeOnline 
        #Pause
    }
}
    function GatherUserInfo ($InactiveMailBoxUser,$RecoverToUser) {
        if ($null -eq $inactiveMailBoxUser) {
            $inactiveMailBoxUser = Read-Host "Please enter the email address of the Inactive Mailbox to Restore:"
            $Global:inactiveMailbox = Get-EXOMailbox -InactiveMailboxOnly -Identity $inactiveMailBoxUser | Where-Object PrimarySmtpAddress -EQ $inactiveMailBoxUser | select-object UserPrincipalName,DisplayName,PrimarySmtpAddress,RecipientType,id,DistinguishedName

            if ($null -eq $Global:inactiveMailbox) {
                Write-Host "Inactive mailbox not found, please relaunch program and try again" -ForegroundColor Red
            }
        }
        $Global:inactiveMailbox = Get-EXOMailbox -InactiveMailboxOnly -Identity $inactiveMailBoxUser| Where-Object PrimarySmtpAddress -EQ $inactiveMailBoxUser | Select-Object UserPrincipalName,DisplayName,PrimarySmtpAddress,RecipientType,id,DistinguishedName
        if ($null -eq $RecoverToUser) {
            $RecoverToUser = Read-Host "Please enter the UPN of the user to restore the mailbox to(without @vvsd.org):"
            $Global:UserRecoveryInfo = Get-ADUser -Identity $RecoverToUser -Properties Name,givenName,sn,displayname,mail,ObjectGUID

            if ($null -eq $Global:UserRecoveryInfo) {
                Write-Host "User not found in Active Directory, please relaunch program and try again" -ForegroundColor Red
            }
        }
        $Global:UserRecoveryInfo = Get-ADUser -Identity $RecoverToUser -Properties Name,givenName,sn,displayname,mail,ObjectGUID
    }

    function PerformRecover () {
        Write-Host "Please Confirm the following is the Inactive Mailbox you want to restore"
        Write-Output $Global:inactiveMailbox | Format-Table UserPrincipalName,DisplayName,DistinguishedName
        Pause

        Write-Host "Pleas Confirm the following user to recovery the Mailbox to"
        Write-Output $Global:UserRecoveryInfo | Format-Table Name,GivenName,sn,displayname,mail,ObjectGUID
        Pause

        $RestoreConfirmation = Read-Host "Perform Restore of Inactive Mail to defined user(y/N):" 

        if ($RestoreConfirmation -eq 'Y') {
            Connect-MsolService 

             #Check for User in RecycleBin Matching GUID
             $RecycleBinCheck = Get-MsolUser -ReturnDeletedUsers |  Where-Object Userprincipalname -eq  $Global:UserRecoveryInfo.mail | Select-Object Userprincipalname,immutableID

             if ($null -ne $RecycleBinCheck) {
                Write-Host "Recycle Bin Check == TRUE, comencing deletion from AzureAD environment..." -ForegroundColor Red
                $confirmDel = Read-Host "Please confirm deletion of $($RecycleBinCheck.Userprincipalname)(y/N)"
                    if ($confirmDel -eq "y") {
                        Remove-MsolUser -UserPrincipalName $RecycleBinCheck.UserPrincipalName -RemoveFromRecycleBin

                        $CheckDeletionAction = Get-MsolUser -ReturnDeletedUsers | Format-List UserprincipalName,immutableID | Where-Object UserprincipalName -eq $RecycleBinCheck.Userprincipalname
                            if ($null -eq $CheckDeletionAction) {
                                Write-Host "User not found... deletion confirmed" -ForegroundColor Green
                            }
                            elseif ($null -ne $checkDeletionAction) {
                                Write-Output $CheckDeletionAction
                                Write-Host "User found in RecycleBin... deletion unsuccessful..." -ForegroundColor Red
                            }
                        }
                }   
        }

            $newpswd = New-RandomPassword -ConvertToSecureString

            $TempAlias = "Alias-$($Global:UserRecoveryInfo.Name)@365u.onmicrosoft.com"
            Write-Output $TempAlias
            $TempName = "Alias-$($Global:UserRecoveryInfo.Name)"
            Write-Output $TempName

           # Try 
            #{
            New-Mailbox -InactiveMailbox $Global:InactiveMailbox.DistinguishedName -Name $TempName -Alias $TempName -DisplayName $Global:UserRecoveryInfo.displayname -MicrosoftOnlineServicesID $TempAlias -Password $newpswd -ResetPasswordOnNextLogon $true
           # }
           # {
                #Write-Host "An Error has Occured when trying to create a mailbox" -ForegroundColor Red
            #}
             $GUID2ImmutableID = [system.convert]::ToBase64String(([GUID]"$($Global:UserRecoveryInfo.ObjectGUID)").tobytearray())

            Write-Host "Pausing for 30 seconds to allow for Account to be created..." -ForegroundColor Yellow
            Start-Sleep -Seconds 30
            #https://github.com/MicrosoftDocs/OfficeDocs-Support/blob/public/Exchange/ExchangeOnline/administration/reconnect-inactive-or-soft-deleted-mailboxes-to-ad.md#reconnect-the-original-inactive-mailbox
            Set-MsolUser -UserPrincipalName $TempAlias -ImmutableId $GUID2ImmutableID
            Start-Sleep -Seconds 5
            #Move AD Object to Sync'd OU
            Write-Host "Moving Local AD Object to OU " -ForegroundColor Yellow -NoNewline
            Write-Host "USERS " -ForegroundColor Magenta -NoNewline
            Write-Host "Please remember to move to proper Home OU after this..."
            Move-ADObject -Identity $Global:UserRecoveryInfo.ObjectGUID -TargetPath "OU=Users,DC=vvsd,DC=org"

            Write-Host "Starting AD Sync to reconnect On-Prem to Cloud" -ForegroundColor Yellow
            RemoteADSync -computerName "DC-TC.vvsd.org"
            Start-Sleep -Seconds 15

                if (Get-MsolUser -UserPrincipalName $TempAlias) {
                    Write-Host "Temporary Alias Removed..." -ForegroundColor DarkGreen -BackgroundColor Green
                    
                }
                if (Get-MsolUser -UserPrincipalName $Global:UserRecoveryInfo.mail) {
                    Write-Host "Sync Successful, account has been registered in AAD" -ForegroundColor DarkGreen -BackgroundColor Green
                }

}
RecoverInactiveMailbox
GatherUserInfo
PerformRecover
Disconnect-ExchangeOnline -Confirm:$true