<# --------------------------------------------
 File:		BulkADEdits.ps1
 Author:    Nathan Carpenter, Network Analyst
 Purpose:	Used for Cleanup of AD and Administration
            of accounts in bulk
 Requires:  .Net Framework 
            
            
             
 VVSD History:
 1.0.0		Created - 06/11/2021 - NC
             
    
-------------------------------------------- #>


function LoadInputFile  ($InitialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv|TXT (*.txt) | *.txt"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

function New-RandomPassword() {
    param(
        [Parameter()]
        [int]$MinimumPasswordLength = 16,
        [Parameter()]
        [int]$MaximumPasswordLength = 24,
        [Parameter()]
        [int]$NumberOfAlphaNumericCharacters = 12,
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

function RemoveAllOfficeLicensing ($infile) {
    
    foreach ($user in $infile) {
        $upn = $user.UserPrincipalName
        
        $licensedusers = Get-MsolUser -UserPrincipalName $upn | Where-Object {$_.isLicensed}
    
        foreach ($user in $licensedusers) {
            Write-Host "$($user.displayname)" -ForegroundColor Yellow
            $licenses = $user.Licenses
            $licenseArray = [System.Collections.ArrayList]::new()
            $licenseArray = @($licenses | foreach-Object {$_.AccountSkuId})
            $licenseString = $licenseArray -join " ,"
            Write-Host "$($user.displayname) has $licenseString" -ForegroundColor Blue

            for ($i = 0; $i -lt $licensearray.Count; $i++) {

                Write-Host "-----------------------------------------------"
                Write-Host "License $i of " $licensearray.count
                Write-Host "Removing" $licensearray[$i] "from" $user.DisplayName
                Write-Host "-----------------------------------------------"

                Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $licenseArray[$i]
                Start-Sleep -Seconds 3
            }
        }

        $verify = Get-MsolUser -UserPrincipalName $upn | Select-Object DisplayName,isLicensed,licenses

        if ($verify.Islicensed -eq "False") {

            Write-Host "---------------------------------------------------"
            Write-Host "All Licensing has been removed from $($verify.displayname)"
            Write-Host "---------------------------------------------------"
            
        }
        else {
            Write-Host "---------------------------------------------------"
            Write-Host "Removal failed, logging result..."
            Write-Host "---------------------------------------------------"

            $verify | Export-Csv "C:\export\FailedToRemoveLicense.csv" -NoTypeInformation -Append
        }
     }
}

function DisableADAccount ($infile) {

    foreach ($user in $infile) {
        $samacc = $user.samaccountname

        $AccountStatus = Get-ADUser -Identity $samacc -Properties name,enabled,whenCreated | Select-Object name,enabled,whenCreated
        

        if ($AccountStatus.enabled -eq "True") {
            Write-Host "----------------------------------------------"
            Write-Host "$samacc is enabled, disabling account..." -ForegroundColor Green
            Write-Host "----------------------------------------------"
                Start-Sleep -s 2
            Disable-ADAccount -Identity $samacc

           # $verify = Get-ADUser -Identity $samacc | Select-Object enabled

        }
        else {
            Write-Host "----------------------------------------------"
            Write-Host "$samacc is already disabled."
            Write-Host "----------------------------------------------"
                Start-Sleep -s 2
        }
    }
}

function Update-ADPassword ($infile) {

    foreach ($user in $infile) {
        $samacc = $user.SAMAccountName

        $TimeBeforeUpdate = Get-ADUser -Identity $samacc -Server DC-TC.vvsd.org -Properties name,enabled,passwordlastset | Select-Object name,enabled,passwordlastset

        $newpswd = New-RandomPassword -ConvertToSecureString

        Set-ADAccountPassword -Identity $samacc -Server DC-TC.vvsd.org -Reset -NewPassword $newpswd

        $TimeAfterupdate = Get-ADUser -Identity $samacc -Server DC-TC.vvsd.org -Properties name,enabled,passwordlastset | Select-Object name,enabled,passwordlastset
        #$todaysdate = Get-Date -Format "MM/dd/yyyy HH:mm"

        if ($TimeAfterupdate.passwordlastset -ge $TimeBeforeUpdate.passwordlastset) {
            Write-Host "----------------------------------------------"
            Write-Host "Password successfully updated for:" $TimeAfterupdate.Name -ForegroundColor Green
            Write-Host "Password last set at: " $TimeAfterupdate.passwordlastset -ForegroundColor Green
            Write-Host "----------------------------------------------"
            Start-Sleep -Seconds 2
        }
        else {
            Write-Host "----------------------------------------------"
            Write-Host "Password Update Failed. Logging Results..." -ForegroundColor Red
            Write-Host "----------------------------------------------"
            Start-Sleep -Seconds 2

            $TimeAfterupdate | Export-Csv "C:\export\PasswordUpdateFailed.csv" -NoTypeInformation -Append
        }

        
    }
}

function Move-UsersToOU ($infile) {
    $credential = Get-Credential

    $TargetOU = Get-ADOrganizationalUnit -filter * -Server DC-TC.vvsd.org |
        Select-Object DistinguishedName, Name | Sort-Object DistinguishedName | Out-GridView -PassThru -Title "Choose OU to Move Users To"
    
    foreach ($user in $infile) {
        $samacc = $user.samaccountname
        $UserGUID = Get-ADUser -Identity $samacc -Properties objectGUID | Select-Object -ExpandProperty objectGUID 

        Move-ADObject -Identity "$UserGUID" -TargetPath $TargetOU.DistinguishedName -Credential $credential
    }
}

Write-Host "Please Select One of the Following options"
Write-Host "1. Delete Users Account"
Write-Host "2. Disabled User Accounts and remove Office Licensing"
Write-Host "3. Disable Account, Change to Random Password, and Remove Office Licensing"
Write-Host "4. Move users to Specific OU"
$Choice = Read-Host "Choice(number)"

switch ($choice) {
    "1" {
        $loadfile = LoadInputFile -InitialDirectory C:\
        $infile = Import-Csv $loadfile

        foreach ($user in $infile) {
            $user.samaccountname
            $samacc = $user.samaccountname

            Remove-ADUser -Identity $samacc -confirm
        }
        ; Break
      }
    "2" {
        $loadfile = LoadInputFile -InitialDirectory "C:\"
        $infile = Import-Csv $loadfile

        Write-Host "Please Login with CloudSU account to manage Office Licensing:"
        Start-Sleep -Seconds 3
        Connect-MsolService

        DisableADAccount -infile $infile

        RemoveAllOfficeLicensing -infile $infile

     ; Break
    }
    "3" {
        $loadfile = LoadInputFile -InitialDirectory "C:\"
        $infile = Import-Csv $loadfile

        DisableADAccount -infile $infile

        Update-ADPassword -infile $infile

        Connect-MsolService
        RemoveAllOfficeLicensing -infile $infile
     
        ; Break
    }
    "4" {
        $loadfile = LoadInputFile -InitialDirectory "C:\"
        $infile = Import-Csv $loadfile

        Move-UsersToOU -infile $infile
    }
    Default {
        Write-Host "Invalid Choice. Please relaunch script."
        End

    }
}