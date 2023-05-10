<# --------------------------------------------
 File:		ADGroupManipulation.ps1
 Author:    Nathan Carpenter, Network Analyst
 Purpose:	Used for removing and adding list of 
            users to grourps in AD
 Requires:  Active Directory Module 
            
            
             
 VVSD History:
 1.0.0		Created - 08/19/2021 - NC
 1.1.0      Addition - 08/31/2021 - NC (Added Controls 3 and 4 below)
 1.2.0      Update - 06.21.22 - NC (Updated LoadInputFile function to open properly)
 1.3.0      Update - 08.17.22 - NC (Added option 5 to the list and cooresponding function)
             
    
-------------------------------------------- #>

function LoadInputFile  ($InitialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}


function Remove-UsersFromGroup ($infile) {
    $infile = LoadInputFile -InitialDirectory "C:\"
    $usersfile = Import-Csv $infile
    [string]$GroupToLeave = Read-Host "Name of AD Group to Remove Members From"
    $counter = 0

    foreach ($User in $Usersfile) {
        $counter++
        $samacc = $user.Userprincipalname

        Remove-ADGroupMember -Identity $GroupToLeave -Members $samacc -Confirm:$false
        Start-Sleep -Seconds 1

        Write-Progress -Activity "Removing Members from $grouptoleave" -CurrentOperation $samacc -PercentComplete (($counter / $usersfile.count) * 100)
    }
}

function Add-AdUsersToGroup ($infile) {
    $infile = LoadInputFile -InitialDirectory "C:\"
    $usersfile = Import-Csv $infile
    [string]$GroupToJoin = Read-Host "Name of AD Group to Add Members to"
    #$credential = Get-Credential
    $counter = 0

    foreach ($User in $Usersfile) {
        $counter++
        $samacc = $user.Userprincipalname

        Add-ADGroupMember -Identity $GroupToJoin -Members $samacc -Confirm:$false <#-server "" -Credential $credential#>
        Start-Sleep -Seconds 1

        Write-Progress -Activity "Adding Members to $grouptoJoin" -CurrentOperation $samacc -PercentComplete (($counter / $usersfile.count) * 100)
    }
}

function CompareGroupMembershipFromCSV ($inputfileUsers,$inputfilegroups) {
    Write-host "Please Select the user File to upload, with the header of 'UserPrincipalName' in the CSV"
        Start-Sleep -Seconds 3
        $inputfileUsers = LoadInputFile -InitialDirectory "C:\"

    Write-Host "Please Select the Groups file to upload, with the header of 'Groups' in the CSV"
        Start-Sleep -Seconds 3
        $inputfilegroups = LoadInputFile -InitialDirectory "C:\"

$users = Import-Csv $inputfileUsers
$groups = Import-Csv $inputfilegroups

    foreach ($group in $groups) {
        $groups = $group.groups
        $members =  @{}
        $members = Get-ADGroupMember -Identity $groups -Recursive | Select-Object -ExpandProperty Name
        Write-Host "$group Members (red for missing, Green for Included)"
        ForEach ($user in $users) {
            $upn = $user.Userprincipalname
            If ($members -contains $upn) {
                Write-Host "$upn exists in the group" -ForegroundColor Green
                Start-Sleep -Seconds 1
        } Else {
            Write-Host "$upn dose not exists in the group" -ForegroundColor Red
            Start-Sleep -Seconds 1
            
           <#  $export = #> Get-ADUser -Identity $upn -Properties Name,DisplayName | Select-Object DisplayName,Name | Export-Csv "C:\export\MembersMissingFrom_$group.csv" -Append -NoTypeInformation
          }
        }
        #$export | Export-Csv "C:\export\MembersMissingFrom_$group.csv" -Append -NoTypeInformation
    }    
}

function DisplayNameToUPN () {
    $inputfile = LoadInputFile -InitialDirectory "C:\"
    $usersfile = Import-Csv $inputfile

    ForEach ($User in $usersfile<# (Get-Content $inputfile | ConvertFrom-CSV -Header FirstName,LastName) #>)
    {  
        $Filter = "givenName -like ""*$($User.FirstName)*"" -and sn -like ""$($User.LastName)"""
        Get-ADUser -Filter $Filter | Select-Object Name,DisplayName | Sort-Object Name | Export-Csv "C:\export\BHSSPEDUPNS.csv" -Append


        }
}

function FirstnameLastname2Group ($infile) {
    $inputfile = LoadInputFile -InitialDirectory "C:\"
    $usersfile = Import-Csv $inputfile
    [string]$GroupToJoin = Read-Host "Name of AD Group to Add Members to"

    $counter = 0
    ForEach ($User in $usersfile<# (Get-Content $inputfile | ConvertFrom-CSV -Header FirstName,LastName) #>)
    {  
        $counter++
        $Filter = "givenName -like ""*$($User.FirstName)*"" -and sn -like ""$($User.LastName)"""
        $UPN = Get-ADUser -Filter $Filter -Properties Name,DisplayName | Select-Object Name,DisplayName
        $UPNDisplayName = $UPN.DisplayName | Out-String
        <# Write-Output $UPN.Name
        Write-Output-$UPN.DisplayName
        Pause #>
        If ($UPN.Name -notcontains "_RTU") {
        Add-ADGroupMember -Identity $GroupToJoin -Members $UPN.Name -Confirm:$false <#-server "" -Credential $credential#>
        Start-Sleep -Seconds 1
        }
        else {
            Write-Host "UPN Contains _RTU account, skipping..." -ForegroundColor Red
            Start-Sleep 1
        }

        Write-Progress -Activity "Adding Members to $grouptoJoin" -CurrentOperation $UPNDisplayName -PercentComplete (($counter / $usersfile.count) * 100)
        }
}

Write-Host "Please Select One of the Following Options..."
Write-Host "1. Remove Members from a Group"
Write-Host "2. Add Members to a Group"
Write-Host "3. Gather UPNs from DisplayNames (CSV needs to have FirstName in one Column and LastName in the other as headers)"
Write-Host "4. Compare Group MemberShips in AD from a CSV of Users (CSV of users needs 'UserPrincipalName' as the header; CSV of Groups needs 'Groups' as the header"
Write-Host "5. Adds Members to group using their lastname and firstname (comnination of tools 2&3)"
$choice = Read-Host "Enter your selection (number)"

switch ($choice) {
    "1" { 
        Remove-UsersFromGroup

        ;Break
     }
     "2" {
         Add-AdUsersToGroup

         ;Break
     }
     "3" {
         DisplayNameToUPN

         ;Break
     }
     "4" {
         CompareGroupMembershipFromCSV

         ;Break
     }
     
     "5" {
        FirstnameLastname2Group
        
        ;Break
     }
    Default {
        Write-Host "Invaild Choice, please launch script again"

        ;Break
    }
}