
Import-Module ImportExcel

$FileName = 'C:\export\Text.xlsx'

#$FunName = (Get-FunctionName)
$IEModule = (Get-Module -ListAvailable -Name "ImportExcel")

If ($IEModule -eq 0) {
    Write-Host "Please install the module ImportExcel to use the -Excel feature" -ForegroundColor Red
    Write-Host "Writing output to 'DomainUserAccounts.csv' instead."
    $Excel = $False
    $FileName = $FileName -replace '[^\\]+$',"$($FunName).csv"
}

$GroupList = [System.Collections.ArrayList]::new()

#$GroupList = Get-ADGroup -SearchBase "OU=Technology Security Groups,OU=Technology,OU=AncillaryLocations,DC=vvsd,dc=org" -Filter 'GroupCategory -eq "Distribution"' -Properties CN | Select-Object -ExcludeProperty CN

$GroupList = Get-ADGroup -Server DC-TC.vvsd.org -SearchBase "OU=District Wide Distribution Lists,OU=District Wide,DC=vvsd,DC=org" -Filter 'GroupCategory -eq "Distribution"' -Properties CN | Select-Object -ExpandProperty CN


#for ($i = 0; $i -lt $groupList.Count; $i++) {

foreach ($group in $GroupList) {

#$GroupInfo = Get-ADGroup -Identity $GroupList.cn[$i] -Properties CN,DisplayName,Info,Mail,ManagedBy,WhenCreated | Select-Object CN,DisplayName,Info,Mail,ManagedBy,WhenCreated

$GroupInfo = Get-ADGroup -Identity $group -Properties CN,DisplayName,Info,Mail,ManagedBy,WhenCreated | Select-Object CN,DisplayName,Info,Mail,ManagedBy,WhenCreated


if ($null -ne $groupInfo.ManagedBy) {
    $DN2ConicalName = Get-AdUser -Filter "distinguishedName -eq ""$($GroupInfo.ManagedBy)""" -Properties SAMAccountName,DistinguishedName,Displayname | Select-Object SAMAccountName,DistinguishedName,Displayname
}
else {
    $DN2ConicalName.Displayname = "No Manager"
}

    for ($ii = 0; $ii -lt $groupList.Count; $ii++) {

        $groupmembers = Get-ADGroupMember -Identity $GroupInfo.CN | Select-Object SAMAccountName,DistinguishedName       

        $membercount = [System.Collections.ArrayList]::new()
        $membercount = $groupmembers.count

        $TableName = $GroupInfo.CN -replace '[^a-zA-Z0-9]', ''

       <#  foreach ($member in $groupmembers) {

            $getDisplayname = [System.Collections.ArrayList]::new()
            $getDisplayname = Get-ADUser -Filter "distinguishedName -eq ""$($groupmembers.DistinguishedName)""" -Properties Displayname | Select-Object DisplayName
            
        }
        $getDisplayname | Export-Excel -Path "$($FileName)" -AutoSize -TableName "$($TableName)" -WorksheetName "$($GroupInfo.CN)" #>
    }

$GroupInfoExport = [PSCustomObject]@{
    'ConicalName' = [String]$GroupInfo.CN
    'DisplayName' = [String]$GroupInfo.DisplayName
    'Description of Group' = [String]$GroupInfo.Info
    'Email Address Used' = [String]$GroupInfo.Mail
    'Managed by' = [String]$DN2ConicalName.DisplayName
    'Created On Date' = $GroupInfo.WhenCreated
    'Number of Members' = [Int]$membercount
}

$GroupInfoExport | Export-Excel -Path "$($FileName)" -AutoSize -TableName "Summary" -WorksheetName "Summary" -Append

}





