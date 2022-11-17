


$ou = Get-ADOrganizationalUnit -filter * -Server DC-TC.vvsd.org | `
 Select-Object DistinguishedName, Name | Sort-Object DistinguishedName | Out-GridView -PassThru -Title "Choose OU to search"

 $computers = get-adcomputer -filter * -SearchBase $OU.DistinguishedName | Select -Property Name

 get-adcomputer -filter * -SearchBase $OU.DistinguishedName | Select -Property Name | foreach { Get-AdmPwdPassword -Computername $_.Name } | Export-Csv -Path C:\Temp\test.csv -NoTypeInformation



 #$forloop = foreach ($computer in $computers) {
  #  $AdmPwd = Get-AdmPwdPassword -Computername $Computer
   # $results +=$AdmPwd
    #}
#$results | Select-Object -Property Computername,Password,ExpirationTimeStamp | Export-Csv C:\Temp\LAPS-$OU.csv
