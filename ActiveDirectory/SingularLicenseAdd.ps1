



$inFileName = C:\temp\DefenderBatch.csv
$outFileName="C:\temp\Users2License-Done.CSV"
$accountSkuId = Read-Host "Please Enter the License SKU for the product you want to add:"

 $users=Import-Csv $inFileName
 #$licenseOptions=New-MsolLicenseOptions -AccountSkuId $accountSkuId -DisabledPlans $planList
 ForEach ($user in $users)
     {
         $user.Userprincipalname
         $upn=$user.UserPrincipalName
         Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $accountSkuId -ErrorAction SilentlyContinue
         Start-Sleep -Seconds 2
         #Set-MsolUserLicense -UserPrincipalName $upn -LicenseOptions $licenseOptions -ErrorAction SilentlyContinue
     }