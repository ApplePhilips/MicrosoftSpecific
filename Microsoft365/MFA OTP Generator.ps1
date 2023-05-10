<# --------------------------------------------
 File:		MFA OTP Generator.ps1
 Author:    Nathan Carpenter, Network Analyst
 Purpose:	Assist in the bulk activation of MFA token
            without needing the token
 Requires:  .Net Framework 
            
            
             
 VVSD History:
 1.0.0		Created - 09/13/2021 - NC
 1.2.0      Updated to allow the generation 
            of TOTP codes from a list every 
            n seconds - 06.01.22 - NC 
             
    
-------------------------------------------- #>



function LoadInputFile  ($InitialDirectory) {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}


function Convert-Base32ToByte {
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Base32
    )

    # RFC 4648 Base32 alphabet
    $rfc4648 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'

    $bits = ''

    # Convert each Base32 character to the binary value between starting at
    # 00000 for A and ending with 11111 for 7.
    foreach ($char in $Base32.ToUpper().ToCharArray()) {
        $bits += [Convert]::ToString($rfc4648.IndexOf($char), 2).PadLeft(5, '0')
    }

    # Convert 8 bit chunks to bytes, ignore the last bits.
    for ($i = 0; $i -le ($bits.Length - 8); $i += 8) {
        [Byte] [Convert]::ToInt32($bits.Substring($i, 8), 2)
    }
}

function Get-TimeBasedOneTimePassword {
    [CmdletBinding()]
    [Alias('Get-TOTP')]
    param
    (
        # Base 32 formatted shared secret (RFC 4648).
        [Parameter(Mandatory = $true)]
        [System.String]
        $SharedSecret,

        # The date and time for the target calculation, default is now (UTC).
        [Parameter(Mandatory = $false)]
        [System.DateTime]
        $Timestamp = (Get-Date).ToUniversalTime(),

        # Token length of the one-time password, default is 6 characters.
        [Parameter(Mandatory = $false)]
        [System.Int32]
        $Length = 6,

        # The hash method to calculate the TOTP, default is HMAC SHA-1.
        [Parameter(Mandatory = $false)]
        [System.Security.Cryptography.KeyedHashAlgorithm]
        $KeyedHashAlgorithm = (New-Object -TypeName 'System.Security.Cryptography.HMACSHA1'),

        # Baseline time to start counting the steps (T0), default is Unix epoch.
        [Parameter(Mandatory = $false)]
        [System.DateTime]
        $Baseline = '1970-01-01 00:00:00',

        # Interval for the steps in seconds (TI), default is 30 seconds.
        [Parameter(Mandatory = $false)]
        [System.Int32]
        $Interval = 60
    )

    # Generate the number of intervals between T0 and the timestamp (now) and
    # convert it to a byte array with the help of Int64 and the bit converter.
    $numberOfSeconds = ($Timestamp - $Baseline).TotalSeconds
    $numberOfIntervals = [Convert]::ToInt64([Math]::Floor($numberOfSeconds / $Interval))
    $byteArrayInterval = [System.BitConverter]::GetBytes($numberOfIntervals)
    [Array]::Reverse($byteArrayInterval)

    # Use the shared secret as a key to convert the number of intervals to a
    # hash value.
    $KeyedHashAlgorithm.Key = Convert-Base32ToByte -Base32 $SharedSecret
    $hash = $KeyedHashAlgorithm.ComputeHash($byteArrayInterval)

    # Calculate offset, binary and otp according to RFC 6238 page 13.
    $offset = $hash[($hash.Length - 1)] -band 0xf
    $binary = (($hash[$offset + 0] -band '0x7f') -shl 24) -bor
    (($hash[$offset + 1] -band '0xff') -shl 16) -bor
    (($hash[$offset + 2] -band '0xff') -shl 8) -bor
    (($hash[$offset + 3] -band '0xff'))
    $otpInt = $binary % ([Math]::Pow(10, $Length))
    $otpStr = $otpInt.ToString().PadLeft($Length, '0')

    return $otpStr
}

function TOTPGeneration ($UPN,$search,$inputfile) {
    if ($null -eq $inputfile) {
        $inputfile = LoadInputFile -InitialDirectory "C:\"
    }
    $workingfile = Import-Csv $inputfile

    Write-Host "Entering TOTP generator..." -ForegroundColor Blue
    #Write-Host ""
    $search = Read-Host "Would you like to search for individual user(Y/n):"

    if ($search -eq "y") {
        do { 
        $SearchByUPN = Read-Host "Enter the full email of the user to get code:"

        [string] $usersecretKey = $workingfile | Where-Object 'UPN' -eq $SearchByUPN | Select-Object 'Secret Key' -ExpandProperty 'Secret Key'
        $totpkey = Get-TimeBasedOneTimePassword -SharedSecret $usersecretKey
        Write-Host $($SearchByUPN) -ForegroundColor Magenta -NoNewline
        Write-Host "'s TOTP code is: " -NoNewline
        Write-Host $totpkey -BackgroundColor Black -ForegroundColor Green
        Start-Sleep -Seconds 2
        $continue = Read-Host "Search for Another User(Y/n)"

        } while ($continue -eq 'Y')
    }
    elseif ($search -eq "n") {

        while ($true) {

            for ($i = 0; $i -lt $workingfile.count; $i++) {
            [String] $usersecretKey = $workingfile[$i] | Select-Object 'Secret Key' -ExpandProperty 'Secret Key'
            [String] $UPNDisplay = $workingFile[$i] | Select-Object 'UPN' -ExpandProperty 'UPN'
            $totpkey = Get-TimeBasedOneTimePassword -SharedSecret $usersecretKey
            Write-Host "$($UPNDisplay)" -ForegroundColor Magenta -NoNewline
            Write-Host "'s TOTP code is: " -NoNewline
            Write-Host $totpkey -BackgroundColor Black -ForegroundColor Green
            }
            Write-Host "Use CRTL + C to stop the loop..." -ForegroundColor Red

            Start-Sleep -Seconds 60
        }
    }
    else {
        Write-Host "Thank you for using the TOTP generator. A huge Breach in security for MFA..." -ForegroundColor Blue
    }
    Write-Host "Thank you for using the TOTP generator. A huge Breach in security for MFA..." -ForegroundColor Blue
}

TOTPGeneration