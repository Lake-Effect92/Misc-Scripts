<#	
	.NOTES
	===========================================================================
	 Created on:   	6/17/2024
	 Created by:   	Colin McMorrow
	 Filename:     	Remove-BitlockerEncryption.ps1
	===========================================================================
	.DESCRIPTION
		Script will check Bitlocker status of volume C: and decrypt if the EncryptionMethod doesn't equal "XtsAes256".
#>


# Function to get encryption statu
function Get-BitLockerStatus {
    param (
        [string]$DriveLetter = "C:"
    )
    
    $bitLockerStatus = Get-BitlockerVolume -MountPoint $DriveLetter
    return $bitLockerStatus
}

# Function to decrypt the drive
function Decrypt-Drive {
    param (
        [string]$DriveLetter = "C:"
    )
    
    Write-Output "Decrypting drive $DriveLetter..."
    Disable-BitLocker -MountPoint $DriveLetter
    Write-Output "Drive $DriveLetter decryption initiated."
}

# Main script

$driveLetter = "C:"
$bitLockerStatus = Get-BitLockerStatus -DriveLetter $driveLetter

if ($bitLockerStatus.ProtectionStatus -ne "On") {
    Write-Output "..."
} else {
    $encryptionMethod = $bitLockerStatus.EncryptionMethod
    Write-Output "Current encryption method: $encryptionMethod"

    if ($encryptionMethod -ne "XtsAes256") {
        Write-Output "Encryption method is not XtsAes256. Proceeding to decrypt the drive."
        Decrypt-Drive -DriveLetter $driveLetter -whatif
    } else {
        try{
		$BLV = Get-BitLockerVolume -MountPoint $env:SystemDrive
        $KeyProtectorID=""
        foreach($keyProtector in $BLV.KeyProtector){
            if($keyProtector.KeyProtectorType -eq "RecoveryPassword"){
                $KeyProtectorID=$keyProtector.KeyProtectorId
                break;
            }
        }

       $result = BackupToAAD-BitLockerKeyProtector -MountPoint "$($env:SystemDrive)" -KeyProtectorId $KeyProtectorID
return $true
}
catch{
     return $false
    }
}
}