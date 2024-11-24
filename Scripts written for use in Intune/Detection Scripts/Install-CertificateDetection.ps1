# Define the certificate thumbprint to check for
$certThumbprint = "98b6adba89305b8b0c11ce181ea20e122c7eeb3e"

# Function to check if a certificate with a specific thumbprint exists in a specific store
function Test-CertificatePresenceByThumbprint {
    param (
        [string]$CertStoreLocation,
        [string]$Thumbprint
    )

    # Get the certificates in the specified store
    $certificates = Get-ChildItem -Path $CertStoreLocation

    # Check if the certificate with the specified thumbprint exists
    foreach ($cert in $certificates) {
        if ($cert.Thumbprint -eq $Thumbprint) {
            return $true
        }
    }

    return $false
}

# Check for the certificate in the Trusted Root Certification Authorities store
$rootCertExists = Test-CertificatePresenceByThumbprint -CertStoreLocation "Cert:\LocalMachine\Root" -Thumbprint $certThumbprint

# Check for the certificate in the Trusted People store
$trustedPeopleCertExists = Test-CertificatePresenceByThumbprint -CertStoreLocation "Cert:\LocalMachine\TrustedPeople" -Thumbprint $certThumbprint

# Determine if both certificates exist
if ($rootCertExists -and $trustedPeopleCertExists) {
    # Certificates are present in both stores, exit with code 0 (success) and STDOUT
    Write-Host "Certificates are Installed"
	exit 0
} else {
    # Certificates are missing in one or both stores, exit with code 1 and no STDOUT
    exit 1
}