<#
.SYNOPSIS
   Installs included certificate on client machine for Allscripts installation
 

.NOTES
    Version:        1.0
    Author:         Colin McMorrow
    Created:        9/3/2024
    Organization:   Marco Technologies, LLC
    Purpose:        Install certificate on client machine in both Root and Trusted People stores

#>

# Define the certificate file path.  This should be the name of the .cer file that is wrapped in .intunewin file
$certificatePath = "ODHC-AS-2.ODHC.local.cer"

# Import the certificate to the Trusted Root Certification Authorities store
Import-Certificate -FilePath $certificatePath -CertStoreLocation Cert:\LocalMachine\Root

# Import the certificate to the Trusted People store
Import-Certificate -FilePath $certificatePath -CertStoreLocation Cert:\LocalMachine\TrustedPeople
