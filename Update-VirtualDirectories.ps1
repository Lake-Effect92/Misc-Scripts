<# 
.SYNOPSIS
    Updates Exchange virtual directory URLs and ensures consistency with the provided domain name.
.DESCRIPTION
    This script prompts the user for an external DNS name, validates it, and allows updating all Exchange virtual 
    directories to use the specified domain for both internal and external URLs. It also verifies the Autodiscover 
    URI and displays the final configuration.
.AUTHOR
    Colin McMorrow
    mcmorrow.colin@gmail.com
.VERSION
    1.0
.NOTES
    - Run this script in the Exchange Management Shell or with Exchange Management tools installed.
    - Requires sufficient permissions to modify Exchange virtual directories.
#>

# Prompt for the external DNS name
$domain = Read-Host "Enter the external DNS name for your Exchange (e.g., mail.domain.com)"
if (-not $domain -or $domain -notmatch '^[a-zA-Z0-9.-]+$') {
    Write-Host "Invalid domain entered. Please run the script again and provide a valid domain." -ForegroundColor Red
    exit
}
Write-Host "You entered: $domain. Please verify it is correct." -ForegroundColor Yellow
if (-not (Read-Host "Is this correct? (Y/N)").ToUpper().StartsWith("Y")) {
    Write-Host "Operation cancelled by user." -ForegroundColor Red
    exit
}

# Confirm the update
$confirmUpdate = Read-Host "Would you like to set all virtual directories to use this domain? (Y/N)"
if (-not $confirmUpdate.ToUpper().StartsWith("Y")) {
    Write-Host "Operation cancelled by user." -ForegroundColor Red
    exit
}

# Prompt for the server name
$server = Read-Host "Enter the Exchange server name"
if (-not (Get-ExchangeServer -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq $server})) {
    Write-Host "The server '$server' is not an Exchange server or cannot be found." -ForegroundColor Red
    exit
}

# Collect virtual directories individually
$virtualDirectories = @(
    Get-OwaVirtualDirectory -Server $server,
    Get-WebServicesVirtualDirectory -Server $server,
    Get-ActiveSyncVirtualDirectory -Server $server,
    Get-OabVirtualDirectory -Server $server,
    Get-PowerShellVirtualDirectory -Server $server,
    Get-OutlookAnywhere -Server $server,
    Get-MAPIVirtualDirectory -Server $server
)

foreach ($vd in $virtualDirectories) {
    if ($vd) {
        $vd.InternalUrl = "https://$domain$($vd.InternalUrl.PathAndQuery)"
        $vd.ExternalUrl = "https://$domain$($vd.ExternalUrl.PathAndQuery)"
        Set-Object -InputObject $vd
        Write-Host "$($vd.Name) updated: InternalUrl and ExternalUrl set to https://$domain$($vd.InternalUrl.PathAndQuery)" -ForegroundColor Green
    }
}

# Verify and update AutodiscoverInternalURI
$cas = Get-ClientAccessServer -Identity $server
if ($cas.AutodiscoverInternalURI -ne "https://$domain/autodiscover/autodiscover.xml") {
    Set-ClientAccessServer -Identity $server -AutodiscoverServiceInternalURI "https://$domain/autodiscover/autodiscover.xml"
    Write-Host "AutodiscoverInternalURI updated to https://$domain/autodiscover/autodiscover.xml" -ForegroundColor Green
} else {
    Write-Host "AutodiscoverInternalURI is already set correctly." -ForegroundColor Green
}

# Print final virtual directory URLs
Write-Host "Final Virtual Directory URLs for server: $server" -ForegroundColor Cyan
foreach ($vd in $virtualDirectories) {
    if ($vd) {
        Write-Host "$($vd.Name): InternalUrl = $($vd.InternalUrl), ExternalUrl = $($vd.ExternalUrl)"
    }
}
