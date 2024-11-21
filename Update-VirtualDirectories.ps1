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

# Process and update each virtual directory type separately
Write-Host "Updating virtual directories for server: $server" -ForegroundColor Cyan

# OWA Virtual Directory
Get-OwaVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/owa"
    $_.ExternalUrl = "https://$domain/owa"
    Set-OwaVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated OWA: InternalUrl and ExternalUrl set to https://$domain/owa" -ForegroundColor Green
}

# EWS Virtual Directory
Get-WebServicesVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/ews/exchange.asmx"
    $_.ExternalUrl = "https://$domain/ews/exchange.asmx"
    Set-WebServicesVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated EWS: InternalUrl and ExternalUrl set to https://$domain/ews/exchange.asmx" -ForegroundColor Green
}

# ActiveSync Virtual Directory
Get-ActiveSyncVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/Microsoft-Server-ActiveSync"
    $_.ExternalUrl = "https://$domain/Microsoft-Server-ActiveSync"
    Set-ActiveSyncVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated ActiveSync: InternalUrl and ExternalUrl set to https://$domain/Microsoft-Server-ActiveSync" -ForegroundColor Green
}

# OAB Virtual Directory
Get-OabVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/oab"
    $_.ExternalUrl = "https://$domain/oab"
    Set-OabVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated OAB: InternalUrl and ExternalUrl set to https://$domain/oab" -ForegroundColor Green
}

# PowerShell Virtual Directory
Get-PowerShellVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/powershell"
    $_.ExternalUrl = "https://$domain/powershell"
    Set-PowerShellVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated PowerShell: InternalUrl and ExternalUrl set to https://$domain/powershell" -ForegroundColor Green
}

# MAPI Virtual Directory
Get-MAPIVirtualDirectory -Server $server | ForEach-Object {
    $_.InternalUrl = "https://$domain/mapi"
    $_.ExternalUrl = "https://$domain/mapi"
    Set-MAPIVirtualDirectory -Identity $_.Identity -InternalUrl $_.InternalUrl -ExternalUrl $_.ExternalUrl
    Write-Host "Updated MAPI: InternalUrl and ExternalUrl set to https://$domain/mapi" -ForegroundColor Green
}

# Outlook Anywhere
Get-OutlookAnywhere -Server $server | ForEach-Object {
    Set-OutlookAnywhere -Identity $_.Identity -InternalHostname $domain -ExternalHostname $domain -InternalClientsRequireSsl $true -ExternalClientsRequireSsl $true -SSLOffloading $false
    Write-Host "Updated Outlook Anywhere: Hostname set to $domain" -ForegroundColor Green
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
Get-OwaVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
Get-WebServicesVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
Get-ActiveSyncVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
Get-OabVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
Get-PowerShellVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
Get-MAPIVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl
