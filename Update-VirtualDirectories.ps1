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

# Prompt for the server name
$server = Read-Host "Enter the Exchange server name"
if (-not (Get-ExchangeServer -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq $server})) {
    Write-Host "The server '$server' is not an Exchange server or cannot be found." -ForegroundColor Red
    exit
}

# Helper function to display and confirm changes
function Show-Change {
    param (
        [string]$VirtualDirectoryName,
        [string]$OldInternalUrl,
        [string]$OldExternalUrl,
        [string]$NewInternalUrl,
        [string]$NewExternalUrl
    )
    
    Write-Host "Updating $VirtualDirectoryName" -ForegroundColor Cyan
    Write-Host "Old Internal URL: $OldInternalUrl" -ForegroundColor Yellow
    Write-Host "New Internal URL: $NewInternalUrl" -ForegroundColor Green
    Write-Host "Old External URL: $OldExternalUrl" -ForegroundColor Yellow
    Write-Host "New External URL: $NewExternalUrl" -ForegroundColor Green
    return (Read-Host "Proceed with this change? (Y/N)").ToUpper().StartsWith("Y")
}

# Process and update each virtual directory type
Write-Host "Processing virtual directories for server: $server" -ForegroundColor Cyan

# OWA Virtual Directory
Get-OwaVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/owa"
    $newExternalUrl = "https://$domain/owa"
    if (Show-Change "OWA Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-OwaVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "OWA Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "OWA Virtual Directory update skipped." -ForegroundColor Red
    }
}

# EWS Virtual Directory
Get-WebServicesVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/ews/exchange.asmx"
    $newExternalUrl = "https://$domain/ews/exchange.asmx"
    if (Show-Change "EWS Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-WebServicesVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "EWS Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "EWS Virtual Directory update skipped." -ForegroundColor Red
    }
}

# ActiveSync Virtual Directory
Get-ActiveSyncVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/Microsoft-Server-ActiveSync"
    $newExternalUrl = "https://$domain/Microsoft-Server-ActiveSync"
    if (Show-Change "ActiveSync Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-ActiveSyncVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "ActiveSync Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "ActiveSync Virtual Directory update skipped." -ForegroundColor Red
    }
}

# OAB Virtual Directory
Get-OabVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/oab"
    $newExternalUrl = "https://$domain/oab"
    if (Show-Change "OAB Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-OabVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "OAB Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "OAB Virtual Directory update skipped." -ForegroundColor Red
    }
}

# PowerShell Virtual Directory
Get-PowerShellVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/powershell"
    $newExternalUrl = "https://$domain/powershell"
    if (Show-Change "PowerShell Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-PowerShellVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "PowerShell Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "PowerShell Virtual Directory update skipped." -ForegroundColor Red
    }
}

# MAPI Virtual Directory
Get-MAPIVirtualDirectory -Server $server | ForEach-Object {
    $newInternalUrl = "https://$domain/mapi"
    $newExternalUrl = "https://$domain/mapi"
    if (Show-Change "MAPI Virtual Directory" $_.InternalUrl $_.ExternalUrl $newInternalUrl $newExternalUrl) {
        Set-MAPIVirtualDirectory -Identity $_.Identity -InternalUrl $newInternalUrl -ExternalUrl $newExternalUrl
        Write-Host "MAPI Virtual Directory updated successfully." -ForegroundColor Green
    } else {
        Write-Host "MAPI Virtual Directory update skipped." -ForegroundColor Red
    }
}
