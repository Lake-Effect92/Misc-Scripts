<# 
.SYNOPSIS
    Retrieves the internal and external URLs of all Exchange virtual directories across all servers in the organization.
.DESCRIPTION
    This script dynamically discovers all Exchange servers in the organization and retrieves the internal and external URLs 
    for OWA, EWS, ActiveSync, OAB, and PowerShell virtual directories.

.AUTHOR
    Colin McMorrow
    mcmorrow.colin@gmail.com

.VERSION
    1.0

.NOTES
    - Requires the Exchange Management tools or Exchange Management Shell.
    - Ensure the user running the script has sufficient permissions to execute Exchange cmdlets.
    - Outputs URLs to the console in a structured format.

#>

# Get all Exchange servers
$servers = Get-ExchangeServer | Select-Object -ExpandProperty Name

# Loop through each server and retrieve virtual directory settings
foreach ($server in $servers) {
    Write-Host "Retrieving virtual directory URLs for server: $server" -ForegroundColor Green
    
    Get-ClientAccessServer -Identity $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "Client Access Server URLs:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }

    # Get OWA virtual directory
    Get-OwaVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "OWA Virtual Directory:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }
    
    # Get EWS virtual directory
    Get-WebServicesVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "EWS Virtual Directory:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }

    # Get ActiveSync virtual directory
    Get-ActiveSyncVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "ActiveSync Virtual Directory:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }

    # Get OAB virtual directory
    Get-OabVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "OAB Virtual Directory:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }

    # Get PowerShell virtual directory
    Get-PowerShellVirtualDirectory -Server $server | Select-Object Name, InternalUrl, ExternalUrl | ForEach-Object {
        Write-Host "PowerShell Virtual Directory:"
        Write-Host "Name: $($_.Name)"
        Write-Host "InternalUrl: $($_.InternalUrl)"
        Write-Host "ExternalUrl: $($_.ExternalUrl)"
        Write-Host "---------------------------"
    }
}
