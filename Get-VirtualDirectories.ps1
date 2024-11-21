# Get all Exchange servers
$servers = Get-ExchangeServer | Select-Object -ExpandProperty Name

# Loop through each server and retrieve virtual directory settings
foreach ($server in $servers) {
    Write-Host "Retrieving virtual directory URLs for server: $server" -ForegroundColor Green
    
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
