$vpnok = Get-VpnConnection
if ($vpnok.ConnectionStatus -eq 'connected')
    {
        Write-Host "Connected" -ForegroundColor Green
     }
    else 
    {
       Write-Host "Not connected" -ForegroundColor Red
    }