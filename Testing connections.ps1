$vpnok = Get-VpnConnection

if ($vpnok.connectionstatus -eq 'connected')
{
Write-Host 'OK....' -BackgroundColor White -ForegroundColor DarkGreen
}

elseif ($locallconn = Test-NetConnection 'wserver') {
   if ($locallconn.pingsucceeded -eq 'true') {
    write-host 'Local conn OK'
   }
    #write-host 'Local conn NOT OK' -ForegroundColor DarkRed -BackgroundColor White
    }
else
{
write-host 'Nothing OK'
}

