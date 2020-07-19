#Get-NetIPConfiguration -All -AllCompartments -Detailed | Where-Object { $_.NetAdapter.Status -eq 'Up' } | Out-File -FilePath "c:\test\ipconfigMikrotik.txt"

#Get-VpnConnection | Out-File -FilePath "C:\test\ipconfigADSL.txt" -Append

Test-NetConnection 41.75.108.157 -TraceRoute