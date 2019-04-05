Get-VpnConnection -name 'WMarine VPN2' | out-file 'c:\test\vpndetails2.txt'
Set-VpnConnection -Name 'WMarine VPN2' -ServerAddress '41.75.108.157'
Remove-VpnConnection -Name 'WMarine VPN2'
#Add-VpnConnection -Name 'WMarine VPN2' -ServerAddress 'remote.circelaunches.co.za' -TunnelType 'sstp' -AuthenticationMethod 'mschapv2'