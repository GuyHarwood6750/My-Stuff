Get-VpnConnection -name 'WMarine VPN2' | out-file 'c:\test\vpndetails2.txt'
Set-VpnConnection -Name 'WMarine VPN2' -ServerAddress '154.126.211.85'
#Add-VpnConnection -Name 'WMarine VPN2' -ServerAddress 'remote.circelaunches.co.za' -TunnelType 'sstp' -AuthenticationMethod 'mschapv2'