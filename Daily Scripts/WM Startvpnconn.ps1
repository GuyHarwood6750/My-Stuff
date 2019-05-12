$vpnname = "WMarine VPN"
$credentials = Get-StoredCredential -AsCredentialObject -Target "VPN"
$vpnusername = $credentials.UserName
$vpnpassword = $credentials.Password
     $vpn = Get-VpnConnection | Where-Object {$_.Name -eq $vpnname}
        if ($vpn.ConnectionStatus -eq "Disconnected")
          {
             $cmd = $env:WINDIR + "\System32\rasdial.exe"
             $expression = "$cmd ""$vpnname"" $vpnusername $vpnpassword"
             Invoke-Expression -Command $expression 
          }