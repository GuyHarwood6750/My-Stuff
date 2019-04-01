$x = 100
$vpnok = Get-VpnConnection
$locallconn = Test-NetConnection -ComputerName "www.circelaunches.co.za" -InformationLevel Detailed

if($vpnok.ConnectionStatus -eq 'connected'){
   write-host("VPN connection good")
} elseif($locallconn.PingSucceeded -eq 'True'){
   write-host("Local Network connection!")
} elseif($x -eq 30){
   write-host("Value of X is 30")
} else {
   write-host("No network!")
}