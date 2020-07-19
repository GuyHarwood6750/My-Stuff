$count = 20
$buffer = 32

#ping kiosk-pc.wmarine.local -n $count -l $buffer | Out-File -FilePath "c:\test\ping\ping_Kiosk $(get-date -f yyyyMMdd-HHmm).txt"
#ping wserver.wmarine.local -n $count -l $buffer | Out-File -FilePath "c:\test\ping\ping_WServer_IP2 $(get-date -f yyyyMMdd-HHmm).txt"
#ping 10.20.30.1 -n $count -l $buffer | Out-File -FilePath "c:\test\ping\ping_Router_IP1 $(get-date -f yyyyMMdd-HHmm).txt"
#ping acer7730.wmarine.local -n $count -l $buffer | Out-File -FilePath "c:\test\ping\ping_Julia $(get-date -f yyyyMMdd-HHmm).txt"
#ping 41.75.108.157 -n 5
#ping guyharwood.wirelessweb.co.za -n 5
ping www.circelaunches.co.za -n 5