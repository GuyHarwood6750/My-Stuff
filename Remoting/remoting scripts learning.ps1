#$cred = Get-Credential
$machine1 = "Guy-7750G"
$sess1 = New-PSSession -ComputerName $machine1
#$hosts = 'guy-7750G,192.168.1.35'
#Set-WSManInstance -ResourceURI winrm/config/client -ValueSet @{TrustedHosts=$hosts}
#Get-WSManInstance -ResourceURI winrm/config/client

#[System.Net.Dns]::GetHostEntry($machine)
#Invoke-Command -Session $sess1 -ScriptBlock{Get-Service winr*}
Remove-PSSession $sess1
#Invoke-Command -ComputerName $machine1 -ScriptBlock {Get-ChildItem -Path 'c:\retriever trials' -Recurse}
#Enter-PSSession $sess1
#Get-PSSession -ComputerName $machine1