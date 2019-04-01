Import-Csv -Path "C:\test\Reference examples\testaccounts.csv" | Get-Member

dir $PSHOME -Filter *.format.ps1xml

Get-WmiObject -Class win32_Operatingsystem | Get-Member

Get-Service | sort status | ft -GroupBy Status

Get-EventLog System -Newest 5 | ft source,message -wrap -auto

Get-WmiObject win32_logicaldisk -Filter "DeviceID='C:'"| fl *

