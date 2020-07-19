<#
Backup Customer Folders
#>
Robocopy "\\wserver\wmarine\Customers" "D:\Circe Launches Backups\Customers" /XO /MIR /copy:dat /xf 'tax invoice*' 'statemen*' 'credi*' /log:"c:\userdata\circe launches\logs\CUSTOMERS $(get-date -f yyyyMMdd-HHmmss).txt"