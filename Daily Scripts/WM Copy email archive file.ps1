<#
Backup Customer Folders
#>
#Robocopy "\\wserver\wmarine\Customers" "D:\Circe Launches Backups\Customers" /XO /MIR /copy:dat /xf 'tax invoice*' 'statemen*' 'credi*' /log:"c:\userdata\circe launches\logs\CUSTOMERS $(get-date -f yyyyMMdd-HHmmss).txt"
#
Robocopy "D:\Circe Launches Backups\Mail Backups" "\\wserver\WMarine\managment\email archives\info gmail" "info@circelaunches.co.za_-_Guy_20200807_165533_(Full).ost" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\infoemailarchive $(get-date -f yyyyMMdd-HHmmss).txt"