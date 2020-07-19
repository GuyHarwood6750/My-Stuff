<#
Backup Scanned vouchers 
#>
#Robocopy "\\wserver\wmarine\kiosk\invoices\01Scanned March Vouchers Invoices 2020" "D:\Circe Launches Backups\Kiosk\Invoices\F2021\01Scanned March Vouchers Invoices 2020" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpinv $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup Scanned Schedules 
#>
#Robocopy "\\wserver\wmarine\kiosk\Schedules\01Scanned March Schedules 2020" "D:\Circe Launches Backups\Kiosk\Schedules\F2021\01Scanned March Schedules 2020" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpschd $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup KIOSK folder (Shutdown period)
#>
Robocopy "\\wserver\wmarine\kiosk" "D:\Circe Launches Backups\Kiosk" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpkiosk $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup Management folder
#>
Robocopy "\\wserver\wmarine\Management" "D:\Circe Launches Backups\Management" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpmngmt $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup Finance folder
#>
Robocopy "\\wserver\wmarine\Finance" "D:\Circe Launches Backups\Finance" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpfinance $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup Photographs folder
#>
Robocopy "\\wserver\wmarine\Photographs" "D:\Circe Launches Backups\Photographs" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpphotos $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup SAMSA folder
#>
Robocopy "\\wserver\wmarine\samsa" "D:\Circe Launches Backups\samsa" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\samsa $(get-date -f yyyyMMdd-HHmmss).txt"
<#
Backup Logos folder
#>
Robocopy "\\wserver\wmarine\logos" "D:\Circe Launches Backups\logos" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\logos $(get-date -f yyyyMMdd-HHmmss).txt"