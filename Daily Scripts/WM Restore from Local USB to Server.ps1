<#
Restore Scanned vouchers from local USB to Server 
#>
Robocopy "D:\Circe Launches Backups\Kiosk\Invoices\F2020\06Scanned August Invoices 2019" "\\wserver\wmarine\kiosk\invoices\F2020\06Scanned August Invoices 2019"  /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\restinv $(get-date -f yyyyMMdd-HHmmss).txt"
Robocopy "D:\Circe Launches Backups\Kiosk\Invoices\F2020\07Scanned September Invoices 2019" "\\wserver\wmarine\kiosk\invoices\F2020\07Scanned September Invoices 2019"  /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\restinv $(get-date -f yyyyMMdd-HHmmss).txt"
Robocopy "D:\Circe Launches Backups\Kiosk\Invoices\F2020\08Scanned October Invoices 2019" "\\wserver\wmarine\kiosk\invoices\F2020\08Scanned October Invoices 2019"  /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\restinv $(get-date -f yyyyMMdd-HHmmss).txt"
