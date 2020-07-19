
#Robocopy "\\wserver\wmarine\kiosk\translations" "D:\Circe Launches Backups\kiosk\translations" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkptranslations.txt"

<#
Backup Scanned Schedules 
#>
#Robocopy "\\wserver\wmarine\kiosk\Schedules\F2020\10Scanned December Schedules 2019" "D:\Circe Launches Backups\Kiosk\Schedules\F2020\10Scanned December Schedules 2019" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpschd.txt"

#Robocopy "\\wserver\wmarine\Management" "D:\Circe Launches Backups\Management" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkpmgmt.txt"

#Robocopy "C:\Pastel18\CIRCE20A" "\\wserver\wmarine\Finance\Pastel Offsite Backup\Pastel Backups\2020\CIRCE20A"/XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\bkppastel.txt"

<#
Backup Calypso Signs 
#>
#Robocopy "\\wserver\wmarine\kiosk\Calypso Signs" "D:\Circe Launches Backups\Kiosk\Calypso Signs" /XO /MIR /copy:dat /log:"C:\test\bkpother.txt"

<#
MOVE FOLDERS (/MIR) & FILES
Commnand below will MOVE directory and files (Specify DESTINATION folder name if you wish to preserve the SOURCE folder name)
#>
#Robocopy "\\wserver\wmarine\kiosk\invoices\10Scanned December Invoices 2019" "\\wserver\wmarine\Kiosk\Invoices\F2020\10Scanned December Invoices 2019" /MIR /MOVE /log:"C:\test\moveinvoices.txt"
#
#Robocopy "\\wserver\wmarine\kiosk\schedules\10Scanned December Schedules 2019" "\\wserver\wmarine\Kiosk\schedules\F2020\10Scanned December Schedules 2019" /MIR /MOVE /log:"C:\test\moveschedules.txt"
#
<#
MOVE ONLY FILES (DO NOT USE /MIR)
#>
#Robocopy "C:\Userdata\Circe Launches\InvWM" "C:\Userdata\Circe Launches\InvWM\Completed" /MOVE /log:"C:\userdata\Circe launches\invwm\logs\log.txt"
#
Robocopy "C:\Userdata\Circe Launches\Logs" "C:\Userdata\Circe Launches\Logs\OLD" /MOVE /log:"C:\test\movelogs.txt"
#
#Robocopy "\\wserver\wmarine\customers\_Prepaid" "D:\Circe Launches Backups\Customers\_Prepaid" /XO /MIR /copy:dat /xf 'tax invoice*' 'statemen*' 'credi*' /log:"C:\test\prepaiddebtors.txt"
#