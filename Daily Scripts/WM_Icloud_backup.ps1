# Backup Affinity files from ICLOUD drive to server
#
Robocopy "C:\Users\Guy\iCloudDrive\Circe Launches" "\\wserver\wmarine\management\advertising" /XO /MIR /copy:dat /log:"c:\userdata\circe launches\logs\icloud $(get-date -f yyyyMMdd-HHmmss).txt"