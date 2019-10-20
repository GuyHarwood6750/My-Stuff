<#
Script to create scheduled task in Windows Task Scheduler
#>
#$trigger = New-ScheduledTaskTrigger -At 21:45pm -Weekly -DaysOfWeek 'Monday', 'Wednesday', 'Friday'
$trigger = New-ScheduledTaskTrigger -At 22:30pm -Daily

#$action = New-ScheduledTaskAction -Execute '"C:\Program files\Microsoft Office\root\Office16\EXCEL.EXE"' -Argument '"z:\userdata\Guy\Barrydale Electricity\Barrydale Electricity.xlsx"'
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-WindowStyle Hidden -file "C:\Users\Guy\Documents\Powershell\Daily Scripts\R62S Dropbox Backup.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "R62S Dropbox Backup" -TaskPath '\guy' -Trigger $trigger -Action $action -RunLevel Highest -Force



