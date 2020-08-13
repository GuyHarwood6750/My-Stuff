<#
Script to create scheduled task in Windows Task Scheduler
#>
$trigger = New-ScheduledTaskTrigger -At 01:15pm -Weekly -DaysOfWeek 'Sunday'
#$trigger = New-ScheduledTaskTrigger -At 21:45pm -Weekly -DaysOfWeek 'Monday', 'Wednesday', 'Friday'
#$trigger = New-ScheduledTaskTrigger -At 00:15am -Daily

#$action = New-ScheduledTaskAction -Execute '"C:\Program files\Microsoft Office\root\Office16\EXCEL.EXE"' -Argument '"c:\userdata\Guy\Health\Readings.xlsx"'
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-WindowStyle Hidden -file "C:\Users\Guy\Documents\Powershell\Daily Scripts\WM_icloud backup.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "WM ICloud backup" -TaskPath '\guy' -Trigger $trigger -Action $action -Force



