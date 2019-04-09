<#
Script to create scheduled task in Windows Task Scheduler
#>
#$trigger = New-ScheduledTaskTrigger -At 21:15pm -Weekly -DaysOfWeek 'Friday'
$trigger = New-ScheduledTaskTrigger -At 21:15pm -Daily

#$action = New-ScheduledTaskAction -Execute '"C:\Program files\Microsoft Office\root\Office16\EXCEL.EXE"' -Argument '"z:\userdata\Guy\Barrydale Electricity\Barrydale Electricity.xlsx"'
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-file "C:\Users\Guy\Documents\Powershell\Daily Scripts\R62S Compress Pastel data F2020.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "R62S Backup Pastel F2020" -TaskPath '\guy' -Trigger $trigger -Action $action -RunLevel Highest -Force



