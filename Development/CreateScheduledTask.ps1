<#
Script to create scheduled task in Windows Task Scheduler
#>
#$trigger = New-ScheduledTaskTrigger -At 21:45pm -Weekly -DaysOfWeek 'Monday', 'Wednesday', 'Friday'
$trigger = New-ScheduledTaskTrigger -At 02:35pm -Daily

#$action = New-ScheduledTaskAction -Execute '"C:\Program files\Microsoft Office\root\Office16\EXCEL.EXE"' -Argument '"z:\userdata\Guy\Barrydale Electricity\Barrydale Electricity.xlsx"'
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-file "C:\Users\Guy\Documents\Powershell\Daily Scripts\WM NoGuideName.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "WM NoGuideName" -TaskPath '\guy' -Trigger $trigger -Action $action -RunLevel Highest -Force



