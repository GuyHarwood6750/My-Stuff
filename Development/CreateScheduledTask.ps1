<#
Script to create scheduled task in Windows Task Scheduler
#>
$trigger = New-ScheduledTaskTrigger -At 5:15pm -Weekly -DaysOfWeek 'Friday'
#$action = New-ScheduledTaskAction -Execute '"C:\Program files\Microsoft Office\root\Office16\EXCEL.EXE"' -Argument '"z:\userdata\Guy\Barrydale Electricity\Barrydale Electricity.xlsx"'
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-file "C:\Users\Guy\Documents\Powershell\Daily Scripts\Copy-invoices2.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "WM Monthly Debtors" -TaskPath '\guy' -Trigger $trigger -Action $action -RunLevel Highest -Force



