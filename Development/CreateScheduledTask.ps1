<#
Script to create scheduled task in Windows Task Scheduler
#>
$trigger = New-ScheduledTaskTrigger -At 8:15am -Daily
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument '-file "C:\Users\Guy\Documents\Powershell\Daily Scripts\Run SQL Backups.ps1"' `
    -WorkingDirectory 'C:\Users\Guy\Documents\Powershell\Daily Scripts'
Register-ScheduledTask -TaskName "WM SQL Backup" -TaskPath '\guy' -Trigger $trigger -Action $action -RunLevel Highest -Force