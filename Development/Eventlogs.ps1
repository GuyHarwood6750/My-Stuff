New-EventLog -LogName MyPowerShell -Source "HROSS"
Write-EventLog -LogName MyPowerShell -Source "HROSS" -EntryType Information -EventId 1 -Message "HROSS script completed"