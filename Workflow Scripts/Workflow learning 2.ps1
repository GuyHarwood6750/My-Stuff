Workflow Get-Eventlogdata {
    Parallel
{ Get-Eventlog -LogName application -Newest 1
        Get-Eventlog -LogName system -Newest 1
        Get-Eventlog -LogName "Windows Powershell" -Newest 1
}}
