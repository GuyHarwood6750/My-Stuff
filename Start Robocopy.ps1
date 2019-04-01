function Start-Robocopy
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,
                    Position=0)]
        [string]$Source,

        [Parameter(Mandatory,
                    Position=1)]
        [string]$Destination,

        [string]$IncludeFile = '*.zip',
        
        [string]$RobocopyPath = 'c:\Windows\system32\robocopy.exe',
        
        [string]$RobocopyParameter = '/E /Z /MIR /V /NP /R:3 /W:5 /MT',
        
        [string]$LogFileName = "C:\Logs\Start-Robocopy-$(Get-Date -Format 'yyyyMMddhhmmss').log",

        [switch]$Tee = $false
    )

    Begin
    {
        # Turn on verbose
        $VerbosePreference = 'Continue'

        # Start Log
        # Begin Logging
        Add-Content -Value "Beginning $($MyInvocation.InvocationName) on $($env:COMPUTERNAME) by $env:USERDOMAIN\$env:USERNAME" -Path $LogFileName

        $RobocopyParameter = "$RobocopyParameter /LOG+:$LogFileName"
        if ($Tee) {
            $RobocopyParameter = "$RobocopyParameter /TEE"
            }
        
        Write-Verbose "Robocopy Parameters: $RobocopyParameter"

    }
    Process
    {
        
        # Build robocopy command line
        $RobocopyExecute = "$RobocopyPath $Source $Destination $IncludeFile $RobocopyParameter"       
        Write-Verbose "Executing Robocopy Command Line: $RobocopyExecute"
        Add-Content "Executing Robocopy Command Line: $RobocopyExecute" -Path $LogFileName
        Invoke-Expression $RobocopyExecute
        
    }
    End
    {

    }
}