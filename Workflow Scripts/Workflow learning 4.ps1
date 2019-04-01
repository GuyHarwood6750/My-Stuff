Workflow get-winfeatures {
    parallel {
        Get-Service -Name WinRM
        sequence {
            InlineScript {$env:COMPUTERNAME}
            Get-Date
            $PSVersionTable.PSVersion
        }
    }
}