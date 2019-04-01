Workflow Get-Compinfo {
    Get-NetAdapter
    Get-Disk | Select-Object FriendlyName, SerialNumber, Model, Size
    Get-Volume
}