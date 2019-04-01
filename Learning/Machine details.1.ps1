#Technique using hash table (ordered) (Page 346, Powershell in Depth)
$comp = $Env:COMPUTERNAME
$os = Get-WmiObject -Class win32_operatingsystem -ComputerName $Comp
$cs = Get-WmiObject -Class win32_computersystem -ComputerName $comp
$bios = Get-WmiObject -Class win32_bios -ComputerName $comp
$proc = Get-WmiObject -Class win32_Processor -ComputerName $comp | Select-Object -First 1
$props = [ordered]@{
    OSversion        = $os.version
    Model            = $cs.Model
    Manufacturer     = $cs.Manufacturer
    BIOSSerial       = $bios.serialnumber
    ComputerName     = $os.CSName
    OSArchitecture   = $os.OSArchitecture
    ProcArchitecture = $proc.addresswidth
}
$obj = New-Object -TypeName PSObject -Property $props

Write-Output $obj | Format-List

$obj | Get-Member