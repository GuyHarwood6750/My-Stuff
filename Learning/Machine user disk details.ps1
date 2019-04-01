#Technique using hash table
$comp = $Env:COMPUTERNAME
$os = Get-WmiObject -Class win32_operatingsystem -ComputerName $Comp
$cs = Get-WmiObject -Class win32_computersystem -ComputerName $comp
$bios = Get-WmiObject -Class win32_bios -ComputerName $comp
$proc = Get-WmiObject -Class win32_Processor -ComputerName $comp | Select-Object -First 1
$disks = Get-WmiObject -Class win32_LogicalDisk -Filter "drivetype=3"
$users = Get-WmiObject -Class win32_useraccount
$diskObjs = @()
foreach ($disk in $disks) {
    $props = @{
        Drive = $disk.DeviceID
        Space = $disk.Size
        Freespace = $disk.Freespace}
     $diskobj = New-Object -TypeName PSObject -Property $props
     $diskObjs += $diskobj   
    }
$userobjs = @()    
foreach ($user in $users) {
    $props = @{
        Username = $user.Name
        Userid = $user.ID}
     $userobj = New-Object -TypeName PSObject -Property $props
     $userObjs += $userobj   
    }

$props = [ordered]@{
    OSversion        = $os.version
    Model            = $cs.Model
    Manufacturer     = $cs.Manufacturer
    BIOSSerial       = $bios.serialnumber
    ComputerName     = $os.CSName
    OSArchitecture   = $os.OSArchitecture
    ProcArchitecture = $proc.addresswidth
    Disks            = $diskObjs
    Users            = $userobjs
       
}
$obj = New-Object -TypeName PSObject -Property $props

Write-Output $obj | Format-List

#$obj | Get-Member