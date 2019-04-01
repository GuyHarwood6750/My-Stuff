workflow GetComputerInfo {
    $computers = "Guy-7750G", "Latitude-E5570"
    foreach -parallel ($cn in $computers)
    {Get-CimInstance -PSComputername $cn -ClassName Win32_ComputerSystem}
      
}