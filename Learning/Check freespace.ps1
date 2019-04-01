<#
Check to see which harddrive has less than 20% free space
#>
Get-WmiObject -Class win32_logicaldisk -filter "Drivetype=3" |
    Where-Object { $_.Freespace / $_.Size -lt .2 } |
    Format-Table @{name = 'DriveLetter'; Expression = {$_.DeviceID}},
    @{name='Size';Expression={$_.Size / 1GB};FormatString='N2'},
    @{name='Freespace';Expression={$_.Freespace / 1gb};FormatString='N2'} -auto