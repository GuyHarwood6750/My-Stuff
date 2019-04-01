$status = Test-Connection wserver
 if ($status.statuscode -eq 0) {
    move-item -Path '\\wdmycloudmirror\public\userdata\circe launches\daily*.xlsx' -Destination 'M:\Daily Reports'
    move-item -Path '\\wdmycloudmirror\public\userdata\circe launches\invoiceonly*.xlsx' -Destination 'M:\Daily Reports'
    move-item -path '\\wdmycloudmirror\Public\Userdata\Circe Launches\invoicevouchers*.xlsx' -Destination '\\wserver\wmarine\Booking Reports\Guy'
    move-item -path '\\wdmycloudmirror\Public\Userdata\Circe Launches\NoGuideName*.xlsx' -Destination '\\wserver\wmarine\Booking Reports\Julia'
 }
 else {
     Write-output "Server not found, aborting...."
 }