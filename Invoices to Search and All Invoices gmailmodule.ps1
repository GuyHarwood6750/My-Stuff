
$vpnok = Get-VpnConnection
if ($vpnok.ConnectionStatus -eq 'connected')
    {
        Copy-Item -Path 'C:\Pastel18\CIRCE19A\Documents\PDF\*.pdf' -Destination "C:\Search Invoices" -Recurse -force
        Move-Item -Path 'C:\Pastel18\CIRCE19A\Documents\PDF\*.pdf' -Destination "\\wserver\wmarine\customers\_All Invoices & Credit Notes" -Force
     }
    else 
    {
              
        Guy-SendGmail "Copy of PDF files to Searches Folder and Offsite location failed" "PLEASE INVESTIGATE"
        
        #$emailMessage.Subject = "Copy of PDF files to Searches Folder and Offsite location failed" 
        #$emailMessage.Body = "PLEASE INVESTIGATE"

    }