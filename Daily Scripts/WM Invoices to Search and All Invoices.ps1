
$vpnok = Get-VpnConnection
$locallconn = Test-NetConnection -ComputerName "wserver" -InformationLevel Detailed

if ($vpnok.ConnectionStatus -eq 'connected')
    {
        Copy-Item -Path 'C:\Pastel18\CIRCE21A\Documents\PDF\*.pdf' -Destination "C:\Search Invoices" -Recurse -force
        Move-Item -Path 'C:\Pastel18\CIRCE21A\Documents\PDF\*.pdf' -Destination "\\wserver\wmarine\customers\_All Invoices & Credit Notes" -Force
     }
    
    elseif($locallconn.PingSucceeded -eq 'True'){
        Copy-Item -Path 'C:\Pastel18\CIRCE21A\Documents\PDF\*.pdf' -Destination "C:\Search Invoices" -Recurse -force
        Move-Item -Path 'C:\Pastel18\CIRCE21A\Documents\PDF\*.pdf' -Destination "\\wserver\wmarine\customers\_All Invoices & Credit Notes" -Force
  }   
   
    else 
    {
        Guy-SendGmail "Copy of PDF files to Searches Folder and Offsite location failed" "PLEASE INVESTIGATE - Script -> WM Invoices to Search and All Invoices"
    }