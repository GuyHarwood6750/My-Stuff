#$datalist = "C:\test\invoices\modified invoices.csv"
$datalist = "C:\userdata\circe launches\monthly invoices\daily modified invoices.csv"

#$fileloc = "C:\test\Invoices"
$fileloc = "\\wserver\wmarine\Customers\_All Invoices & Credit Notes"  

$data = Get-Content -Path $datalist

foreach ($renamef in $data) {
    $invoicename = "Tax Invoice IN$renamef.pdf"
    $renamemod1 = "Tax Invoice IN$renamef-1.pdf"
    $renamemod2 = "Tax Invoice IN$renamef-2.pdf"
    $renamemod3 = "Tax Invoice IN$renamef-3.pdf"

    $filex = Get-ChildItem -Path $fileloc -Filter $invoicename
    $filey = Get-ChildItem -path $fileloc -Filter $renamemod1
    $filez = Get-ChildItem -path $fileloc -Filter $renamemod2

    if ($filex.Name -eq $invoicename) {
        Rename-Item -path $fileloc\$invoicename -NewName $renamemod1
        #Write-Host "$filex was found"
    }
    elseif ($filey.name -eq $renamemod1) {
        Rename-Item -Path $fileloc\$renamemod1 -NewName $renamemod2
        #Write-Host "$filey was found"
    }       
    elseif ($filez.Name -eq $renamemod2) {
        Rename-Item -Path $fileloc\$renamemod2 -NewName $renamemod3
        #Write-Host "$filez was found"
    }
    else {
        Write-Output "Problem..............."
    } 
   
}
#Guy-SendGmail "Rename of modified invoices failed" "Please investigate"
    