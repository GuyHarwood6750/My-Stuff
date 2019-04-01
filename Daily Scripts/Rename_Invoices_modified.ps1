$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()

#$datalist = "C:\test\invoices\modified invoices.csv"
$datalist = "C:\userdata\circe launches\monthly invoices\daily modified invoices.csv"

#$fileloc = "C:\test\Invoices"
$fileloc = "\\wserver\wmarine\customers\_All Invoices & Credit Notes"  

$data = Get-Content -Path $datalist

foreach ($renamef in $data) {
    

    $filex = Get-ChildItem -Path $fileloc -Filter *$renamef*
        
    foreach ($item in $filex) {
        IF ($item.Name) {
            $newname = $item.basename + "-1.pdf"
            $oldname = $item.Name
            Rename-Item -Path $fileloc\$oldname -NewName $newname
            #Write-Host $item.Name 
        }
        else {
            #Write-Host $item.Name
        }
   
        #Write-Output "Problem..............."
    }     
    
}
#Guy-SendGmail "Rename of modified invoices failed" "Please investigate"
$elapsed = $StopWatch.Elapsed
Write-Host "Elapsed time: $elapsed"
    