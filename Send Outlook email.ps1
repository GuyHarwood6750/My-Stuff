$OL = New-Object -ComObject outlook.application
#
$mItem = $OL.CreateItem("olMailItem")
$mItem.To = 'harwoodg123@gmail.com'
$mItem.Subject = 'Powershell mail testing'
$mItem.Body = "Sent from Powershell script"
#
$mItem.Send()
#$OL.Quit()
#ReleaseComObject

#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OL) | Out-Null
