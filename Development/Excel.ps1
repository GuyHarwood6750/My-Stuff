$xl = New-Object -ComObject excel.application
$xl.visible = $true
$wb = $xl.Workbooks.Add()
$ws = $wb.worksheets["sheet1"]
$ws.cells[1, 1].value = "id"
$ws.cells[1, 2].value = "product"
$ws.cells[1, 3].value = "qty"
$ws.cells[1, 4].value = "price"
$ws.cells[1, 5].value = "value"
$ws.cells[2, 1].value = 12001
$ws.cells[2, 2].value = "Nails"
$ws.cells[2, 3].value = 37
$ws.cells[2, 4].value = 3.99
$ws.cells[2, 5].Formula = "=c2*d2"
$ws.cells[3, 1].value = 12002
$ws.cells[3, 2].value = "Hammer"
$ws.cells[3, 3].value = 5
$ws.cells[3, 4].value = 45.29
$ws.cells[3, 5].Formula = "=c3*d3"
$ws.cells[4, 1].value = 12003
$ws.cells[4, 2].value = "Saw"
$ws.cells[4, 3].value = 10
$ws.cells[4, 4].value = 55.29
$ws.cells[4, 5].Formula = "=c4*d4"