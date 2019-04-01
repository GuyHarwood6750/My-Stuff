"Hello world" | Export-Excel
    [PSCustomObject]@{Data ="Hallo you stupid world"} | Export-Excel
    get-process | Export-Excel '.\demo.xlsx'
    Invoke-Item '.\demo.xlsx'
    remove-item '.\demo.xlsx' -ErrorAction ignore
    Get-Process | select company,pm,handles | Export-Excel '.\demo.xlsx' -Show
    $ps=get-process | Select company, pm, Handles
    $ps | Export-Excel '.\demo.xlsx' -Show -IncludePivotTable -IncludePivotChart -PivotRows Company -PivotData @{handles='sum'}
    $ps | Export-Excel '.\demo.xlsx' -Show -ChartType PieExploded3D -IncludePivotTable -IncludePivotChart -PivotRows Company -PivotData @{Handles='sum'}
    $ps | Export-Excel '.\demo.xlsx' -Show -ChartType PieExploded3D -Nolegend -showcategory -showpercent -IncludePivotTable -IncludePivotChart -PivotRows Company -PivotData @{Handles='sum'}
    $dat = Get-Content .\data.csv
    Import-Csv .\data.csv | Export-Excel test.xlsx -Show -TableName items -AutoSize
    Remove-Item .\test.xlsx
    $data = import-csv .\data.csv | Export-Excel test.xlsx -Show -AutoNameRange
    $barchart = New-ExcelChart -XRange Product -YRange Total -Height 250
    $pieChart = New-ExcelChart -ChartType PieExploded3D -Row 15 -Height 250 -XRange Product -YRange Total -NoLegend -ShowCategory -ShowPercent
    $data | Export-Excel .\test.xlsx -Show -AutoNameRange -ExcelChartDefinition $barchart, $pieChart
    $data2 = import-csv .\Status2.csv | Format-Table
    $chart = New-ExcelChart -XRange Customer -YRange QTY -Title "Total by Company" -Column 9 -Height 600 -Width 500
    $excelparams = @{show=$true; AutoSize=$true; AutoNameRange=$true; TableName="Sales"}
    Import-Csv .\Status2.csv | Select Customer, QTY | Export-Excel @excelparams demo2.xlsx -ExcelChartDefinition $chart
    Remove-Item .\demo2.xlsx
    $data3 = Get-Service | select Status,Name,DisplayName,StartType
    $data3 | Export-Excel test3.xlsx -show -AutoSize
    Remove-Item .\test3.xlsx
    $check1 = New-ConditionalText stop
    $data3 | Export-Excel .\test3.xlsx -Show -AutoSize -ConditionalText $check1
    $check2 = New-ConditionalText runn blye cyan
    $data3 | Export-Excel .\test3.xlsx -show -AutoSize -ConditionalText $check1, $check2
    $check3 = New-ConditionalText svc wheat green
    $data3 | Export-Excel .\test3.xlsx -show -AutoSize -ConditionalText $check1, $check2, $check3
    $data4 = Get-Process | select Company,name,pm,Handles, *mem*
    $cfmt = New-ConditionalFormattingIconSet -Range "c:c" -ConditionalFormat ThreeIconSet -IconType Arrows
    $data4 | Export-Excel test3.xlsx -Show -AutoSize -ConditionalFormat $cfmt
    $url="http://www.science.co.il/ptelements.asp"
    $index = 0
    start $url
    Import-Html $url $index #this fails on my machine
    $data5 = import-csv .\data.csv
    $data5 | BarChart -title "a fabulous bar chart"
    $data5 | PieChart -title "a pie chart"
    $data5 | LineChart -title "line chart"
    $data5 | ColumnChart -title "column chart"



