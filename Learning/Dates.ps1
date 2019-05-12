#$datefilter = (get-date).AddDays(-1).ToString(("dd/MM/yyyy"))
#$datefilter = (get-date).ToString(("dd/MM/yyyy"))
#$date = $datefilter.tostring("yyyy/MM/dd")
$invoice = '21/12/2019'
$newdate = [datetime]::ParseExact($invoice, 'dd/MM/yyyy', $null)
#$newdate | gm
#$newdate
#$newdate.Month
#$newdate.Year
#$newdate.Day
#$newdate.Date
#$newdate.DayOfWeek
