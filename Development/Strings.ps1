####################################################

# Strings vs. Stringbuilder

####################################################

<#

	$obj = "String"

	$obj = "String" + " " + "Builder"

	$obj = "$obj Builder"

	Strings are Immutable – cannot be changed



	[System.Text.StringBuilder]



	$obj = new-object –typename System.Text.StringBuilder –Args 4096

	$obj.Append(" ") or $obj.AppendLine(" ")

	$obj.AppendFormat("{0} {1}", "one", "two")



	Returns an object

#>



$str1 = "This is a string"

$str2 = " and I added it to another"



$str3 = $str1 + " " + $str2

$str3 = "$str1 $str3"



$str3 = New-Object -TypeName System.Text.StringBuilder -Args 8192



$null = $str3.Append($str1)

$null = $str3.Append(" ")

$null = $str3.Append($str2)



$null = $str3.AppendFormat("{0} {1}", $str1, $str2)



$str3.ToString()



$bigstring = ""

Measure-Command -Expression { 1 .. 10000 | ForEach-Object { $bigstring += $_ } }

Measure-Command -Expression { 1..10000 | ForEach-Object { $null = $str3.Append($_) } }

