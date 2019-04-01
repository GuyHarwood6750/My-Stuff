$files = Get-ChildItem "C:\Test" -Recurse -file | group Extension
$files | sort count -Descending | select -First 5 Count, Name, @{Name="Size";Expression={($_.group | Measure-Object Length -sum).sum}}