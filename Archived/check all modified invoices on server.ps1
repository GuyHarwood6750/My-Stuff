Get-ChildItem -path "W:\customers\_all invoices & credit notes" -filter "*10404*.pdf" | select name | sort name -Descending | export-csv -Path "c:\test\invoices\modflist.csv"