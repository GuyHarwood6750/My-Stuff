New-PSDrive -Name "K" -PSProvider "FileSystem" -Root "\\wserver\wmarine\Kiosk" -Persist -Scope Global
New-PSDrive -Name "L" -PSProvider "FileSystem" -Root "\\wserver\wmarine\Customers" -Persist -Scope Global
New-PSDrive -Name "M" -PSProvider "FileSystem" -Root "\\wserver\wmarine\Management" -Persist -Scope Global