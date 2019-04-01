$Source = "C:\TEST\SEARCH"
$DestArchive = "C:\TEST\DestF"

Get-ChildItem $Source | ForEach-Object {

    $filename = $_.Name

    if (Test-Path "$DestArchive\$filename")
    {
        Copy-Item $_.FullName -Destination $DestArchive
        Write-Host 'file exists - not copying'
    }
    else
    {
        Copy-Item $_.FullName -Destination $DestArchive
        Write-Host 'file not found'
    }

}