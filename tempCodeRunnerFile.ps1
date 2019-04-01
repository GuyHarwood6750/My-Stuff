$collection = New-Object -TypeName System.Collections.ArrayList
$random = New-Object -TypeName System.Random
foreach ($i in 1..1000) {
    $null = $collection.Add($random.Next(0,1000))
}