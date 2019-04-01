$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$elapsed = $StopWatch.Elapsed
Write-Host "All completed. Elapsed time: $elapsed"