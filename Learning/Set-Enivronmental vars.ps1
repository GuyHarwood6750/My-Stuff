[System.Environment]::setEnvironmentVariable('Firstname', 'Guy', [System.EnvironmentVariableTarget]::User)
[System.Environment]::setEnvironmentVariable('Lastname', 'Harwood', [System.EnvironmentVariableTarget]::User)
[System.Environment]::setEnvironmentVariable('City', 'Barrydale', [System.EnvironmentVariableTarget]::User)
[System.Environment]::setEnvironmentVariable('FullName', 'Guy Harwood', [System.EnvironmentVariableTarget]::User)

$env:Firstname
$env:Lastname
$env:City
$env:fullname
$env:HOMEPATH
$env:Path
$env:PSModulePath
[System.Math]::Round( (1904002170088 / 1gb), 2)