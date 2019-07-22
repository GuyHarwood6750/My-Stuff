$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
$playsoung.playsync()
$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
$playsoung.playsync()
#