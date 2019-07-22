Add-Type -AssemblyName presentationCore
$mediaPlayer = New-Object system.windows.media.mediaplayer
$mediaPlayer.open('C:\test\mp3\ES_All Of Us - Daxten.mp3')
$mediaPlayer.Play()
