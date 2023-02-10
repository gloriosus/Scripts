Set-Location -Path $args[0]

$Playlist = Get-Content ".\playlist.csv" | ConvertFrom-Csv -Delimiter ";"

$Seek = 0

foreach ($Track in $Playlist) {
    $TrackNumber = "{0:d2}" -f [int]$Track.Track
    $Duration = ([TimeSpan]$Track.Duration).TotalSeconds

    Invoke-Command { ffmpeg -ss $Seek -i ".\playlist.mp3" -t $Duration -c:a libmp3lame -b:a 320k -metadata track="$($Track.Track)" -metadata artist="$($Track.Artist)" -metadata title="$($Track.Title)" -metadata album="$($Track.Album)" -metadata date="$($Track.Year)" ".\$TrackNumber. $($Track.Artist) - $($Track.Title) - $($Track.Album).mp3" }

    $Seek = $Duration
}