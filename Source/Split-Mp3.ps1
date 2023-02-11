$Location = $args[0]

if ($null -eq $Location) {
    $Location = Read-Host "Specify location of a playlist"
}

while (-not (Test-Path -Path $Location)) {
    $Location = Read-Host "Please, specify the correct location of a playlist"
}

$CsvExists = Test-Path -Path (Join-Path -Path $Location -ChildPath "playlist.csv")
$Mp3Exists = Test-Path -Path (Join-Path -Path $Location -ChildPath "playlist.mp3")

while ((-not $CsvExists) -or (-not $Mp3Exists)) {
    Write-Host "Please, make sure that the playlist.csv and playlist.mp3 files are present in the specified location and press Enter"
    if ([Console]::ReadKey().Key -eq "Enter") {
        $CsvExists = Test-Path -Path (Join-Path -Path $Location -ChildPath "playlist.csv")
        $Mp3Exists = Test-Path -Path (Join-Path -Path $Location -ChildPath "playlist.mp3")
    }
}

Set-Location -Path $Location

$Playlist = Get-Content ".\playlist.csv" | ConvertFrom-Csv -Delimiter ";"

$Seek = 0

foreach ($Track in $Playlist) {
    $TrackNumber = "{0:d2}" -f [int]$Track.Track
    $Duration = ([TimeSpan]$Track.Duration).TotalSeconds

    Invoke-Command { ffmpeg -ss $Seek -y -i ".\playlist.mp3" -t $Duration -c:a libmp3lame -b:a 320k -metadata track="$($Track.Track)" -metadata artist="$($Track.Artist)" -metadata title="$($Track.Title)" -metadata album="$($Track.Album)" -metadata date="$($Track.Year)" ".\$TrackNumber. $($Track.Artist) - $($Track.Title) - $($Track.Album).mp3" }

    $Seek = $Duration
}