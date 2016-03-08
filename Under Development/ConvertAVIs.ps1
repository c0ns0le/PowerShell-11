$uTorrentHome = $env:USERPROFILE+"\Downloads\Torrents\"
$completed = $uTorrentHome+'Completed\'
$toConvert = $uTorrentHome+'To Convert\'

sl $completed
$files = ls *.avi -Recurse

foreach ($f in $files) {
  mv -LiteralPath $f.FullName $toConvert
}