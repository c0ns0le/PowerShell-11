$dropBox = "D:\Dropbox"
$ffmpeg = $dropBox+'\Applications\Open Source\Video Apps\ffmpeg\bin\ffmpeg.exe'
$showsdir = 'C:\Users\Steve\Downloads\Torrents\Untagged'

sl $showsdir
$shows = gci -Include '*.mkv' -Recurse
foreach( $show in $shows ){
  $newMp4Path = $show.Name -replace "\.(mkv)$",".mp4"
  $ffmpegParams = '-i',$show,'-acodec','copy','-vcodec','copy',$newMp4Path

  try {
    & $ffmpeg $ffmpegParams
  } catch {
    "Exception: " + $_.Exception.Message
    "Item: " + $_.Exception.ItemName
  }
}