################################################################################
#                                                                              #
# This script losslessly embeds SRT files into m4v files using ffmpeg          #
#                                                                              #
# Created Date: 09 MAY 2015                                                    #
# Version: 1.0                                                                 #
#                                                                              #
################################################################################

$mp4Box = 'C:\Program Files\My MP4Box GUI\Tools\MP4Box.exe'
sl 'D:\Movies\'
$subs = ls *.srt
foreach ($sub in $subs) {
  $m4v = $sub.FullName.Replace('eng.srt','m4v')
  $subFile = $sub.FullName
  $subtitles = $sub.FullName+':lang=eng'
  $mp4BoxParams = '-add',$subtitles,$m4v
  if (Test-Path $m4v) {
    & $mp4Box $mp4BoxParams
    rm $subFile -Force
  }
}
