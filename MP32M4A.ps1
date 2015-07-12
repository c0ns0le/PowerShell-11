param (
  $directory = 'C:\Users\Steve\Downloads'
)
$ffmpeg = 'D:\Dropbox\Applications\Open Source\Video Apps\MkvToMp4_0.224\Tools\ffmpeg\x64\ffmpeg.exe'

sl $directory
$mp3s = ls *.mp3 -Recurse

foreach ($mp3 in $mp3s) {
  $origFile = $mp3.Fullname
  $newFile = $origFile.Replace('.mp3','.m4a')
  $params = @($origFile,'-c:a','libfdk_aac','-b:a','256k','-y',$newFile)
  & $ffmpeg $params
}