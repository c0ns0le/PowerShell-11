$moviesFolder = 'D:\Movies\'
$outputFolder = 'D:\Encoded\'
$exifTool = 'D:\Dropbox\Applications\Open Source\Utilities\exiftool.exe'
$handbrake = "C:\Program Files\Handbrake\HandBrakeCLI.exe"
$roundingMultiple = 8;

sl $moviesFolder
$movies = ls *.m4v,*.mp4
for ($i = 0; $i -lt ($movies.Length); $i++) {
  $movieName = $movies[$i].Name
  Write-Progress -activity "Processing $movieName" -PercentComplete (($i / $movies.Length)  * 100)    
  $fileAttrHash = @{}
  $fileAttr = & $exifTool $movies[$i]
  foreach ($attr in $fileAttr) {
    if ($attr) {
      $key = $attr -split ':',2
      $fileAttrHash += @{$key[0].Trim()=$key[1].Trim()}
    }
  }  
  [int]$videoWidth = $fileAttrHash['Source Image Width']
  [int]$displayWidth = $fileAttrHash['Image Width']
  [int]$displayHeight = $fileAttrHash['Image Height']
  
  if ($videoWidth -ne $displayWidth) {
    if ($videoWidth -le 720) {
      $width = 1024
      if ($displayHeight -ge 576) {
        $height = 576
      } else {
        $height = $displayHeight
      }
      $input = '"'+$movies[$i].FullName+'"'
      $output = $input.Replace($moviesFolder,$outputFolder)

      $params = '-i',$input,'-t','1','--angle','1','-c','1','-o',$output,'-f','mp4','-O','-w',$width,'-l',$height,`
      '--crop','0:0:0:0','--modulus','16','-e','x264','-q','22','--vfr','-a','1,1','-E','copy:ac3,copy:aac','-6', `
      'none,none','-R','Auto,Auto','-B','0,0','-D','0,0','--gain','0,0','--audio-fallback','ac3',`
      '--encoder-tune="film"','--encoder-level="4.0"','--encoder-profile=high','--verbose=1'
      Write-Host "Encoding $input to $output..."
      [Diagnostics.Process]::Start($handbrake, $params).WaitForExit()
    }
  }  
}