$moviesFolder = 'D:\Movies\'
$moviesCSV = Import-CSV -Header 'Index','File','Title' 'C:\Users\Steve\Desktop\Movies.csv'
$digitalOnly = Import-CSV -Header 'Index','File','Title' 'C:\Users\Steve\Desktop\DigitalOnly.csv'
$outputFolder = 'D:\Encoded\'
$todo = $outputFolder+"Re-EncodeList.txt"
$exifTool = 'D:\Dropbox\Applications\Open Source\Utilities\exiftool.exe'
$handbrake = "C:\Program Files\Handbrake\HandBrakeCLI.exe"
$roundingMultiple = 8;

sl $moviesFolder
$movies = ls *.m4v,*.mp4
$reEncode = @()
for ($i = 0; $i -lt ($movies.Length); $i++) {
  $movieName = $movies[$i].Name
  $perc = (($i / $movies.Length) * 100)
  Write-Progress -activity "Processing $movieName" -Status "$perc% Completed" -PercentComplete $perc
  $fileAttrHash = @{}
  if ($movies[$i]) {
    $fileAttr = & $exifTool $movies[$i]
    foreach ($attr in $fileAttr) {
      if ($attr) {
        $key = $attr -split ':',2
        $fileAttrHash += @{$key[0].Trim()=$key[1].Trim()}
      }
    }  
    [int]$videoWidth = $fileAttrHash['Source Image Width']
    [int]$videoHeight = $fileAttrHash['Source Image Height']
    [int]$displayWidth = $fileAttrHash['Image Width']
    [int]$displayHeight = $fileAttrHash['Image Height']
    
    if ([int]$videoWidth -ne [int]$displayWidth) {
      if ([int]$videoWidth -le 720) {
        $width = 1024
        if ([int]$displayHeight -ge 576) {
          $height = 576
        } else {
          $height = $displayHeight
        }
       
        $input = '"'+$movies[$i].FullName+'"'
        $output = $input.Replace($moviesFolder,$outputFolder)
        
        $digital = $false
        foreach ($file in $digitalOnly) {
          if ($file.File -like '*'+$movieName+'*') {
            $digital = $true
            break
          }
        }
        if ($digital -eq $true) {
          if (!(Test-Path $output)) {
            $params = '-i',$input,'-t','1','--angle','1','-c','1','-o',$output,'-f','mp4','-O','-w',$width,'-l',$height,`
            '--crop','0:0:0:0','--modulus','16','-e','x264','-q','22','--vfr','-a','1,1','-E','copy:ac3,copy:aac','-6', `
            'none,none','-R','Auto,Auto','-B','0,0','-D','0,0','--gain','0,0','--audio-fallback','ac3',`
            '--encoder-tune="film"','--encoder-level="4.0"','--encoder-profile=high','--verbose=1'
            Write-Host "Encoding $input to $output..."
            [Diagnostics.Process]::Start($handbrake, $params).WaitForExit()
          }
        } else {
          $index = $null
          $title = $null
          foreach ($row in $moviesCSV) {
            if ($row.File -like '*'+$movieName+'*') {
              $index = "{0:D3}" -f [int]$row.Index
            }
          }
          $reEncode += $index+"`t"+$movieName+"`t"+[string]$videoWidth+"x"+[string]$videoHeight
        }
      }
    } else {
      if ($videoWidth -ne 0) {
        $digital = $false
        foreach ($file in $digitalOnly) {
          if ($file.File -like '*'+$movieName+'*') {
            $digital = $true
            break
          }
        }
        if ($digital -ne $true) {
          if (([int]$videoWidth -le 720) -and ([int]$displayHeight -le 405)) {
            $index = $null
            $title = $null
            foreach ($row in $moviesCSV) {
              if ($row.File -like '*'+$movieName+'*') {
                $index = "{0:D3}" -f [int]$row.Index
              }
            }
            $reEncode += $index+"`t"+$movieName+"`t"+[string]$videoWidth+"x"+[string]$videoHeight
          }
        }
      }
    }
  }
}
$reEncode = $reEncode | Sort-Object
$reEncode | out-file $todo