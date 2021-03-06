$exifTool = $env:USERPROFILE+'\Dropbox\Applications\Open Source\Utilities\exiftool.exe'
sl "$env:USERPROFILE\My Documents\My Pictures\My Photos"
$photos = ls *.jpg,*.png -Recurse
foreach ($photo in $photos) {
  if ($photo.Name -notmatch "[0-9]{4}[-+][0-9]{1,2}[-+][0-9]{1,2}\s*[0-9]{2}\.[0-9]{2}\.[0-9]{2}\.[a-zA-Z]{3}[-+]*?[0-9]*?") {
    $photoDate = (& $exifTool -TAG -DateTimeOriginal $photo)
    if ($photoDate) {
      $date = $photoDate.Replace('Date/Time Original              : ','')
      $dateString = $date.SubString(0,4)+"-"+$date.SubString(5,2)+"-"+$date.SubString(8,2)+" "+$date.SubString(11,2)+"."+$date.SubString(14,2)+"."+$date.SubString(17,2)
      (New-Object PsObject -Property @{Photo=$photo.FullName;Date=$dateString}) | Export-Csv -Path "$env:USERPROFILE\Desktop\UnnamedPhotos.csv" -NoTypeInformation -Append
    }
  }  
}