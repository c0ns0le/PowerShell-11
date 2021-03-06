$files = Get-ChildItem -Path $env:USERPROFILE+'\Pictures\iCloud Photos\My Photo Stream' -Filter *.jpg
foreach ($file in $files) {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $i = new-object System.Drawing.Bitmap $file.FullName
    $i.RotateFlip('Rotate180FlipNone')
    $i.Save($file.FullName)
    $dateString =[System.Text.Encoding]::ASCII.GetString($i.GetPropertyItem(36867).Value)
    $dateString = ($dateString.Replace(':','-').SubString(0,11))+($dateString.Replace(':','.').SubString(10,9))
}