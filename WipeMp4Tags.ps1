param (
  $directory = 'D:\Buffy the Vampire Slayer'
)

$dropBox = "D:\Dropbox"
$atomicParsley = $dropBox+'\Applications\Open Source\Video Apps\AtomicParsley\AtomicParsley.exe'

sl $directory
$files = ls *.m4v,*.mp4 -Recurse
  
foreach ($file in $files) {
  try {
    $atomicParams = $file,'--metaEnema','--overWrite'
    & $atomicParsley $atomicParams
  } catch {
    Write-Host "`r`nError: "$_.Exception.Message"`r`n"$_.Exception.ItemName
  }
}