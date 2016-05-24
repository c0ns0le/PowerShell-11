$aacgain = 'D:\Dropbox\Applications\Open Source\Audio Apps\aacgain_1_9\aacgain.exe'

sl 'D:\TV Shows'
$shows = gci -Include '*.mp4' -Recurse
$showCount = $shows.Count
$counter = 0;
foreach ($show in $shows) {
	  $counter++
	  Write-Progress -activity "Analysing ""$($show.BaseName)""..." -status "$counter of $showCount" -percentComplete (($counter / $showCount) * 100)
    $show.BaseName
    
    try {
      $aacParams = '/k','/r',$show
      & $aacgain $aacParams
    } catch {
      Write-Host "Exception: $_.Exception.Message"
      Write-Host "Item: $_.Exception.ItemName"      
    }    
}


