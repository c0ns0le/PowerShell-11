#################################################################################
#
# This script goes through the PLEX played list for shows which have been
# watched that are on my 'to delete' list and deletes them from the computer.
#
# Created Date: 11 OCT 2013
# Version: 1.6
#
#################################################################################

param (
  $plexToken = 'dgnK5Hx8GfBnsjcZaASw',
  $tvShowsDir = 'D:\TV Shows'
)

# Variables
$ip = 'localhost' #'192.168.1.2'
$port = '32400'
$plex = "C:\Program Files (x86)\Plex\Plex Media Server\Plex Media Scanner.exe"
[int]$section = ((& $plex 'Media\' 'Scanner\' '--list') -Like '*TV Shows').Split("{:}")[0]
$showsToDelete = 'offspring',
  'greys anatomy',
  'marvels agents of shield',
  'top gear',
  'comic book men',
  'survivor',
  'the middle',
  'selling houses australia',
  'grand designs',
  'american horror story',
  'play it again, dick',
  'the comeback'

$url = 'http://'+$ip+':'+$port+'/library/sections/'+$section+'/recentlyViewed?X-Plex-Token='+$plexToken
[xml]$recentlyWatched = (New-Object System.Net.WebClient).DownloadString($url)
$watchedShows = $recentlyWatched.GetElementsByTagName('Video')

foreach ($show in $watchedShows) {
  # $lastViewed = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($show.lastViewedAt))
  $watchedTitle = $show.grandparentTitle.Replace("'","").Replace(".","").Replace(":","").ToLower()
  if ($showsToDelete -contains $watchedTitle) {
    [int]$views = $show.getAttribute("viewCount")    
    if ($views -gt 0) {
      [array]$pendingDeletes += $show.GetElementsByTagName('Media').getElementsByTagName("Part").getAttribute("file")
    }
  }
}

foreach ($delete in $pendingDeletes) {
  rm $delete -Force 
}

If ($pendingDeletes) {
  # Rescan after deletes.
  $plex = "C:\Program Files (x86)\Plex\Plex Media Server\Plex Media Scanner.exe"
  $plexParams = '--scan','--section',$section
  [Diagnostics.Process]::Start($plex,$plexParams).WaitForExit()
}

sl $tvShowsDir
Get-ChildItem -recurse | Where {$_.PSIsContainer -and `
@(Get-ChildItem -Lit $_.Fullname -r | Where {!$_.PSIsContainer}).Length -eq 0} |
Remove-Item -recurse -Force
