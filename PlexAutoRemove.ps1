#################################################################################
#
# This script goes through the PLEX played list for shows which have been
# watched and deletes them from the computer.
#
# Created Date: 09 JUL 2015
# Version: 1.0
#
#################################################################################

param (
  $plexToken,
  $tvShowsDir
)

# Variables
$ip = 'localhost' #'192.168.1.2'
$port = '32400'
$plex = "C:\Program Files (x86)\Plex\Plex Media Server\Plex Media Scanner.exe"
[int]$section = ((& $plex 'Media\' 'Scanner\' '--list') -Like '*TV Shows').Split("{:}")[0]

$url = 'http://'+$ip+':'+$port+'/library/sections/'+$section+'/recentlyViewed?X-Plex-Token='+$plexToken
[xml]$recentlyWatched = (New-Object System.Net.WebClient).DownloadString($url)
$watchedShows = $recentlyWatched.GetElementsByTagName('Video')

foreach ($show in $watchedShows) {
  [int]$views = $show.getAttribute("viewCount")
  if ($views -gt 0) {
    $lastViewed = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($show.lastViewedAt))
    # If show was watched over a month ago delete it.
    if ($lastViewed -lt (Get-Date).AddMonths(-1)) {
      rm ($show.GetElementsByTagName('Media').getElementsByTagName("Part").getAttribute("file")) -Force
    }
  }  
}

sl $tvShowsDir
Get-ChildItem -recurse | Where {$_.PSIsContainer -and `
@(Get-ChildItem -Lit $_.Fullname -r | Where {!$_.PSIsContainer}).Length -eq 0} |
Remove-Item -recurse
