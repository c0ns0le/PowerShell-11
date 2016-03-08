#################################################################################
#
# This script goes through the PLEX played list for shows which have been
# watched that are on my 'to delete' list and deletes them from the computer.
#
# Created Date: 11 OCT 2013
# Version: 1.8
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

foreach ($watchedShow in $watchedShows) {
  # $lastViewed = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($show.lastViewedAt))
  $watchedTitle = $watchedShow.grandparentTitle.Replace("'","").Replace(".","").Replace(":","").ToLower()
  if ($showsToDelete -contains $watchedTitle) {
    [int]$views = $watchedShow.getAttribute("viewCount")    
    if ($views -gt 0) {
      # Get the path of the previous episode to delete.  Keeps the file in 'On Deck'.
      $parentPath = (Split-Path -Parent $watchedShow.GetElementsByTagName('Media').getElementsByTagName("Part").getAttribute("file")) + "\"
      $episodeData = $watchedShow.GetElementsByTagName('Media').getElementsByTagName("Part").getAttribute("file") -split " - "
      $seasonEp = $episodeData[1].split("sSeE")
      # If episode is greater than 1, get previous episode.      
      if ([int]$seasonEp[2] -gt 1) {
        # Get Episode number
        [string]$episode = [int]$seasonEp[2] - 1
        # Create build string to search for previous episode.
        $newFileName = ($episodeData[0] + " - S" + $seasonEp[1] + "E" + $episode.PadLeft(2,"0")).Replace($parentPath,"")
        # Get previous episode.
        $lastEp = Get-ChildItem -Path $parentPath -Filter "$newFileName*"
        # Add episode to pending deletes.
        if ($lastEp) {
          [array]$pendingDeletes += $lastEp.FullName
        }
      # Else get previous episode of previous season.
      } elseif (([int]$seasonEp[1] -gt 1) -and ([int]$seasonEp[2] -eq 1) ) {
        $parentSplit = $parentPath -split "Season "
        # Get season number.
        [string]$season = [int]$seasonEp[1] - 1
        $newParent = $parentSplit[0] + "Season " + $season.PadLeft(2,"0") + "\"
        # Find last episode of previous season.
        $lastEp = Get-ChildItem -Path $newParent | Sort-Object Name -Descending | Select-Object -First 1
        # Add episode to pending deletes.
        if ($lastEp) {
          [array]$pendingDeletes += $lastEp.FullName        
        }
      }
    }
  }
}

# Find shows in deleted eps which haven't been watched in over a year.
$url = 'http://'+$ip+':'+$port+'/library/sections/'+$section+'/all?X-Plex-Token='+$plexToken
[xml]$shows = (New-Object System.Net.WebClient).DownloadString($url)
[array]$allShows = $shows.ChildNodes[1].Directory
foreach ($show in $allShows) {
  if (($showsToDelete -contains $show.Title)) {
    $lastViewedAt = ([datetime]'1970-01-01 00:00:00').AddSeconds($show.LastViewedAt)
    if ($lastViewedAt -lt (Get-Date).AddYears(-1)) {
      $ratingKey = $show.ratingKey
      $showUrl = 'http://'+$ip+':'+$port+'/library/metadata/'+$ratingKey+'/allLeaves?X-Plex-Token='+$plexToken
      [xml]$eps = (New-Object System.Net.WebClient).DownloadString($showUrl)
      [array]$pendingDeletes += @($eps.GetElementsByTagName('Video').GetElementsByTagName('Media').getElementsByTagName("Part").getAttribute("file"))      
    }
  } 
}

foreach ($delete in $pendingDeletes) {
  if (Test-Path $delete) {
    rm $delete -Force 
  }
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
