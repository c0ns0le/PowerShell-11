#################################################################################
#                                                                               #
# This script gets passed some arguments from uTorrent, which it uses to        #
# rename, tag with MetaX, move to the Media drive and add to iTunes.            #
#                                                                               #
# Created Date: 20 FEB 2014                                                     #
# Updated Date: 20 FEB 2014                                                     #
# Created By: devonuto                                                          #
# Version: 1.0                                                                  #
#                                                                               #
#################################################################################

#################################################################################
#                                                                               #
# PRE-REQUISITES                                                                #
#                                                                               #
# Add the following line to uTorrent on finish.                                 #
#   powershell -executionpolicy remotesigned -file                              #
#   "$env:USERPROFILE+"\My Documents\WindowsPowerShell\TVShows.ps1"             #
#   "%D" "%N" "%K" "%F"                                                         #
#                                                                               #
# Install WinRAR here: C:\Program Files\WinRAR\UnRAR.exe                        #
# Install Handbrake here: C:\Program Files (x86)\Handbrake\HandBrakeCLI.exe     #
# Put AtomicParsley here:                                                       #
#  $env:USERPROFILE+"\My Documents\Applications\AtomicParsley\AtomicParsley.exe #
# Add this:                                                                     #
# C:\Windows\Microsoft.NET\Framework\uTorrentAPI\Release\AnyCPU\UTorrentAPI.dll # 
#                                                                               #
# Create Folders for uTorrent/Powershell:                                       #
#   $env:USERPROFILE+"\Downloads\Torrents\                                      #
#   $env:USERPROFILE+"\Downloads\Torrents\Tagging\                              #
#   $env:USERPROFILE+"\My Documents\WindowsPowerShell\                          #
#                                                                               #
#################################################################################

#################################################################################
#                                                                               #
# CHANGELOG                                                                     #
#                                                                               #
# 1.0 Initial build.                                                            #
#                                                                               #
#################################################################################

#################################################################################
# PARAMETERS                                                                    #
#################################################################################

param(
  $directory = $null,
  $title = 'Autotag',
  $kind = $null,
  $torrentFile = $null
)

#################################################################################
# GLOBAL VARIABLES                                                              #
#################################################################################

$global:tvdburl = "http://thetvdb.com/"
$global:tvdbAPIKey = "503C63FB12629A81"

#################################################################################
# FUNCTIONS                                                                     #
#################################################################################

# Get the iTunes path.
function GetiTunesPath {
  $itunesLibXml = Get-ChildItem -Path $env:USERPROFILE -Filter 'Itunes Library.xml' -Recurse
  if ($itunesLibXml) {
    [xml]$itunesLib = Get-Content $itunesLibXml.PSPath
    $path = $itunesLib.plist.Item(1).dict.String[1].Replace('file://localhost/','').Replace('/','\').Replace('%20',' ')
  } else {
    return $false
  }
  return $path
}

# Get the Series ID from theTVDb for the TV series specified.
function GetSeries($name) {
  try {
    $url=$tvdburl+"api/GetSeries.php?seriesname="+$name+"&language=English"
    [xml]$series_ws = (New-Object System.Net.WebClient).DownloadString($url)
    if ($series_ws) {
      # Get series id based on name.
      foreach ($Series in $series_ws.Data.Series){
        Write-Host "Reading $series.SeriesName"
          if ($series.SeriesName -eq $name){
          # Set the global series id variable.
          $id = $series.seriesid
          break
        }
      }
    }
  } catch {
    return $false
  }
  return $id
}

# Encode the specified file with HandBrake for AppleTV.
function EncodeWithHB($inputStr,$outputStr){
  $handbrake = "C:\Program Files (x86)\Handbrake\HandBrakeCLI.exe"
  # Add quoted paths to Handbrake (as without them it breaks)
  $inputStr = '"'+$inputStr+'"'
  $outputStr = '"'+$outputStr+'"'
  # Set up handbrake prameters.
  $params = '-i',$inputStr,'-t','1','-c','1','-o',$outputStr,'-f','mp4',
    '-4','--deinterlace="fast"','-w','720','-l','576','--custom-anamorphic',
        '--display-width','1024','--keep-display-aspect','-e','x264','-q','18',
    '--vfr','-a','1','-E','faac','-B','160','-6','dpl2','-R','Auto','-D',
    '0','--gain=0','--audio-copy-mask','none','--audio-fallback','ffac3',
    '-x','ref=1:weightp=1:subq=2:rc-lookahead=10:trellis=0:8x8dct=0',
    '--verbose=1'
  try {
    Write-Host "Encoding $inputStr to $outputStr..."
    [Diagnostics.Process]::Start($handbrake, $params).WaitForExit()
  } catch {
    return $false
  }  
}

# Unpack the specified compressed files to the output directory provided.
function UnRarFiles($directory,$outputDir,$file) {
  $unRar = "C:\Program Files\WinRAR\UnRAR.exe"
  if (!(Test-Path $outputDir)) {
    Write-Host "Creating directory $outputDir"
    md $outputdir
  }
  sl -LiteralPath $directory
  $unRarParams = 'x',$file,'*.mp4',$outputDir,'-o+'
  try {
    Write-Host "Upacking $directory to $outputDir"
    [Diagnostics.Process]::Start($unrar, $unRarParams).WaitForExit()
  } catch {
    Write-Error -Message "$unRar failed to unpack $file" -Category OpenError
    return $false
  }
}

# A function to remove old files from passed in path and filename type.
function RemoveOldFiles() {
  param ($path,$files = '*.*')

  sl $path
  $date = (Get-Date).AddDays(-7)    
  $children = Get-Childitem -Path $path -Recurse -Include @($files) | Where-Object { $_.CreationTime -le $date }
  if ($children) {
    foreach($child in $children) {
      rm $child -ErrorAction SilentlyContinue
    }        
  }
}

# Wait for file to be unlocked
function WaitForFile([string]$fullPath) {
  [int]$numTries = 0
  $success = $true
  while ($true) {
    ++$numTries
    try {
      # Check if the file is locked.
      $filelocked = $false
      $fileInfo = New-Object System.IO.FileInfo $fullPath
      $fileStream = $fileInfo.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
      if ($fileStream) {
        $fileStream.Close()
      }            
      # If we got this far the file is ready            
      break            
    }
    catch {
      if ($numTries -gt 10) {
        Write-Error -Message "Quitting after $numTries attempts to gain lock" -Category WriteError
        $success = $false;
        break
      }
      # Wait for the lock to be released
      Write-Host "Waiting for lock release"
      if ($fileStream) {
        $fileStream.Close()
      }            
      Start-Sleep -Seconds 5
    }
  }
  return $success
}

function GetSeriesId($seriesName) {
  # Get the Series ID based off the series name provided.
  try {
    Write-Host "Getting series id from theTVDb.com using series name for $seriesname."
    $url=$tvdburl+"api/GetSeries.php?seriesname="+$seriesName+"&language=English"
    [xml]$series_ws = (New-Object System.Net.WebClient).DownloadString($url)
    if ($series_ws) {
      # Get series id based on name.
      foreach ($Series in $series_ws.Data.Series){
         if ($series.SeriesName -eq $seriesName){
          # Set the global series id variable.
          $seriesId = $series.seriesid
          break
        } # end if $series.SeriesName equal to $seriesName
      } # end foreach $Series in $series_ws.Data.Series.
    } # end if $series_ws.
  } catch {
    Write-Error -Message "Unable to find $seriesName on www.thetvdb.com" -Category ConnectionError
    break
  } # end try catch.
  return $seriesId
}

function GetBannerImage($seriesId, $parentDir, $season) {
  $imgFile = $parentDir+'\'+$seriesId+'.jpg'
  
  if (!(Test-Path $imgFile)) {
    # Get the highest rated season banner image.
    $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/banners.xml"
    [xml]$banners_ws = (New-Object System.Net.WebClient).DownloadString($url)
    if ($banners_ws) {
      $bannerData = $banners_ws.Banners.Banner
      [int]$rating = 0
      [int]$ratingCount = 0
      $rated = $false
      # Get the highest ranked banner image for this season.
      Write-Host "Searching for season image..."
      foreach ($banner in $bannerData) {
        if ($banner.BannerType -eq 'season') {
          if ($banner.Season) {
            if (($banner.Season -eq [int]$season) -and ($banner.Language -eq 'en')){
              Write-Host "Matched season"
              # TODO fix this rating count comparison so it works.
              if ($banner.RatingCount -ne '0') {
                if ([int]$banner.RatingCount -gt $ratingCount) {
                  if ([int]$banner.Rating -gt $rating) {
                    $rating = $banner.Rating
                    Write-Host "Rating: $rating"
                    $ratingCount = $banner.RatingCount
                    Write-Host "Rating count: $ratingCount"
                  }
                }
              }
            }
          }
        }
      }
      foreach ($banner in $bannerData) {
        if (($banner.BannerType -eq 'season') -and ($banner.Season -eq [int]$season) -and ([int]$banner.Rating -eq $rating) -and ([int]$banner.RatingCount -eq $ratingCount)) {
          $bannerPath = $banner.BannerPath
          Write-Host "Banner path: $bannerPath"
          $seasonImage = $url=$tvdburl+"banners/"+$bannerPath
          break                
        }
      }
    }
    # Download banner image
    if ($seasonImage) {
      $image_wc = New-Object System.Net.WebClient
      Write-Host "Downloading season image."
      $image_wc.DownloadFile($seasonImage, $imgFile)
    }
  }
  return $imgFile
}

# Rename the provided series with information gathered from theTVDb.
function RenameAndTagEpisode($date,$episode,$episode2,$file,$season,$seriesName,$recording,$psDir) {
  $atomicParsley = $dropBox+'\Applications\Open Source\Video Apps\AtomicParsley\AtomicParsley.exe'
  $parentDir = Split-Path $file -parent
  Write-Host "Parent directory: $parentDir"
  $itunesPath = GetiTunesPath
  if ($itunesPath) {
    $iTunesAddFolder = $itunesPath+'\Automatically Add to iTunes\'
    Write-Host "Automatically add to iTunes folder path: $iTunesAddFolder"
  } else {
    Write-Error -Message "Unable to locate iTunes folder path"
    break
  } # end if else $itunesPath
  
  $secondEp = $false    
  if ($seriesName) {
    $seriesId = GetSeriesId $seriesName
    Write-Host "Series ID: $seriesId"
    if ($seriesId) {
      try {
        # Get the base series record.
        $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/en.xml"
        [xml]$baseseries_ws = (New-Object System.Net.WebClient).DownloadString($url)
        if ($baseseries_ws) {
          $baseSeriesData = $baseseries_ws.Data.Series
          $actors = $baseSeriesData.Actors
          Write-Host "Actors: $actors"
          $contentRating = $baseSeriesData.ContentRating
          Write-Host "Rating: $contentRating"
          $genre = $baseSeriesData.Genre
          Write-Host "Genre: $genre"
          $network = $baseSeriesData.Network
          Write-Host "Network: $network"
        } # end if $baseseries_ws.
      } catch {
        Write-Error -Message "Unable to get series data for $seriesName from the www.thetvdb.com" -Category ConnectionError
        exit
      } # end try catch.
    }
  } else {
    Write-Error -Message  "No series name provided" -Category NotSpecified
    break
  } # end if else $seriesName.

  # If series id is not there, exit out.
  if ($seriesId) {
    # If season and episode provided (i.e. torrented shows), tag based on that information.
    if ($season -and $episode) {
      $seasonId = [int]$season
      $episodeId = [int]$episode
      if ($episode2) {
        $episode2Id = [int]$episode2
      } #end if $episode2
      try {
        Write-Host "Downloading episode data from theTVDb.com"
        $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/default/"+$seasonId+"/"+$episodeId+"/en.xml"
        [xml]$episode_ws = (New-Object System.Net.WebClient).DownloadString($url)
        if ($episode_ws) {
          $episodeData = $episode_ws.Data.Episode
          $episodeName = $episodeData.EpisodeName
          Write-Host "Episode name: $episodeName"
          $description = $episodeData.Overview
          Write-Host "Episode description: $description"
          $airDate = $episodeData.FirstAired
          Write-Host "Air date: $airDate"
        } # end if $episode_ws.
        if ($episode2Id) {
          Write-Host "Downloading episode 2 data from theTVDb.com"
          $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/default/"+$seasonId+"/"+$episode2Id+"/en.xml"
          [xml]$episode2_ws = (New-Object System.Net.WebClient).DownloadString($url)
          if ($episode2_ws) {
            $episode2Data = $episode2_ws.Data.Episode
            $episode2Name = $episode2Data.EpisodeName
            Write-Host "Episode 2 name: $episode2Name"
            $description = $description+' '+$episode2Data.Overview
            Write-Host "Combined description: $description"
          } # end if $episode2_ws.
        } # end if $episode2Id.
      } catch {
        Write-Error -Message "Unable to download episode data from theTVDb.com" -Category ConnectionError
        break
      }
    } else {
      # If no episode or series data provided (i.e. recorded shows), tag by provided airDate.
      $url=$tvdburl+"api/GetEpisodeByAirDate.php?apikey="+$tvdbAPIKey+"&seriesid="+$seriesId+"&airdate="+$date+"&[language=English]"
      [xml]$episode_ws = (New-Object System.Net.WebClient).DownloadString($url)
      if ($episode_ws) {
      $episodeData = $episode_ws.Data.Episode
      $episodeName = $episodeData.EpisodeName
        $description = $episodeData.Overview
        $airDate = $episodeData.FirstAired
        Write-Host "Air date: $airDate"
        $episode = [int32]::Parse($episodeData.EpisodeNumber)
        Write-Host "Episode: $episode"
        $season = [int32]::Parse($episodeData.SeasonNumber)
        Write-Host "Season: $season"
      } # end if $episode_ws.
    } # end if else $season and #episode.

    # Add leading zeros to season and episode numbers.
    if (($season -lt 10) -and ($season.Length -lt 2)) {
      $season = '0'+$season
    } # end if $season less than 10, and shorter than 2 characters.
    if (($episode -lt 10) -and ($episode.Length -lt 2)) {
      $episode = '0'+$episode
    } # end if $episode less than 10, and shorter than 2 characters.
    if ($episode2 -and ($episode2 -lt 10) -and ($episode2.Length -lt 2)) {
      $episode2 = '0'+$episode2
    } # end if $episode exists and is less than 10, and shorter than 2 characters.

    # Clean unprintable characters.
    $episodeName = $episodeName.Replace(':','').Replace('?','').Trim()
    if ($episode2Name) { 
      $episode2Name = $episode2Name.Replace(':','').Replace('?','').Trim()
    } # end if $episode2Name.

    # If double-episode, rename with both names, else just the one.
    if ($episode2Id) {      
      $fileFormat = $seriesName+" - S"+$season+"E"+$episode+"-E"+$episode2+" - "+$episodeName+" & "+$episode2Name+".mp4"    
    } else {
      $fileFormat = $seriesName+" - S"+$season+"E"+$episode+" - "+$episodeName+".mp4"    
    } # end if else $episode2Id.
    Write-Host "Formatted file name: $fileFormat"
    try {
      # Rename the file
      sl -LiteralPath $parentDir
      Write-Host "Renaming $file to $fileFormat"
      Move-Item -Path $file $fileFormat -Force -Verbose
    } catch {
      Write-Error -Message "Unable to rename $file to $fileFormat, an error has occurred." -Category WriteError
    } # end try catch.
  } # end if $seasonId.

  # If not renamed, use original name format.
  if (!($fileFormat)) {
    $fileFormat = (Split-Path $file -Leaf -Resolve)
    Write-Host "Formatted file name: $fileFormat"
  } # end if not $fileFormat.
  Write-Host "Deleting existing metadata from $fileFormat."
  $fullFilePath = $parentDir+'\'+$fileFormat
  Write-Host "File path: $fullFilePath"
  & $atomicParsley $fullFilePath --metaEnema --overWrite | Write-Host -ErrorAction SilentlyContinue
  # If atomic parsley fails due to file errors, re-encode the file using Handbrake.
  if ($LastExitCode -ne 0) {
    Write-Host "An error occurred attempting to wipe $fullFilePath of metatags.  Re-encoding file."
    $newFile = 'C:\Users\Steve\Downloads\Torrents\Tagging\'+(Split-Path $fullFilePath -Leaf -Resolve).Replace('.ts','.mp4')
    Write-Host "New file: $newFile"
    EncodeWithHB $fullFilePath $newFile
    Remove-Item $fullFilePath -Force -Verbose
    return
  }
  Start-Sleep -Seconds 5

  # Download the series image file.
  if ($seriesId -and $parentDir -and $season) {
    $imgFile = GetBannerImage $seriesId $parentDir $season
    Write-Host "Image file: $imgFile"
  }
  
  # Set up variables for tagging.
  $episodeDesc = [string]$episode+[string]$episode2+' - '+$episodeName
  
  $title = $fileFormat.Replace('.mp4','')
  $album = $seriesName+', Season '+[int]$season
  $actors = $actors.Replace('|',', ').Trim(',',' ')

  # Gather Parameters
  $parsleyParams = @($fullFilePath)
  $genre = $genre.Replace('|',', ').Trim(',',' ')
  if (!($genre)) { 
    $genre = 'TV Show' 
  } # end if !$genre.
  $parsleyParams += @('--genre',$genre,'--stik','TV Show')
  if (!($actors)) { 
    $actors = 'N/A' 
  } # end if !$actors.
  $parsleyParams += @('--artist',$actors)
  if ($airDate) {
    $airDate += 'T12:00:00Z'
  } elseif ($date) {
    $airDate = $date+'T12:00:00Z'
  } else {
    $airDate = (Get-Date -Format s)+'Z'
  } # end if elseif else $airdate.
  $parsleyParams += @('--year',$airDate)
  if ($episodeDesc) {
    $parsleyParams += @('--TVEpisode',$episodeDesc)
  } # end if $episodeDesc.
  if (!($episodeId)) {
    $episodeId = [int]$episode
  } # end if !$episodeId.
  $parsleyParams += @('--TVEpisodeNum',$episodeId,'--tracknum',$episodeId)
  if(!($seasonId)) {
    $seasonId = [int]$seasons
  } # end if !$seasonId.
  $parsleyParams += @('--TVSeason',$seasonId)
  if ($contentRating) {
    $parsleyParams += @('--contentRating',$contentRating)
  } # end if $contentRating.
  if ($title) {
    $parsleyParams += @('--title',$title)
  } # end if $title
  if ($seriesName) {
    $parsleyParams += @('--albumArtist',$seriesName,'--TVShowName',$seriesName)
  } # end if $seriesName.
  if ($album) {
    $parsleyParams += @('--album',$album)
  } # end if $album.
  if ($network) {
    $parsleyParams += @('--TVNetwork',$network)
  } # end if $network.
  if ($description) {
    $parsleyParams += @('--description',$description,'--longdesc',$description)
  } # end if $description.
  if ($imgFile) {
    $parsleyParams += @('--artwork',$imgFile)
  } # end if $imgFile
  $parsleyParams += @('--overWrite')
      
  # Tag the file.
  Write-Host "Writing tags to $fileFormat"
  & $atomicParsley $parsleyParams | Write-Host
  Start-Sleep -Seconds 5
  
  # Delete the image file.
  Remove-Item $imgFile -ErrorAction SilentlyContinue -Verbose

  Write-Host "Destination: $iTunesAddFolder"
  # Move the file to provided destination.
  try {  
    if (Test-Path $fileFormat) {
      Move-Item -Path $fileFormat -Destination $iTunesAddFolder -Force -Verbose
    } # end if Test-Path $fileFormat.
  } catch {
    Write-Error -Message "Unable to move $fileFormat to $iTunesAddFolder, an error has occurred." -Category WriteError
  } # end try catch.
}

# Delete any watched episodes of TV from PC etc.
function DeleteWatchedTV {
  $iTunes = New-Object -ComObject iTunes.application
  # Start-Sleep -Seconds 30
  if ($iTunes) {
    $LibrarySource = $iTunes.LibrarySource
    $playList = $LibrarySource.Playlists | Where-Object { $_.Name -eq "TV Shows" }  
    $tracks = $playList.Tracks
    $trackCount = $tracks.Count
    $counter = 0
    $showsFound = 0
    $removed = @()
    foreach ($track in $tracks) {
      # Reset variables
      $parent = $null
      $children = $null
      ++$counter
      # Update progress bar.
      Write-Progress -activity "Searching for TV Shows to delete. . ." -status "Read: $counter of $trackCount" -percentComplete (($counter / $trackCount) * 100)
      $path = $track.Location
      # Delete the track from iTunes.
      $track.Delete()
      Write-Host "$name removed from iTunes."
        
      # If the file is actually there, delete it.
      if (Test-Path $path) {
        rm -Path $path -Force
        Write-Host "$path deleted from the computer."       
      }
      # If the parent folder is empty, delete that too.
      $parent = (Split-Path $path -parent)
      $children = Get-ChildItem -Path $parent -Recurse
      if ($children -eq $null) {
        try {
          rm $parent -Force
          Write-Host "$parent deleted from PC"
        } catch {
          Write-Error -Message "Unable to delete $parent" -Category ResourceBusy
        }
      }         
    }
    # Release COM objects.
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($itunes)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($LibrarySource)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($playList)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($track)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tracks)
  }
}

#################################################################################
# VARIABLES                                                                     #
#################################################################################

$uTorrentPW = 'torrents'
$uTorrentUser = 'jpicker'
$uTorrentWebUI = 'http://localhost:8080/gui'
$uTorrentHome = $env:USERPROFILE+"\Downloads\Torrents\"
$dropFolder = $uTorrentHome+'Tagging\'
$dropBox = GetDropBox
$powershellDir  = $env:USERPROFILE+"\My Documents\WindowsPowerShell\"
[int]$runthrough = 1
$exitScript = $false
$endBracket = $null
$folderRename = $null
$length = $null
$retries = 10
$seeding = 4,5,7
$startBracket = $null
$timeout = 30
$unpacked = $uTorrentHome+'Unpacked\'
$uTorrentPath = $null

# Load uTorrentAPI assembly
Add-Type -Path "C:\Windows\Microsoft.NET\Framework\uTorrentAPI\Release\AnyCPU\UTorrentAPI.dll" 
$utorrentclient = New-Object -TypeName UTorrentAPI.UTorrentClient -ArgumentList $uTorrentWebUI,$uTorrentUser,$uTorrentPW 
$torrents = $utorrentclient.Torrents 

# If torrent folder doesn't exist, create it.
if (!(Test-Path $uTorrentHome)){
  md $uTorrentHome    
}
#################################################################################
# PROCESS FLOW                                                                  #
#################################################################################

do { 
  # If the script is run manually, use the drop folder to tag dropped files.
  if (($directory -eq $null) -and ($kind -eq $null)) {
    $directory = $dropFolder
    Write-Host "Directory: $directory"
    $kind = 'multi'
  }
  # If directory is missing closing slash, add it.
  If ($directory[-1] -ne "\") {
    $directory += "\"
  }
  $length = $null
  # Strip bracketed website descriptions from title and directory if they exist.
  if ($title.Substring(0,1) -eq "[") {
    Write-Host "Stripping brackets from torrent title."
    $endBracket = $title.IndexOf("]")+1
    $length = $title.Length-$endBracket
    $title = $title.SubString($endBracket,$length).Replace(' - ','').TrimStart()
    # Get directory position of brackets for removal.
    $startBracket = $directory.IndexOf("[")
    if ($startBracket -ge 0) {
      Write-Host "Stripping brackets from torrent folder."
      $endBracket = $directory.IndexOf("]")+1
      $length = $directory.Length-$endBracket
      $leftString = $directory.Substring(0,$startBracket)
      $rightString = $directory.Substring($endBracket,$length).Replace(' - ','').TrimStart()
      $folderRename = $leftString+$rightString
      $uTorrentPath = $directory
      copy -LiteralPath $directory -Recurse $folderRename -Force -Verbose
      $directory = $folderRename
      Write-Host "Directory: $directory"
    }
  }
  if ($kind -eq 'multi'){
    # wait for directory to be available.
    $wait = 0
    Write-Host "Testing if $directory exists.  If not, wait until it does."
    do {
      Test-Path $directory -Verbose
      Start-Sleep 1
      $wait = $wait+1      
    } while (!$directory -or ($wait -eq $timeout))
    sl -LiteralPath $directory
    # Find all uncompressed mp4 & m4v files and move them to unpacked folder.
    $files = ls *.mp4,*.m4v -Recurse
    foreach ($file in $files) {
      if ($file.length -gt 100mb) {
        if ($directory -eq $dropFolder) {
          Write-Host "Moving $file to $unpacked"
          Move-Item -literalPath $file $unpacked -Force -Verbose
        } else {
          Write-Host "Copying $file to $unpacked"
          Copy-Item -LiteralPath $file $unpacked -Force -Verbose
        }
      }
    }
    # Find all mkv & avi files, convert them to mp4 and move them to the unpacked folder
    $files = ls *.mkv,*.avi -Recurse
    if ($files) {
      foreach ($file in $files) {
        $output = $unpacked+$title+".mp4"
        Write-Host "Output: $output"
        EncodeWithHB $file $output 
      }
    }
    # Find if there are any archived files within the folder, if so, extract them.
    $files = ls *.rar -Recurse
    if ($files) {
      foreach ($file in $files) {
        # Use unRarFiles(filepath, destination) function to extract archive
        UnRarFiles $directory $unpacked $file 
      }      
    }
  } elseif ($kind -eq 'single') {
    # Wait 10 seconds to assure files have been copied across.
    Start-Sleep -Seconds 10
    # Move files to unpacked folder.
      sl -LiteralPath $directory
      if ($torrentFile.EndsWith('mp4') -or $torrentFile.EndsWith('m4v')) {
        $filePath = $directory+$torrentFile
        if (Test-Path -LiteralPath $filePath) {
          Copy-Item -LiteralPath $filePath $unpacked -Force 
        }
      }
    # If single is an mkv file, convert it to mp4 and move it to the unpacked folder
    $files = ls *.mkv -Recurse
    if ($files) {
      foreach ($file in $files) {
        $input = $directory+$file
        $output = $unpacked+$file.Replace("mkv","mp4")
        # Use encodeWithHB(filepath, destination) function encode the file with Handbrake.                                     
        EncodeWithHB $input $output 
      }
    }
  }
  Start-Sleep -Seconds 10 -Verbose
  sl -LiteralPath $unpacked
  $files = ls *.mp4
  # If files are found, continue.
  if ($files) {
    # For each file in unpacked folder do the following.
    foreach ($file in $files) {
      # If the file has  website name at beginning, rename it.
      if ($file.Name.Substring(0,3) -eq "www") {
        Write-Host "Stripping website name from filename."
        $dotCom = $file.IndexOf(".com")+4
        $length = $file.Length-$dotCom
        $fileRename = $file.Name.Substring($dotCom,$length).Replace(' - ','').TrimStart()
        Move-Item -literalPath $file.FullName $fileRename -Verbose
        # If the file has any brackets, rename the file without them.
        $newFile = $fileRename.FullName.Replace("[","").Replace("]","")      
        Move-Item -literalPath $fileRename.FullName $newFile -Verbose
      } else {
        # If the file has any brackets, rename the file without them.
        $newFile = $file.FullName.Replace("[","").Replace("]","")
        if ($file.FullName -ne $newFile) {
          Move-Item -literalPath $file.FullName $newFile -Verbose
        }
      }
      # Initilise Do/While loop variables.
      [int]$retry = 0
      $exitLoop = $false
      # Try renaming file, if file in use wait 30 seconds and try again.
      do {
        try {
          # Remove read-only permission on file.
          Write-Host "Removing read-only property from $newFile"
          sp $newFile IsReadOnly $false -Verbose
          # TODO: if ($isTV) {       
          # Loop through filename array and pull out the series name.
          $seriesName = $null
          $season = $null
          $episode = $null
          $episode2 = $null
          $date = (Get-Date).ToUniversalTime().AddHours(-4).ToString("yyyy-MM-dd")
          Write-Host "Date: $date"
          $titleArray = (Split-Path $newFile -Leaf -Resolve).Split('.-_')
          Write-Host "Reading season and episode numbers"
          foreach ($item in $titleArray) {
            $item = $item.Trim()
            if ($item.ToUpper() -Match 'S(\d{1,2})E(\d{1,2})E(\d{1,2})') {
              $details = $item.Split('sSeEeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              $episode2 = $details[3].TrimStart()              
              break
            } elseif ($item.ToUpper() -Match 'S(\d{1,1})E(\d{1,2})') {
              $details = $item.Split('sSeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              break
            } elseif ($item.ToUpper() -Match 'S(\d{1,2})E(\d{1,2})') {
              $details = $item.Split('sSeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              break
            } elseif ($item.ToUpper() -Match 'S(\d{1,2})') {
              $details = $item.Split('sS')
              $season = $details[1].TrimStart()
              break
            } elseif ($item.ToUpper() -Match '(\d{1,2})X(\d{1,2})') {
              $details = $item.Split('xX')
              $season = $details[0].TrimStart()
              $episode = $details[1].TrimStart()
              break
            } elseif ($item -Match '(^(19|20)\d{2}$)') {
              # Add brackets around years in series name.
              $seriesName = $seriesName+" "+"("+$item+")"
            } elseif ($item -Match '(^\d{1,3}$)') {
              $season = '0'+$item.Substring(0,1)
              $episode = $item.Substring(1,2)
              break
            } elseif ($item.ToUpper() -Match '(^\d{8}$)') {
              $date = $item.Substring(0,4)+'-'+$item.Substring(4,2)+'-'+$item.Substring(6,2)
              break
            } else {
              if ($seriesName -eq $null){
                $seriesName = $item
              } else {
                $seriesName = $seriesName+" "+$item
              }
            }            
          }
          Write-Host "Series Name: $seriesName"
          Write-Host "Season: $season"
          Write-Host "Episode: $episode"
          Write-Host "Episode 2: $episode2"
          
          # If the file doesn't contain a series name, exit the script.
          if ($seriesName -eq $null) {
            break
          }
          # Convert Series Name to Title Case
          $seriesName = (Get-Culture).TextInfo.ToTitleCase($seriesName.ToLower()).Trim()
          Write-Host "Series Name (Title Case): $seriesName"
          # A few specific replace Strings for specific problem shows.  
          $seriesName = $seriesName.Replace("Greys","Grey's").Replace("Its","It's")
          Write-Host "Series Name (Corrected): $seriesName"
          # Tag, rename & move file to correct folder location.
          # function RenameAndTagEpisode($date,$episode,$episode2,$file,$season,$seriesName)
          RenameAndTagEpisode $date $episode $episode2 $newFile $season $seriesName $false $powershellDir -Verbose
          $exitLoop = $true
        } catch {
          if ($retry -eq $retries) {
            $exitLoop = $true
            Write-Error -Message "Failed to rename "+$newFile+" after "+$retries+" retries."
            return $false
          } else {
            Start-Sleep -Seconds 30 -Verbose
            $retry = $retry+1
          }
        }
      } while ($exitLoop -eq $false) 
    }
  }
  if ($uTorrentPath) {
    Write-Host "Deleting $directory"
    del -Recurse $directory -Force -Verbose
  }
  # Reset variables to run through script a second time to get any files that have been re-encoded.
  $directory = $null
  $kind = $null
  if ($runthrough -ge 2) {
    $exitScript = $true
  }
  ++$runthrough
} while ($exitScript -eq $false) 
