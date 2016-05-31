################################################################################
#                                                                              #
# This script gets passed some arguments from uTorrent, which it uses to       #
# rename, tag then move to the Media drive and add to Plex.                    #
#                                                                              #
# Created Date: 26 SEP 2013                                                    #
# Version: 4.0                                                                 #
#                                                                              #
################################################################################

################################################################################
# FUNCTIONS                                                                    #
################################################################################

function refreshPlex() {
  $plex = "C:\Program Files (x86)\Plex\Plex Media Server\Plex Media Scanner.exe"
  $plexParams = '--scan','--section','2'
  [Diagnostics.Process]::Start($plex,$plexParams).WaitForExit()
  # Start-Process -FilePath $plex -ArgumentList $plexParams -Wait -NoNewWindow
}

# Encode with ffmpeg
function EncodeWithFFMPEG($inputStr,$outputStr) {
  $ffmpeg = 'D:\Dropbox\Applications\Open Source\Video Apps\ffmpeg\bin\ffmpeg.exe'
  $params = '-i',$inputStr,$outputStr
  try {
    Write-Host "Encoding $inputStr to $outputStr..."
    [Diagnostics.Process]::Start($ffmpeg, $params).WaitForExit()
  } catch {
    "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
    "Exception: "+$_.Exception.Message >> $errorLog
    "Item: "+$inputStr >> $errorLog
    if ([bool](get-process  notepad -ea 0)) {
      Stop-Process notepad 
      Invoke-Item $errorLog
    }
    return $false
  }
  return $true
}

# Encode the specified file with HandBrake for AppleTV.
function EncodeWithHB($inputStr,$outputStr){

  $handbrake = "C:\Program Files\Handbrake\HandBrakeCLI.exe"
  # Add quoted paths to Handbrake (as without them it breaks)
  $inputStr = '"'+$inputStr+'"'
  $outputStr = '"'+$outputStr+'"'
  # Set up handbrake prameters.
  $params = '-i',$inputStr,'-t','1','--angle','1','-c','1','-o',$outputStr,'-f',
    'mp4','-o','--decomb','-w','720','--crop','0:0:0:0','--modulus','2','-e','x264','-b',
    '1000','--vfr','-a','1','-E','faac,copy:ac3','-6','dpl2,auto','-R',
    'Auto,Auto','-B','256,0','-D','0,0','--gain','-2,0','--audio-fallback',
    'ffac3','--x264-profile=high','--h264-level="4.1"','--verbose=1'

  try {
    Write-Host "Encoding $inputStr to $outputStr..."
    [Diagnostics.Process]::Start($handbrake, $params).WaitForExit()
  } catch {
    "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
    "Exception: "+$_.Exception.Message >> $errorLog
    "Item: "+$inputStr >> $errorLog
    if ([bool](get-process  notepad -ea 0)) {
      Stop-Process notepad 
      Invoke-Item $errorLog
    }
    return $false
  }
  return $true
}

# Unpack the specified compressed files to the output directory provided.
function UnRarFiles($torrentDirectory,$outputDir,$file) {

  $unRar = "C:\Program Files (x86)\Unrar\UnRAR.exe"
  if (!(Test-Path $outputDir)) {
    Write-Host "Creating directory $outputDir"
    md $outputdir
  }
  sl -LiteralPath $torrentDirectory
  $unRarParams = 'x',$file,'*.mp4',$outputDir,'-o+'
  try {
    Write-Host "Upacking $torrentDirectory to $outputDir"
    [Diagnostics.Process]::Start($unrar, $unRarParams).WaitForExit()
  } catch {
    "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
    "Exception: "+$_.Exception.Message >> $errorLog
    "Item: "+$file >> $errorLog
    if ([bool](get-process  notepad -ea 0)) {
      Stop-Process notepad 
      Invoke-Item $errorLog
    }
    return $false
  }
}

function GetSeriesData($showName, [int]$season, [int]$episode, [int]$episode2) {
  $tvdbAPIKey = $env:TVDB_API_KEY
  $tvdburl = "http://thetvdb.com/"
  if ($showName -and $season -and $episode) {
    $continue = $false
    # Get the Series data based off the series name provided.
    try {
      $lastSeriesId = 0    
      Write-Host "Getting series id from theTVDb.com using series name for $showName."
      $url=$tvdburl+"api/GetSeries.php?seriesname="+$showName+"&language=English"
      try {
        [xml]$series_ws = $wc.DownloadString("$url")
      } catch [System.Net.WebException] {
        $statusCode = [int]$_.Exception.Response.StatusCode
        $html = $_.Exception.Response.StatusDescription
        "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
        "Error: "+$statuscode+" : "+$html >> $errorLog
      }
      $showName = ($showName+"*").Replace(' ','*')
      # Get series id based on name.
      $matches = @()
      foreach ($Series in $series_ws.Data.Series) {
        $tempName = ($series.SeriesName).Replace("'","")
        if ($tempName -like $showName){
          # Set the global series id variable.
          [array]$matches += $series      
        }
      }
      if ($series_ws.Data.HasChildNodes -eq $true) {   
        while (!$continue) {       
          $showData = @{}
          $releaseYear = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
          foreach ($match in $matches) {
            if (((Get-date $match.FirstAired) -gt $releaseYear) -and $match.seriesid -ne $lastSeriesId) {
              $releaseYear = Get-date $match.FirstAired 
              $seriesId = $match.seriesid
            }
          }
          # Get the base series record.
          $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/en.xml"
          try {
            [xml]$baseseries_ws = $wc.DownloadString("$url")
          } catch [System.Net.WebException] {
            $statusCode = [int]$_.Exception.Response.StatusCode
            $html = $_.Exception.Response.StatusDescription
            "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
            "Error\: $statuscode\:$html" >> $errorLog          
          }        
          if ($baseseries_ws.Data.HasChildNodes -eq $true) {
            $baseSeriesData = $baseseries_ws.Data.Series
            $showData.Add("SeriesId", $baseSeriesData.id)
            Write-Host "Series ID: "$baseSeriesData.id
            $showData.Add("SeriesName", $baseSeriesData.SeriesName)
            Write-Host "Show: "$baseSeriesData.SeriesName
            $showData.Add("Actors",$baseSeriesData.Actors)
            Write-Host "Actors: "$baseSeriesData.Actors
            $showData.Add("ContentRating",$baseSeriesData.ContentRating)
            Write-Host "Rating: "$baseSeriesData.ContentRating
            $genre = $baseSeriesData.Genre
            $showData.Add("Genre", $baseSeriesData.Genre)
            Write-Host "Genre: "$baseSeriesData.Genre
            $showData.Add("Network", $baseSeriesData.Network)
            Write-Host "Network: "$baseSeriesData.Network
            # Download episode data.
            Write-Host "`nDownloading episode data from theTVDb.com"
            $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/default/"+$season+"/"+$episode+"/en.xml"
            try {
              [xml]$episode_ws = $wc.DownloadString("$url")
              if ($episode_ws.Data.HasChildNodes -eq $true){
                $episodeData = $episode_ws.Data.Episode
                $episodeName = $episodeData.EpisodeName -replace '[/</>/:/"//\\/|/?/*]', "_"
                $showData.Add("EpisodeName", $episodeName)
                Write-Host "Episode name: "$episodeData.EpisodeName
                $showData.Add("Synopsis", $episodeData.Overview)
                Write-Host "Episode description: "$episodeData.Overview
                $showData.Add("AirDate", $episodeData.FirstAired)
                Write-Host "Air date: "$episodeData.FirstAired
                $continue = $true
                if ($episode2) {
                  Write-Host "Downloading episode 2 data from theTVDb.com"
                  $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/default/"+$season+"/"+$episode2+"/en.xml"
                  try {
                    [xml]$episode2_ws = [System.Web.HttpUtility]::HtmlDecode($wc.DownloadString("$url"))
                  } catch [System.Net.WebException] {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                    # Ignore 404 error.  We want to retry in this instance.
                    if ($statusCode -ne 404) {
                      $html = $_.Exception.Response.StatusDescription
                      "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                      "Error\: $statuscode\:$html" >> $errorLog
                    }
                  }
                  if ($episode2_ws.Data.HasChildNodes -eq $true) {
                    $episode2Data = $episode2_ws.Data.Episode
                    $episode2Name = $episode2Data.EpisodeName -replace '[/</>/:/"//\\/|/?/*]', "_"
                    $showData.Add("Episode2Name",$episode2Name)
                    Write-Host "Episode 2 name: $episode2Name"
                    $episode1Synopsis = $showData.get_Item("Synopsis")
                    $synopsis = $episode1Synopsis+' & '+$episode2Data.Overview
                    $showData.set_Item("Synopsis", $synopsis)
                    Write-Host "Combined synopsis: $synopsis"
                  } else {
                    "No data found for Episode: "+$episode2.ToString()+" of season "+$season.ToString()+" for "+$showName+" on the TVBD. Exiting." >> $errorLog                  
                  }                  
                }
              } else {
                "No data found for Episode: "+$episode.ToString()+" of season "+$season.ToString()+" for "+$showName+" on the TVBD. Exiting." >> $errorLog
                return $false
              }
            } catch [System.Net.WebException] {
              $lastSeriesId = $seriesId
              Write-Host "Incorrect series with ID: "$lastSeriesId". Retrying."
              $statusCode = [int]$_.Exception.Response.StatusCode
              # Ignore 404 error.  We want to retry in this instance.
              if ($statusCode -ne 404) {
                $html = $_.Exception.Response.StatusDescription
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Error\: $statuscode\:$html" >> $errorLog
              }
            }
          } else {          
            "No data found for Series with Series ID: "+$seriesId+" on the TVBD. Exiting." >> $errorLog
            return $false
          }    
        } 
        return $showData
      } else {
        "No Show called "+$ShowName+" found on The TVDB. Exiting." >> $errorLog
        return $false
      }  
    } catch {
      "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
      "Exception: "+$_.Exception.Message >> $errorLog
      "Item: "+$showName >> $errorLog
      if ([bool](get-process  notepad -ea 0)) {
        Stop-Process notepad 
        Invoke-Item $errorLog
      }
      return $false
    }  
  } else {
    "Not enough information provided to complete search.  `nNeed show name, season and episode number. Exiting." >> $errorLog
    return $false  
  }
}

function GetBannerImage($seriesId, $parentDir, $season) {
  $imgFile = $parentDir+$seriesId+'.jpg'  
  if (!(Test-Path $imgFile)) {
    # Get the highest rated season banner image.
    $url=$tvdburl+"api/"+$tvdbAPIKey+"/series/"+$seriesId+"/banners.xml"
    [xml]$banners_ws = $wc.DownloadString("$url")
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

function DownloadFile($url, $targetFile = 'c:\temp\xml.xml') {
  $uri = New-Object "System.Uri" "$url"
  $request = [System.Net.HttpWebRequest]::Create($uri)
  $request.set_Timeout(15000) #15 second timeout
  $response = $request.GetResponse()
  $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
  $responseStream = $response.GetResponseStream()
  $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create -
  $buffer = new-object byte[] 10KB
  $count = $responseStream.Read($buffer,0,$buffer.length)
  $downloadedBytes = $count
  while ($count -gt 0) {
    $targetStream.Write($buffer, 0, $count)
    $count = $responseStream.Read($buffer,0,$buffer.length)
    $downloadedBytes = $downloadedBytes+$count
    Write-Progress -activity "Downloading file '$($url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)    
  }
  Write-Progress -activity "Finished downloading file '$($url.split('/') | Select -Last 1)'"
  $targetStream.Flush()
  $targetStream.Close()
  $targetStream.Dispose()
  $responseStream.Dispose()
}

################################################################################
# VARIABLES ETC                                                                #
################################################################################

$dropBox = "D:\Dropbox"
$aacgain = $dropBox+'\Applications\Open Source\Audio Apps\aacgain_1_9\aacgain.exe'
$destDrive = 'D:\'
$atomicParsley = $dropBox+'\Applications\Open Source\Video Apps\AtomicParsley\AtomicParsley.exe'
$enc = [system.Text.Encoding]::UTF8
$endBracket = $null
$exitScript = $false
$ffmpeg = $dropBox+'\Applications\Open Source\Video Apps\ffmpeg\bin\ffmpeg.exe'
$folderRename = $null
$length = $null
$retries = 10
$seeding = 4,5,7
$startBracket = $null
$timeout = 30
$tmdbApiKey = $env:TMDB_API_KEY
$tmdburl = "http://www.themoviedb.org/"
$uTorrentHome = $env:USERPROFILE+"\Downloads\Torrents\"
$completed = $uTorrentHome+'Completed\'
$unpacked = $uTorrentHome+'Unpacked\'
$untagged = $uTorrentHome+'Untagged\'
$toConvert = $uTorrentHome+'To Convert\'
$connectionLog = $uTorrentHome+"ConnectionIssues.txt"
$script:errorLog = $uTorrentHome+"TVShows_ErrorLog.txt"
$script:waitFolder = $uTorrentHome+"Untagged\"
$uTorrentPW = $env:UTORRENT_PW
$uTorrentUser = $env:UTORRENT_USER
$uTorrentWebUI = 'http://localhost:8080/gui'
$script:wc = New-Object System.Net.WebClient
$script:wc.Encoding = [System.Text.Encoding]::UTF8

if (([environment]::OSVersion.Version.Major -eq 6) -and ([environment]::OSVersion.Version.Minor -ge 2)) {
  $videoWidthVar = 301 
} else {
  $videoWidthVar = 285
}

# Load uTorrentAPI assembly
Add-Type -Path "C:\Windows\Microsoft.NET\Framework\uTorrentAPI\Release\AnyCPU\UTorrentAPI.dll" 
$utorrentclient = New-Object -TypeName UTorrentAPI.UTorrentClient -ArgumentList $uTorrentWebUI,$uTorrentUser,$uTorrentPW 
$torrents = $utorrentclient.Torrents 

# If torrent folders don't exist, create them.
if (!(Test-Path $uTorrentHome)){
  md $uTorrentHome    
}
if (!(Test-Path $unpacked)) {
  mkdir $unpacked
}
if (!(Test-Path $untagged)) {
  mkdir $untagged
}

################################################################################
# PROCESS FLOW                                                                 #
################################################################################

# Check if Google is available.
[int]$retry = 0
while ((!(Test-Connection -computer 8.8.8.8 -count 1 -quiet)) -and ($retry -lt $retries)) {
  Start-Sleep -Seconds 10
  $retry++
}
if ($retry -eq $retries) {
  $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss:ff"
  $timestamp+", Unable to establish internet connection.  Exiting.`n" >> $connectionLog
  exit
}

# Scan through torrents.
foreach ($torrent in $torrents) {
  # If unable to connect to the tvdb, exit the script.
  [int]$retry = 0
  while ((!(Test-Connection -computer 'thetvdb.com' -count 1 -quiet)) -and ($retry -lt $retries)) {
    Start-Sleep -Seconds 10 
    $retry++
  }
  if ($retry -eq $retries) {
    $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss:ff"
    $timestamp+", Unable to establish connection to theTVDb.com.  Exiting.`n" >> $connectionLog
    Invoke-Item $connectionLog
    exit
  }
  
  $torrentDirectory = $torrent.SavePath+"\"
  
  # If complete continue.
  if([long]($torrent.RemainingBytes) -eq 0) {
    if ($torrent.Name -match "(.*)([sS](\d{1,2})[eE](\d{1,2})|(\d{1,2})[xX](\d{1,2})|(?:^|\W)season(?:$|\W)(\d{1,2}))(.*)") {
      # Get number of files.
      $torrentFiles = @($torrent.Files | ? { 
        ($_.Path -notlike '*sample*' -and ($_.Path -like '*.mp4' -or $_.Path -like '*.m4v' -or $_.Path -like '*.mkv')) 
      }).Count
      foreach ($f in $torrent.Files) {      
        if($f.Path -notlike '*sample*') {
          $remove = $false
          if(($f.Path).Endswith('.rar')) {
            # Use unRarFiles(filepath, destination) function to extract archive
            UnRarFiles $torrentDirectory $unpacked $f.Path          
          } elseif(($f.Path).Endswith('.mp4') -or ($f.Path).Endswith('.m4v')) {
            $torrentPath = $torrentDirectory+$f.Path
            # Rename file using Torrent Title, strip brackets etc.
            if ($torrentFiles -gt 1) {
              $torrentName = (Split-Path $f.Path -leaf).ToLower() -replace "\.(mp4|m4v)$",""
            } else {
              $torrentName = $torrent.Name
            }      
            $newMp4Path = $unpacked+$torrentName+'.mp4'
            if (Test-Path -LiteralPath $torrentPath) {
              Copy-Item -LiteralPath $torrentPath $newMp4Path -Force -Verbose
              if (Test-Path -LiteralPath $newMp4Path) {
                $remove = $true
                if (Test-Path -LiteralPath $torrentDirectory) {
                  if($torrentDirectory -ne $completed) {
                    Write-Host "Deleting $torrentDirectory"
                    del -Recurse $torrentDirectory -Force
                  }
                }
              }
            }
          } elseif(($f.Path).Endswith('.mkv')) {
            $torrentPath = $torrentDirectory+$f.Path
            # Rename file using Torrent Title, strip brackets etc.
            if ($torrentFiles -gt 1) {
              $torrentName = (Split-Path $f.Path -leaf).ToLower() -replace "\.(mkv)$",""
            } else {
              $torrentName = $torrent.Name
            }      
            $newMp4Path = $unpacked+$torrentName+'.mp4'
            if (Test-Path -LiteralPath $torrentPath) {
              $ffmpegParams = '-i',$torrentPath,'-acodec','copy','-vcodec','copy',$newMp4Path
              Write-Host "Converting mkv to mp4"
              try {
                & $ffmpeg $ffmpegParams
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$newMp4Path >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              Start-Sleep -Seconds 5
              if (Test-Path -LiteralPath $newMp4Path) {
                $remove = $true
                if (Test-Path -LiteralPath $torrentDirectory) {
                  if($torrentDirectory -ne $completed) {
                    Write-Host "Deleting $torrentDirectory"
                    del -Recurse $torrentDirectory -Force
                  }
                }
              }
            }
          } elseif(($f.Path).Endswith('.avi')) {
            $torrentPath = $torrentDirectory+$f.Path
            if ($torrentFiles -gt 1) {
              $torrentName = (Split-Path $f.Path -leaf).ToLower() -replace "\.(avi)$",""
            } else {
              $torrentName = $torrent.Name
            }      
            $newMp4Path = $unpacked+$torrentName+'.mp4'
            if (Test-Path -LiteralPath $torrentPath) {
              $ffmpegParams = '-i',$torrentPath,'-acodec','copy','-vcodec','libx264','-preset','slow',$newMp4Path
              Write-Host "Converting avi to mp4"
              try {
                & $ffmpeg $ffmpegParams
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$f.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              Start-Sleep -Seconds 5
              if (Test-Path -LiteralPath $newMp4Path) {
                $remove = $true
                if (Test-Path -LiteralPath $torrentDirectory) {
                  if($torrentDirectory -ne $completed) {
                    Write-Host "Deleting $torrentDirectory"
                    del -Recurse $torrentDirectory -Force
                  } 
                } 
              }
            }
          }
          if ($remove -eq $true) {
            $utorrentclient.Torrents.Remove($torrent, [UTorrentAPI.TorrentRemovalOptions]::TorrentFileAndData)
          }       
        } 
      }
    } else { 
      "Failed regex on "+$torrent.name >> $errorLog
      if ([bool](get-process  notepad -ea 0)) {
        Stop-Process notepad 
        Invoke-Item $errorLog
      }
    }     
  } 
} 
  
Start-Sleep -Seconds 3
$directories = $unpacked,$untagged
foreach ($dir in $directories) {
  sl -LiteralPath $dir
  $files = ls *.mp4,*.m4v
  # If files are found, continue.
  if ($files) {
    # For each file in unpacked folder do the following.
    foreach ($file in $files) {
      If ($file.IsReadOnly) {
        $file.IsReadOnly = $false
      }      
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
          # Loop through filename array and pull out the series name.
          [string]$seriesName = $null
          [string]$season = $null
          [string]$episode = $null
          [string]$episode2 = $null
          $date = (Get-Date).ToUniversalTime().AddHours(-4).ToString("yyyy-MM-dd")
          Write-Host "Date: $date"
          $titleArray = (Split-Path $newFile -Leaf -Resolve).Split('.-_() ')
          Write-Host "Reading season and episode numbers"
          foreach ($item in $titleArray) {
            $item = $item.Trim()
            # S##E##E##
            if ($item.ToUpper() -Match 'S(\d{1,2})E(\d{1,2})E(\d{1,2})') {
              $details = $item.Split('sSeEeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              $episode2 = $details[3].TrimStart()              
              break
            # S##E##-E##
            } elseif ($item.ToUpper() -Match 'S(\d{1,2})E(\d{1,2})-E(\d{1,2})') {
              $details = $item.Split('sSeE-eE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              $episode2 = $details[3].TrimStart()              
              break
            # S####E##
            } elseif ($item.ToUpper() -Match 'S(\d{1,4})E(\d{1,2})') {
              $details = $item.Split('sSeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()           
              break
            # S#E##
            } elseif ($item.ToUpper() -Match 'S(\d{1,1})E(\d{1,2})') {
              $details = $item.Split('sSeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              break
            # S##E##
            } elseif ($item.ToUpper() -Match 'S(\d{1,2})E(\d{1,2})') {
              $details = $item.Split('sSeE')
              $season = $details[1].TrimStart()
              $episode = $details[2].TrimStart()
              break
            # S##
            } elseif ($item.ToUpper() -Match 'S(\d{1,2})') {
              $details = $item.Split('sS')
              $season = $details[1].TrimStart()
              break
            # ##X##
            } elseif ($item.ToUpper() -Match '(\d{1,2})X(\d{1,2})') {
              $details = $item.Split('xX')
              $season = $details[0].TrimStart()
              $episode = $details[1].TrimStart()
              break
            # 20## (year)
            } elseif ($item -Match '(^(19|20)\d{2}$)') {
              # Add brackets around years in series name.
              $seriesName = $seriesName+" "+"("+$item+")"
            # ###
            } elseif ($item -Match '(^\d{1,3}$)') {
              $season = '0'+$item.Substring(0,1)
              $episode = $item.Substring(1,2)
              break
            # ####
            } elseif ($item -Match '(^\d{1,4}$)') {
              $season = $item.Substring(0,2)
              $episode = $item.Substring(2,2)
              break
            # ########
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
          if (($seriesName -eq $null) -or (($season -eq $null) -and ($episode -eq $null))) {
            break
          }
          # Convert Series Name to Title Case
          $seriesName = (Get-Culture).TextInfo.ToTitleCase($seriesName.ToLower()).Trim().Replace('_',' ')
          Write-Host "Series Name (Title Case): $seriesName"
          # A few specific replace Strings for specific problem shows.
          if ($seriesName -like 'Ballers*') {
            $seriesName = 'Ballers'
          } elseif ($seriesName -like 'Louie*') {
            $seriesName = 'Louie (2010)'
          }
          $seriesName = $seriesName.Replace("S H I E L D","S.H.I.E.L.D.")

          # Tag, rename & move file to correct folder location.
          Write-Host "Parent directory: $untagged"
          if($seriesName) {
            $data = GetSeriesData $SeriesName $season $episode $episode2
            if (!$data) {
              Write-Host "Unable to connect to thetvdb.com, aborting tag."
              $waitFolder = $env:USERPROFILE+"\Downloads\Torrents\Untagged\"
              Write-Host "Moving $file.name to $waitFolder."
              Move-Item -LiteralPath $file $waitFolder -Force -Verbose
              Invoke-Item $errorLog
              break
            } else {
              # Add leading zeros to season and episode numbers.
              $season = $season.PadLeft(2,"0")            
              $episode = $episode.PadLeft(2,"0")
              if ($episode2) {
                $episode2 = $episode2.PadLeft(2,"0")
              }              
              # Set destination folder based on the show name and season.
              Write-Host "Setting destination directory based on series name, season and episode numbers."
              $destination = $destDrive+'TV Shows\'+$seriesName+"\Season "+$season+"\"
              if ($episode2Id) {      
                $fileFormat = $seriesName+" - S"+$season+"E"+$episode+"-E"+$episode2+" - "+$episodeName+" & "+$episode2Name+".mp4"    
              } else {
                $fileFormat = $seriesName+" - S"+$season+"E"+$episode+" - "+$episodeName+".mp4"
              }
              Write-Host "Formatted file name: $fileFormat"
              
              try {
                # Rename the file
                Write-Host "Renaming $file to $fileFormat"
                Rename-Item -Path $newFile $fileFormat -Force -Verbose
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$file.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              # If not renamed, use original name format.
              if (!($fileFormat)) {
                $fileFormat = (Split-Path $newFile -Leaf -Resolve)
                Write-Host "Formatted file name: $fileFormat"
              }
              Write-Host "Deleting existing metadata from $fileFormat."
              $fullFilePath = $dir+$fileFormat
              Write-Host "File path: $fullFilePath"
              try {
                $atomicParams = $fullFilePath,'--metaEnema','--overWrite'
                & $atomicParsley $atomicParams
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$file.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              # If atomic parsley fails due to file errors, re-encode the file using Handbrake.
              if (($LastExitCode -ne 0) -and ($LastExitCode -ne $null)) {
                Write-Host "An error occurred attempting to wipe $fullFilePath of metatags.  Re-encoding file."
                $newFile = 'C:\Users\Steve\Downloads\Torrents\Tagging\'+(Split-Path $fullFilePath -Leaf -Resolve).Replace('.ts','.mp4')
                Write-Host "New file: $newFile"
                if ($newFile -eq $fullFilePath) {
                  $newFile.Replace('.mp4','(2).mp4')
                }
                EncodeWithHB $fullFilePath $newFile
                if (Test-Path $newFile) {
                  Remove-Item $fullFilePath -Force -Verbose
                  mv $newFile $fullFilePath -Force -Verbose
                }
                return
              }
              Start-Sleep -Seconds 5
              # If destination doesn't exist, create it.
              if (!(Test-Path $destination)){
                New-Item -ItemType directory -Path $destination -Verbose
              }
              Write-Host "Destination: $destination"
              # Download the season image file.
              $imgFile = $destination+'folder.jpg'
              if (!(Test-Path $imgFile )) {
                $seriesId = $data.get_Item("SeriesId")
                $image = GetBannerImage $seriesId $dir $season 
                if (Test-Path $image) {
                  mv $image $imgFile
                  # Set folder.jpg to hidden.
                  $file = get-item $imgFile -Force
                  $file.Attributes = 'Hidden'
                }              
              }
              
              # Set up variables for tagging.
              if ($data.get_Item("Episode2Name")) {
                $episodeDesc = [string]$episode+[string]$episode2+' - '+$data.get_Item("EpisodeName")+" & "+$data.get_Item("Episode2Name")
              } else {             
                $episodeDesc = [string]$episode+' - '+$data.get_Item("EpisodeName")
              }  
              $title = $fileFormat.Replace('.mp4','')
              $album = $seriesName+', Season '+[int]$season
              
              # Gather Parameters              
              $atomicParams = @($fullFilePath)
              if ($data.get_Item("Actors")) {
                $actors = $data.get_Item("Actors").Replace('|',', ').Trim(',',' ')
              }
              if (!($actors)) { 
                $actors = 'N/A' 
              }
              $atomicParams += @('--artist',$actors)
              if ($data.get_Item("Genre")) {
                $genre = $data.get_Item("Genre").Replace('|',', ').Trim(',',' ')
              } else { 
                $genre = 'TV Show' 
              }
              $atomicParams += @('--genre',$genre,'--stik','TV Show')
              if ($data.get_Item("AirDate")) {
              $airDate = $data.get_Item("AirDate")+'T12:00:00Z'
              } else {
                $airDate = (Get-Date -Format s)+'Z'
              }
              $atomicParams += @('--year',$airDate)
              if ($episodeDesc) {
                $atomicParams += @('--TVEpisode',$episodeDesc)
              }
              if (!($episodeId)) {
                $episodeId = [int]$episode
              }
              $atomicParams += @('--TVEpisodeNum',$episodeId,'--tracknum',$episodeId)
              if(!($seasonId)) {
                $seasonId = [int]$seasons
              }
              $atomicParams += @('--TVSeason',$seasonId)
              if ($data.get_Item("ContentRating")) {
                $atomicParams += @('--contentRating',$data.get_Item("ContentRating"))
              }
              if ($title) {
                $atomicParams += @('--title',$title)
              }
              if ($data.get_Item("SeriesName")) {
                $atomicParams += @('--albumArtist',$data.get_Item("SeriesName"),'--TVShowName',$data.get_Item("SeriesName"))
              }
              if ($album) {
                $atomicParams += @('--album',$album)
              }
              if ($data.get_Item("Network")) {
                $atomicParams += @('--TVNetwork',$data.get_Item("Network"))
              }
              if ($data.get_Item("Synopsis")) {
                $atomicParams += @('--description',$data.get_Item("Synopsis"),'--longdesc',$data.get_Item("Synopsis"))
              }
              if (Test-Path $imgFile) {
                $atomicParams += @('--artwork',$imgFile)
              }
              $atomicParams += @('--overWrite')
              
              # Tag the file.
              Write-Host "Writing tags to $fileFormat"
              try {
                & $atomicParsley $atomicParams
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$file.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              Start-Sleep -Seconds 5
              
              # Move the file to provided destination.
              try {  
                if (Test-Path $fileFormat) {
                  Move-Item -Path $fileFormat -Destination $destination -Force -Verbose
                }
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$file.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }
              $newMp4Path = $destination+$fileFormat
              
              # Normalize Audio
              try {
                $aacParams = '/k','/r',$newMp4Path
                & $aacgain $aacParams
              } catch {
                "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
                "Exception: "+$_.Exception.Message >> $errorLog
                "Item: "+$file.FullName >> $errorLog
                if ([bool](get-process  notepad -ea 0)) {
                  Stop-Process notepad 
                  Invoke-Item $errorLog
                }
              }          
              $exitLoop = $true
            }
          }
        } catch  {
          "Line: "+$_.InvocationInfo.ScriptLineNumber >> $errorLog
          "Exception: "+$_.Exception.Message >> $errorLog
          "Item: "+$_.Exception.ItemName >> $errorLog
          if ([bool](get-process  notepad -ea 0)) {
            Stop-Process notepad 
            Invoke-Item $errorLog
          }
          
          mv $newFile $waitFolder -Force
          $exitLoop = $true
          return $false      
        }
      } while ($exitLoop -eq $false)
    }
  }
}

# Call a Plex Refresh after adding the files.
refreshPlex
