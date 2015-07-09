################################################################################
#                                                                              #
# This script picks up videos and photos taken on mobile devices synced to the #
# computer via dropbox and moves them to their correct locations               #
#                                                                              #
# Created Date: 27 SEP 2013                                                    #
# Version: 1.3                                                                 #
#                                                                              #
################################################################################

$dropBox = 'D:\Dropbox'
$dropboxFiles = $dropBox+"\Mum Shared"
$errorFile = $dropBox+"\Applications\WindowsPowerShell\Error Logs\"+(Get-Date).ToString("yyyy-MM-dd")+"_error_log.csv"
$photoSync = "D:\Pictures\PhotoSync\"
$PlexSync = "D:\Plex Media Server\Media Upload\"
$iCloud = "D:\Pictures\iCloud Photos\My Photo Stream\"
$videoPath = "D:\Videos\My Home Vids\Phone, Camera, Short Videos\"
$unsortedPath = "D:\Videos\My Home Vids\Phone, Camera, Short Videos\*"
$photoPath = "D:\Pictures\My Photos\Phone Pics\"
$imported = "D:\Pictures\Imported\"
$videoFiles = "*.mov","*.mp4","*.m4v","*.MOV","*.MP4","*.M4V"
$photoFiles = "*.jpg","*.png","*.jpeg","*.JPG","*.PNG","*.JPEG"
$retries = 10
$break = $false
[int]$retry = 0

#sl $cameraUploads
$dropboxVideos = Get-ChildItem $dropboxFiles -Include $videoFiles -Recurse
$dropboxPhotos = Get-ChildItem $dropboxFiles -Include $photoFiles -Recurse
$PlexPhotos = Get-ChildItem $PlexSync -Include $photoFiles -Recurse
$UnsortedVideos = Get-ChildItem $unsortedPath -Include $videoFiles
$importedPhotos = Get-ChildItem $imported -Include $photoFiles -Recurse
$importedVideos = Get-ChildItem $imported -Include $videoFiles -Recurse
$syncPhotos = Get-ChildItem $photoSync -Include $photoFiles -Recurse
$syncVideos = Get-ChildItem $photoSync -Include $videoFiles -Recurse

# If video folder doesn't exist, create it.
if (!(Test-Path $videoPath)) {
	$null = New-Item -Path $videoPath -ItemType Directory
}

# If photo folder doesn't exist, create it.
if (!(Test-Path $photoPath)) {
	$null = New-Item -Path $photoPath -ItemType Directory
}		

$counter = 0
$photos = $PlexPhotos.Count
foreach ($photo in $PlexPhotos) {
  $counter++
  Write-Progress -activity "Reading photos..." -status "$counter of $photos" -percentComplete (($counter / $photos) * 100)
  sl $PlexSync
  $year = [string](($photo.LastWriteTime).Year)
  $month = "{0:D2}" -f (($photo.LastWriteTime).Month)
  $day = "{0:D2}" -f (($photo.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($photo.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($photo.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($photo.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $photoOutput = $photoPath + $year + '\'
  try {
    $newFileName = $dateString+($photo.Extension).ToLower()
    $newFile = $photoOutput+$newFileName
    if (Test-Path $newFile) {
      if ((Get-Item $newFile).Length -eq $photo.Length) {
        Write-Host $newFileName" already exists in $photoOutput.  Deleting."
        rm $photo -Force -Verbose
      } else {
        if (!(Test-Path $photoOutput)) {
          $null = New-Item -Path $photoOutput -ItemType Directory
        }
        # Change the name of the file and test again
        $newFileName = $newFileName.Replace('.jpg','-1.jpg')
        $newFile = $photoOutput+$newFileName
        $newFile = $newFile.Replace('.jpg','-1.jpg')
        if (Test-Path $newFile) {
          if ((Get-Item $newFile).Length -eq $photo.Length) {
            Write-Host $newFileName" already exists in $photoOutput.  Deleting."
            rm $photo -Force -Verbose
          } else {
            $newFileName = $newFileName.Replace('-1.jpg','-2.jpg')
            $newFile = $photoOutput+$newFileName
            if (!(Test-Path $newFile)) {
              Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
              mv $photo $newFile -Force
            } else {
              Write-Host ($dateString+$photo.Extension).ToLower()" already exists in $photoOutput.  Deleting."
              rm $photo -Force -Verbose
            }
          }
        } else {
          Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
          mv $photo $newFile -Force
        }
      }	  
    } else {
      Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
      if (!(Test-Path $photoOutput)) {
        $null = New-Item -Path $photoOutput -ItemType Directory
      }
      mv $photo $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $photo, aborting."
  }
}


$counter = 0
$videos = $UnsortedVideos.Count
foreach ($video in $UnsortedVideos) {
  $counter++
  Write-Progress -activity "Reading videos..." -status "$counter of $videos" -percentComplete (($counter / $videos) * 100)
  sl $videoPath
  $year = [string](($video.LastWriteTime).Year)
  $month = "{0:D2}" -f (($video.LastWriteTime).Month)
  $day = "{0:D2}" -f (($video.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($video.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($video.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($video.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $videoOutput = $videoPath + $year + '\'
  try {
    $newFileName = $dateString+($video.Extension).ToLower()
    $newFile = $videoOutput+$newFileName
    if (Test-Path $newFile) {
      Write-Host ($dateString+$video.Extension).ToLower()" already exists in $videoOutput.  Deleting."
      rm $video
    } else {
      Write-Host "Moving "($dateString+$video.Extension).ToLower()" to $videoOutput"
      if (!(Test-Path $videoOutput)) {
        $null = New-Item -Path $videoOutput -ItemType Directory
      }
      mv $video $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $video, aborting."
  }       
}

$counter = 0
$photos = $importedPhotos.Count
foreach ($photo in $importedPhotos) {
  $counter++
  Write-Progress -activity "Reading photos..." -status "$counter of $photos" -percentComplete (($counter / $photos) * 100)
  sl $PlexSync
  $year = [string](($photo.LastWriteTime).Year)
  $month = "{0:D2}" -f (($photo.LastWriteTime).Month)
  $day = "{0:D2}" -f (($photo.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($photo.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($photo.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($photo.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $photoOutput = $photoPath + $year + '\'
  try {
    $newFileName = $dateString+($photo.Extension).ToLower()
    $newFile = $photoOutput+$newFileName
    if (Test-Path $newFile) {
      Write-Host ($dateString+$photo.Extension).ToLower()" already exists in $photoOutput.  Deleting."
	  rm $photo -Force -Verbose
    } else {
      Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
      if (!(Test-Path $photoOutput)) {
        $null = New-Item -Path $photoOutput -ItemType Directory
      }
      mv $photo $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $photo, aborting."
  }
}

$counter = 0
$videos = $importedVideos.Count
foreach ($video in $importedVideos) {
  $counter++
  Write-Progress -activity "Reading videos..." -status "$counter of $videos" -percentComplete (($counter / $videos) * 100)
  sl $videoPath
  $year = [string](($video.LastWriteTime).Year)
  $month = "{0:D2}" -f (($video.LastWriteTime).Month)
  $day = "{0:D2}" -f (($video.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($video.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($video.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($video.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $videoOutput = $videoPath + $year + '\'
  try {
    $newFileName = $dateString+($video.Extension).ToLower()
    $newFile = $videoOutput+$newFileName
    if (Test-Path $newFile) {
      Write-Host ($dateString+$video.Extension).ToLower()" already exists in $videoOutput.  Deleting."
      rm $video
    } else {
      Write-Host "Moving "($dateString+$video.Extension).ToLower()" to $videoOutput"
      if (!(Test-Path $videoOutput)) {
        $null = New-Item -Path $videoOutput -ItemType Directory
      }
      mv $video $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $video, aborting."
  }       
}

$counter = 0
$photos = $syncPhotos.Count
foreach ($photo in $syncPhotos) {
  $counter++
  Write-Progress -activity "Reading photos..." -status "$counter of $photos" -percentComplete (($counter / $photos) * 100)
  sl $PlexSync
  $year = [string](($photo.LastWriteTime).Year)
  $month = "{0:D2}" -f (($photo.LastWriteTime).Month)
  $day = "{0:D2}" -f (($photo.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($photo.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($photo.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($photo.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $photoOutput = $photoPath + $year + '\'
  try {
    $newFileName = $dateString+($photo.Extension).ToLower()
    $newFile = $photoOutput+$newFileName
    if (Test-Path $newFile) {
      Write-Host ($dateString+$photo.Extension).ToLower()" already exists in $photoOutput.  Deleting."
	  rm $photo -Force -Verbose
    } else {
      Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
      if (!(Test-Path $photoOutput)) {
        $null = New-Item -Path $photoOutput -ItemType Directory
      }
      mv $photo $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $photo, aborting."
  }
}

$counter = 0
$videos = $syncVideos.Count
foreach ($video in $syncVideos) {
  $counter++
  Write-Progress -activity "Reading videos..." -status "$counter of $videos" -percentComplete (($counter / $videos) * 100)
  sl $videoPath
  $year = [string](($video.LastWriteTime).Year)
  $month = "{0:D2}" -f (($video.LastWriteTime).Month)
  $day = "{0:D2}" -f (($video.LastWriteTime).Day)
  $hour = "{0:D2}" -f (($video.LastWriteTime).Hour)
  $minute = "{0:D2}" -f (($video.LastWriteTime).Minute)
  $second = "{0:D2}" -f (($video.LastWriteTime).Second)
  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
  $videoOutput = $videoPath + $year + '\'
  try {
    $newFileName = $dateString+($video.Extension).ToLower()
    $newFile = $videoOutput+$newFileName
    if (Test-Path $newFile) {
      Write-Host ($dateString+$video.Extension).ToLower()" already exists in $videoOutput.  Deleting."
      rm $video
    } else {
      Write-Host "Moving "($dateString+$video.Extension).ToLower()" to $videoOutput"
      if (!(Test-Path $videoOutput)) {
        $null = New-Item -Path $videoOutput -ItemType Directory
      }
      mv $video $newFile -Force
    }
  } catch {
    Write-Host "No date string found for $video, aborting."
  }       
}

#
#$counter = 0
#$photos = $dropboxPhotos.Count
#foreach ($photo in $dropboxPhotos) {
#  $counter++
#  Write-Progress -activity "Reading photos..." -status "$counter of $photos" -percentComplete (($counter / $photos) * 100)
#  sl $PlexSync
#  $year = [string](($photo.LastWriteTime).Year)
#  $month = "{0:D2}" -f (($photo.LastWriteTime).Month)
#  $day = "{0:D2}" -f (($photo.LastWriteTime).Day)
#  $hour = "{0:D2}" -f (($photo.LastWriteTime).Hour)
#  $minute = "{0:D2}" -f (($photo.LastWriteTime).Minute)
#  $second = "{0:D2}" -f (($photo.LastWriteTime).Second)
#  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
#  $photoOutput = $photoPath + $year + '\'
#  try {
#    $newFileName = $dateString+($photo.Extension).ToLower()
#    $newFile = $photoOutput+$newFileName
#    if (Test-Path $newFile) {
#      Write-Host ($dateString+$photo.Extension).ToLower()" already exists in $photoOutput.  Deleting."
#	  rm $photo -Force -Verbose
#    } else {
#      Write-Host "Moving "($dateString+$photo.Extension).ToLower()" to $photoOutput"
#      if (!(Test-Path $photoOutput)) {
#        $null = New-Item -Path $photoOutput -ItemType Directory
#      }
#      mv $photo $newFile -Force
#    }
#  } catch {
#    Write-Host "No date string found for $photo, aborting."
#  }
#}
#
#$counter = 0
#$videos = $dropboxVideos.Count
#foreach ($video in $dropboxVideos) {
#  $counter++
#  Write-Progress -activity "Reading videos..." -status "$counter of $videos" -percentComplete (($counter / $videos) * 100)
#  sl $videoPath
#  $year = [string](($video.LastWriteTime).Year)
#  $month = "{0:D2}" -f (($video.LastWriteTime).Month)
#  $day = "{0:D2}" -f (($video.LastWriteTime).Day)
#  $hour = "{0:D2}" -f (($video.LastWriteTime).Hour)
#  $minute = "{0:D2}" -f (($video.LastWriteTime).Minute)
#  $second = "{0:D2}" -f (($video.LastWriteTime).Second)
#  $dateString = $year+"-"+$month+"-"+$day+" "+$hour+"."+$minute+"."+$second
#  $videoOutput = $videoPath + $year + '\'
#  try {
#    $newFileName = $dateString+($video.Extension).ToLower()
#    $newFile = $videoOutput+$newFileName
#    if (Test-Path $newFile) {
#      Write-Host ($dateString+$video.Extension).ToLower()" already exists in $videoOutput.  Deleting."
#      rm $video
#    } else {
#      Write-Host "Moving "($dateString+$video.Extension).ToLower()" to $videoOutput"
#      if (!(Test-Path $videoOutput)) {
#        $null = New-Item -Path $videoOutput -ItemType Directory
#      }
#      mv $video $newFile -Force
#    }
#  } catch {
#    Write-Host "No date string found for $video, aborting."
#  }       
#}


# If errors print error output to file.
if($error) {
  foreach ($err in $Error) {
    $timestamp = Get-Date -format "yyyy-MM-dd HH:mm:ss:ff"
    (New-Object PsObject -Property @{TimeStamp=$timestamp;Title='Moving iPhone Pics';Script=$err.InvocationInfo.ScriptName;Error=$err.ToString();Pos=$err.InvocationInfo.Line }) | Export-Csv -Path $errorFile -NoTypeInformation -Append
  }
}
