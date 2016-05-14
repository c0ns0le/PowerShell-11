function ConnectiTunes() {
  param ($script:pListName = "Music")
  $script:iTunes = New-Object -ComObject iTunes.application -ErrorAction SilentlyContinue
  if ($script:iTunes) {
    $script:playList = $script:iTunes.LibrarySource.Playlists.ItemByName($script:pListName)
  }
}

function CloseiTunes() {
  # Release COM objects.
  if ($script:iTunes) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:iTunes)
  }
  if ($script:playList) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:playList)
  }
}

function getiTunesMusic() {
  $songArray = @()
  $tracks = $script:playList.Tracks
  foreach ($track in $tracks) {
    if ($track.Location -like "*.m4p") {
      $albumArtist = $null
      if ($track.AlbumArtist -eq $null) {
        $albumArtist = $track.Artist
      } else {
        $albumArtist = $track.AlbumArtist
      }
      $songArray += (New-Object PsObject -Property @{
        Track=$track.Name;
        Album=$track.Album;     
        AlbumArtist=$albumArtist;
        Artist=$track.Artist;
        TrackNumber=$track.TrackNumber;
	      TrackCount=$track.TrackCount;
        Year=$track.Year;
	      DiscNumber=$track.DiscNumber;
	      DiscCount=$track.DiscCount;		
        Path=$track.Location;
      })
    }
  }
  return $songArray
}

$appleMusic = 'D:\iTunes Media\Apple Music'
$output = 'D:\Converted\'
$metaflac = 'D:\Dropbox\Applications\Open Source\Utilities\flac-1.3.1-win\win64\metaflac.exe'
$tagFile = $output+'m4p_tags.csv'

$flacFiles = gci -Path $appleMusic -Include '*.flac' -Recurse

if (!(Test-Path $output)) {
  mkdir $output
}

ConnectiTunes
$tags = @(getiTunesMusic)
CloseiTunes

foreach ($flac in $flacFiles) {
  $tagYear = $null
  $tagArtist = $null
  $tagAlbumArtist = $null
  $tagAlbum = $null
  $tagTrackNum = $null
  $tagTrackCount = $null
  $tagTitle = $null
  $tagDiscNum = $null
  $tagDiscCount = $null
  
  $m4pPath = ($flac.Fullname).Replace('flac','m4p')
  foreach($key in ($tags.GetEnumerator() | Where-Object {$_.Path -eq $m4pPath})) {
    $tagTitle = $key.Track
    $tagAlbum = $key.Album
    $tagArtist = $key.AlbumArtist
    $tagAlbumArtist = $key.AlbumArtist
    $tagTrackNum = $key.TrackNumber
    $tagTrackCount = $key.TrackCount
    $tagYear = $key.Year
    $tagDiscNum = $key.DiscNumber
    $tagDiscCount = $key.DiscCount	
  }
  if ($tagTitle) {
    $year = '--set-tag="DATE='+$tagYear+'"'
    $artist = '--set-tag="ARTIST='+$tagArtist.Replace('"',"'")+'"'
    $albumArtist = '--set-tag="ALBUM ARTIST='+$tagAlbumArtist.Replace('"',"'")+'"'
    $album = '--set-tag="ALBUM='+$tagAlbum.Replace('"','')+'"'
    $trackNum = '--set-tag="TRACKNUMBER='+$tagTrackNum+'"'
    $trackCount = '--set-tag="TRACKCOUNT='+$tagTrackCount+'"'
    $title = '--set-tag="TITLE='+$tagTitle.Replace('"','')+'"'
    $discNum = '--set-tag="DISCNUMBER='+$tagDiscNum+'"'
    $discCount = '--set-tag="DISCCOUNT='+$tagDiscCount+'"'
    
    #Tag File
    $exec = '& "'+$metaflac+'" --remove-all-tags '+$artist+' '+$album+' '+$albumArtist+' '+$trackNum+' '+$trackCount+' '+$title+' '+$year+' '+$discNum+' '+$discCount+' "'+$flac+'"'
    Invoke-Expression $exec
       
    # Move file to destination
    mv -LiteralPath $flac $output -Force
    if ($?) {
      # Delete the original m4p
      $shell = new-object -comobject "Shell.Application"
      $item = $shell.Namespace(0).ParseName("$m4pPath")
      $item.InvokeVerb("delete")
      
      # Delete empty folders in Apple Music directory.
      sl $appleMusic
      Get-ChildItem -recurse | Where {
        $_.PSIsContainer -and `
        @(Get-ChildItem -Lit $_.Fullname -r | 
          Where {!$_.PSIsContainer}).Length -eq 0
      } | Remove-Item -recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Invoke-Item $output
