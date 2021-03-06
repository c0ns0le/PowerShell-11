﻿################################################################################
#
# This script allows me to rate my songs against each other, and rank them
# using the Elo rating system.
#
# Created Date: 30 OCT 2013
# Version: 1.1
#
################################################################################

################################################################################
# VARIABLES ETC                                                                #
################################################################################

# TODO: Remove all script-wide variables and pas them between functions as required.

$script:dropBox = 'D:\Dropbox'
$script:powershellDir  = $script:dropBox+"\Applications\WindowsPowerShell\Files\"
$script:ratingCsv = $script:powershellDir+"MusicRatings.csv"
$script:changedCsv = $script:powershellDir+"ChangedRatings.csv"
$script:songCount = $null
$script:songs = $null
$script:rated = $false

################################################################################
# FUNCTIONS                                                                    #
################################################################################

function PlaySong() {
  param ($songPath)
  # Play  song in iTunes for 20 seconds.
  if ($script:iTunes) {
    $script:iTunes.PlayFile($songPath)
    Start-Sleep -Seconds 20
    $script:iTunes.Stop()
  } else {
    $message = "iTunes is not detected on this computer.`n`nMusic playback is temporarily disabled."
    [System.Windows.Forms.MessageBox]::Show($message, 'iTunes Not Located', 'Ok', 'Information')
  }
}

function UpdateiTunes() {
  # TODO: Update songs in iTunes from the CSV data.
  # Narrow further by getting the songs which now have a different iTunes rating as a result.  
  # This process should be limited to when there is significant data to update.
  # After update to iTunes delete ChangedCSV from files directory.
}

function UpdateCSV() {
  # TODO: Search through iTunes, find any missing tracks and add them to CSV.
}

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

function PopulateCsv() {
  $songArray = @()
  $counter = 0
  $tracks = $script:playList.Tracks        
  $trackCount = $tracks.Count
	foreach ($track in $tracks) {
    $counter += 1
    $artist = $track.Artist
    $album = $track.Album
		$name = $track.Name
    $path = $track.Location
    $trackid = $track.TrackID
    $rateCount = 0
    switch($track.Rating) {
      0 {$elo = 1600}
      20 {$elo = 600}
      40 {$elo = 1100}
      60 {$elo = 1600}
      80 {$elo = 2100}
      100 {$elo = 2600}
    }
    $songArray += (New-Object PsObject -Property @{
      Artist=$artist;
      Album=$album;
      ELO=$elo;
      Path=$path;
      RateCount=$rateCount;
      RowID=$counter;
      Track=$name;
      TrackID=$trackid;
    }) | Export-Csv -Path $script:ratingCsv -NoTypeInformation -Append
    Write-Progress -activity "Collecting Music Tracks . . ." -status "Collected: $counter of $trackCount" -percentComplete (($counter / $trackCount) * 100)
	}
  if ($tracks) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tracks)
  }
  if ($track) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($track)
  }
}

function MainWindow() {    
  [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
  [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

  $objForm = New-Object System.Windows.Forms.Form 
  $objForm.Text = "Music Rating System"
  $objForm.Size = New-Object System.Drawing.Size(800,295)
  $objForm.StartPosition = "CenterScreen"
  $objForm.KeyPreview = $True
  $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})
  $objForm.Add_KeyDown({if ($_.KeyCode -eq "1") {$VoteSong1Button.Click()}})
  $objForm.Add_KeyDown({if ($_.KeyCode -eq "2") {$VoteSong2Button.Click()}})

  $objDescriptionLabel = New-Object System.Windows.Forms.Label
  $objDescriptionLabel.Location = New-Object System.Drawing.Size(150,10) 
  $objDescriptionLabel.AutoSize = $true
  $objDescriptionLabel.TextAlign = "TopCenter"
  $objDescriptionLabel.Text = "Devo's iTunes Music Rating Program"
  $objDescriptionLabel.Font = New-Object System.Drawing.Font("Tahoma",18,[System.Drawing.FontStyle]::Bold)
  $objForm.Controls.Add($objDescriptionLabel)

  # Get random song from songs array.
  # TODO: Make this a function somehow which can be called without having to open and close the form.
  $rand1 = Get-Random -Minimum 0 -Maximum $script:songCount
  $rand2 = Get-Random -Minimum 0 -Maximum $script:songCount
  $song1 = $script:songs.Item($rand1)
  $song2 = $script:songs.Item($rand2)
  $song1Name = $song1.Artist+"`r`n"+$song1.Album+"`r`n"+$song1.Track
  $song2Name = $song2.Artist+"`r`n"+$song2.Album+"`r`n"+$song2.Track
  
  $objSong1Text = New-Object System.Windows.Forms.TextBox
  $objSong1Text.Location = New-Object System.Drawing.Size(10,60)
  $objSong1Text.Size = New-Object System.Drawing.Size(375,60)
  $objSong1Text.MultiLine = $true
  $objSong1Text.Enabled = $false
  $objSong1Text.Text = $song1Name
  $objSong1Text.WordWrap = $false
  $objSong1Text.Font = New-Object System.Drawing.Font("Tahoma",10)
  $objForm.Controls.Add($objSong1Text)

  $objSong2Text = New-Object System.Windows.Forms.TextBox
  $objSong2Text.Location = New-Object System.Drawing.Size(395,60)
  $objSong2Text.Size = New-Object System.Drawing.Size(375,60)
  $objSong2Text.MultiLine = $true
  $objSong2Text.Enabled = $false
  $objSong2Text.Text = $song2Name
  $objSong2Text.WordWrap = $false
  $objSong2Text.Font = New-Object System.Drawing.Font("Tahoma",10)
  $objForm.Controls.Add($objSong2Text)

  $VoteSong1Button = New-Object System.Windows.Forms.Button
  $VoteSong1Button.Location = New-Object System.Drawing.Size(10,130)
  $VoteSong1Button.Size = New-Object System.Drawing.Size(375,30)
  $VoteSong1Button.Text = "Vote 1"
  $VoteSong1Button.Font = New-Object System.Drawing.Font("Tahoma",10)
  $VoteSong1Button.Add_Click({$script:continue=$true;CalculateWin $song1 $song2;$objForm.Close()})
  $objForm.Controls.Add($VoteSong1Button)

  $VoteSong2Button = New-Object System.Windows.Forms.Button
  $VoteSong2Button.Location = New-Object System.Drawing.Size(395,130)
  $VoteSong2Button.Size = New-Object System.Drawing.Size(375,30)
  $VoteSong2Button.Text = "Vote 2"
  $VoteSong2Button.Font = New-Object System.Drawing.Font("Tahoma",10)
  $VoteSong2Button.Add_Click({$script:continue=$true;CalculateWin $song2 $song1;$objForm.Close()})
  $objForm.Controls.Add($VoteSong2Button)

  $PlaySong1Button = New-Object System.Windows.Forms.Button
  $PlaySong1Button.Location = New-Object System.Drawing.Size(10,170)
  $PlaySong1Button.Size = New-Object System.Drawing.Size(375,30)
  $PlaySong1Button.Text = "Play Snippet"
  $PlaySong1Button.Font = New-Object System.Drawing.Font("Tahoma",10)
  $PlaySong1Button.Add_Click({PlaySong $song1.Path})
  $objForm.Controls.Add($PlaySong1Button)

  $PlaySong2Button = New-Object System.Windows.Forms.Button
  $PlaySong2Button.Location = New-Object System.Drawing.Size(395,170)
  $PlaySong2Button.Size = New-Object System.Drawing.Size(375,30)
  $PlaySong2Button.Text = "Play Snippet"
  $PlaySong2Button.Font = New-Object System.Drawing.Font("Tahoma",10)
  $PlaySong2Button.Add_Click({PlaySong $song2.Path})
  $objForm.Controls.Add($PlaySong2Button)

  $UpdateiTunesButton = New-Object System.Windows.Forms.Button
  $UpdateiTunesButton.Location = New-Object System.Drawing.Size(10,210)
  $UpdateiTunesButton.Size = New-Object System.Drawing.Size(247,30)
  $UpdateiTunesButton.Text = "Update iTunes"
  $UpdateiTunesButton.Font = New-Object System.Drawing.Font("Tahoma",10)
  $UpdateiTunesButton.Add_Click({UpdateiTunes})
  $objForm.Controls.Add($UpdateiTunesButton)

  $UpdateCSV = New-Object System.Windows.Forms.Button
  $UpdateCSV.Location = New-Object System.Drawing.Size(267,210)
  $UpdateCSV.Size = New-Object System.Drawing.Size(247,30)
  $UpdateCSV.Text = "Import New Songs"
  $UpdateCSV.Font = New-Object System.Drawing.Font("Tahoma",10)
  $UpdateCSV.Add_Click({UpdateCSV})
  $objForm.Controls.Add($UpdateCSV)

  $ExitButton = New-Object System.Windows.Forms.Button
  $ExitButton.Location = New-Object System.Drawing.Size(522,210)
  $ExitButton.Size = New-Object System.Drawing.Size(247,30)
  $ExitButton.Text = "Close"
  $ExitButton.Font = New-Object System.Drawing.Font("Tahoma",10)
  $ExitButton.Add_Click({$script:continue=$false;$objForm.Close()})
  $objForm.Controls.Add($ExitButton)

  # Set form position on screen and activate.
  $objForm.Topmost = $True
  $objForm.Add_Shown({$objForm.Activate()})
  [void] $objForm.ShowDialog()

  # on vote button click, calculate win, then clear both the songs.
  return $script:continue
}

function CalculateWin() {
  param ($winner, $loser)

  $battleKFactor = CalculateKFactor $loser.ELO $winner.ELO

  $winnerExpScore = 1 / (1 + [math]::Pow(10,(($winner.ELO - $loser.ELO) / 400)))
  $loserExpScore = 1 / (1 + [math]::Pow(10,(($loser.ELO - $winner.ELO) / 400)))
  $winnerElo = [int]$winner.ELO + ($battleKFactor * (1 - $winnerExpScore))
  $loserElo = [int]$loser.ELO + ($battleKFactor * (0 - $loserExpScore))
  
  $winnerElo = [math]::Round($winnerElo)
  $loserElo = [math]::Round($loserElo)

  [int]$winner.ELO = $winnerElo
  [int]$loser.ELO = $loserElo
  [int]$winner.RateCount += 1
  [int]$loser.RateCount += 1    

  UpdateSongs $winner $loser
}

# Calculate the K-Factor based on the values provided.
function CalculateKFactor() {
  param ($song1Elo, $song2KElo)

  if (($song1Elo -lt 2100) -or ($song2KElo -lt 2100)) {
    $kFactor = 32
  } elseif (($song1Elo -le 2400) -or ($song2Elo -le 2400)) {
    $kFactor = 24
  } else {
    $kFactor = 16
  }
  return $kFactor
}

# Write the battle results back to the CSV
function UpdateSongs() {
  param ($song1,$song2)
  
  # Update Song 1 in array
  $script:songs.Item($song1.RowID).ELO = $song1.ELO
  $script:songs.Item($song1.RowID).RateCount = $song1.RateCount
  if ($song1.ELO -gt 2600) {
    $script:songs.Item($song1.RowID).Rating = 100
  } elseif ($song1.ELO -gt 2100) {
    $script:songs.Item($song1.RowID).Rating = 80
  } elseif ($song1.ELO -gt 1600) {
    $script:songs.Item($song1.RowID).Rating = 60
  } elseif ($song1.ELO -gt 1100) {
    $script:songs.Item($song1.RowID).Rating = 40
  } else {
    $script:songs.Item($song1.RowID).Rating = 20
  }

  #Update Song 2 in array
  $script:songs.Item($song2.RowID).ELO = $song2.ELO
  $script:songs.Item($song2.RowID).RateCount = $song2.RateCount
  if ($song2.ELO -gt 2600) {
    $script:songs.Item($song2.RowID).Rating = 100
  } elseif ($song2.ELO -gt 2100) {
    $script:songs.Item($song2.RowID).Rating = 80
  } elseif ($song2.ELO -gt 1600) {
    $script:songs.Item($song2.RowID).Rating = 60
  } elseif ($song2.ELO -gt 1100) {
    $script:songs.Item($song2.RowID).Rating = 40
  } else {
    $script:songs.Item($song2.RowID).Rating = 20
  }
  
  # Check if it exists in the 'changed' array.  If it does, delete and replace it.
  if ($script:changed.TrackID -contains $song1.TrackID) {
    $script:changed = @($script:changed |? {$_.TrackID -ne $song1.TrackID})
  }
  $script:changed += $song1
  if ($script:changed.TrackID -contains $song2.TrackID) {
    $script:changed = @($script:changed |? {$_.TrackID -ne $song2.TrackID})
  }
  $script:changed += $song2

  if ($script:rated -eq $false) {
    $script:rated = $true
  }
}

################################################################################
# PROCESS FLOW                                                                 #
################################################################################

# Connect to iTunes COM object.
ConnectiTunes

if ((!(Test-Path $script:ratingCsv)) -and ($script:iTunes)) {
  PopulateCsv
} 
# Load CSV for processing.
try {
  $script:songs = Import-Csv $script:ratingCsv
} catch {
  # Handle error on load.
  $message = "An error occured attempting to load $script:ratingCsv."
  [System.Windows.Forms.MessageBox]::Show($message, 'CSV Load Error', 'Ok', 'Error')
}

$script:songCount = (($script:songs).Count)-1
if (Test-Path $script:changedCsv) {
  try {
    $script:changed = Import-Csv $script:changedCsv
  } catch {
    # Handle error on load.
    $message = "An error occured attempting to load $script:changedCsv."
    [System.Windows.Forms.MessageBox]::Show($message, 'CSV Load Error', 'Ok', 'Error')
  }
} else {
  $script:changed = @()
}


do {
  $script:continue = MainWindow
} while ($script:continue -eq $true)

if ($script:rated) {
  try {
    $script:songs | Export-Csv  $script:ratingCsv -NoTypeInformation
    $script:changed | Export-Csv  $script:changedCsv -NoTypeInformation
  } catch {
    # Handle error on export.
    $message = "Failed to write results back to the CSV files.`n`nVotes from this session have been lost."
    [System.Windows.Forms.MessageBox]::Show($message, 'CSV Write Error', 'Ok', 'Error')
  }
}

# Close the iTunes COM objects.
CloseiTunes


