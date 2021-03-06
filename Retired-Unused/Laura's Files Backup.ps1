param (
  $drive = 'G:\'
)
sl $drive
$exts = @('*.3dm','*.3ds','*.3g2','*.3gp','*.7z','*.accdb','*.ai','*.aif','*.apk','*.app','*.asf','*.asp','*.aspx','*.asx','*.avi','*.bak','*.bat','*.bin','*.bmp','*.c','*.cab','*.cbr','*.cer','*.cfg','*.cfm','*.cgi','*.class','*.com','*.cpl','*.cpp','*.crdownload','*.crx','*.cs','*.csr','*.css','*.csv','*.cue','*.cur','*.dat','*.db','*.dbf','*.dds','*.deb','*.dem','*.deskthemepack','*.dmg','*.dmp','*.doc','*.docx','*.drv','*.dtd','*.dwg','*.dxf','*.eps','*.exe','*.fla','*.flv','*.fnt','*.fon','*.gadget','*.gam','*.gbr','*.ged','*.gif','*.gpx','*.gz','*.h','*.hqx','*.htm','*.html','*.icns','*.ico','*.ics','*.iff','*.indd','*.ini','*.iso','*.jar','*.java','*.jpg','*.js','*.jsp','*.key','*.keychain','*.kml','*.kmz','*.lnk','*.log','*.lua','*.m','*.m3u','*.m4a','*.m4v','*.max','*.mdb','*.mdf','*.mid','*.mim','*.mov','*.mp3','*.mp4','*.mpa','*.mpg','*.msg','*.msi','*.nes','*.obj','*.odt','*.otf','*.pages','*.part','*.pct','*.pdb','*.pdf','*.php','*.pif','*.pkg','*.pl','*.plugin','*.png','*.pps','*.ppt','*.pptx','*.prf','*.ps','*.psd','*.pspimage','*.py','*.ra','*.rar','*.rm','*.rom','*.rpm','*.rss','*.rtf','*.sav','*.sdf','*.sh','*.sitx','*.sln','*.sql','*.srt','*.svg','*.swf','*.swift','*.sys','*.tar','*.tar.gz','*.tax2012','*.tax2014','*.tex','*.tga','*.thm','*.tif','*.tiff','*.tmp','*.toast','*.torrent','*.ttf','*.txt','*.uue','*.vb','*.vcd','*.vcf','*.vcxproj','*.vob','*.wav','*.wma','*.wmv','*.wpd','*.wps','*.wsf','*.xcodeproj','*.xhtml','*.xlr','*.xls','*.xlsx','*.xml','*.yuv','*.zip','*.zipx')
$files = gci -Include $exts -Recurse -ErrorAction SilentlyContinue | Where-Object { !($_.FullName).StartsWith($drive+"Windows") -and !($_.FullName).StartsWith($drive+"Program Files") } 
$counter = $files.Count
for($i = 0; $i -lt $counter; $i++) {
  Write-Progress -activity "Backing up files . . ." -status "Processing: $i of $counter" -percentComplete (($i / $counter) * 100)
  $file = $files[$i]
  $original = $file.Fullname
  $newPath = ($file.DirectoryName).Replace("G:\Users\Laura","D:\Laura's Backup\").Replace("G:\","D:\Laura's Backup\")+"\"
  $new = $newPath+$file.Name
  if (!(Test-Path $new)) {
    if(!(Test-Path $newPath)) {
      mkdir $newPath
    }
    Write-Host $original" -> `n`t"$new
    cp $original $new 
  }  
}