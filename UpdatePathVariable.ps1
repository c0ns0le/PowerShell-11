# Import path variable.
$path = ($env:Path | Sort-Object).split(';')

# Get all local drives.
$drives = @((get-wmiobject Win32_LogicalDisk | ? {$_.drivetype -eq 3} | % {get-psdrive $_.deviceid[0]}).Root)

# Set up initial search locations.
$locations = @('C:\Users','C:\Program Files','C:\Program Files (x86)','C:\My Documents')

# Add in additional local drives.
$locations += $drives | Where-Object {
  $_.ToString() -ne 'C:\'
}

# Iterate through locations and collect executables.
foreach ($location in $locations) {
  if (Test-Path $location) {
    sl $location
    $executables += gci *.exe -Recurse -ErrorAction SilentlyContinue
  }
}

# Itertate through executables list, ignoring some location paths.
foreach ($exe in $executables) {
  $parent = Split-Path -parent $exe
  if (($parent -notlike '*Adobe*') `
    -and ($parent -notlike '*AppData*') `
    -and ($parent -notlike '*backup*') `
    -and ($parent -notlike '*Common Files*') `
    -and ($parent -notlike '*Debug*') `
    -and ($parent -notlike '*Drivers*') `
    -and ($parent -notlike '*Install*') `
    -and ($parent -notlike '*Java*') `
    -and ($parent -notlike '*Microsoft*') `
    -and ($parent -notlike '*ProgramData*') `
    -and ($parent -notlike '*Setup*') `
    -and ($parent -notlike '*Update*') `
    -and ($parent -notlike '*Windows*') `
    -and ($parent -notmatch("(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"))){
    if ($path -notcontains $parent) {
      $path += $parent
    }
  }
}

# Set up new path variable string.
$newPath = ($path | Sort-Object) -join ';'

# Save path variable.
$env:Path = $newPath