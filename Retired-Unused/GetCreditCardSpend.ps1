#################################################################################
#
# Prerequisites:`t                                                               
# Go to https://ibanking.stgeorge.com.au/ibank/loginPage.action and login       
#   then go to:`thttps://ibanking.stgeorge.com.au/ibank/accountDetails.action?index=3#transaction-all
#   then:`thttps://ibanking.stgeorge.com.au/ibank/accountDetails.action?index=3#transHistExport
#
# This will export a copy of all of the historic transactions.
#
#################################################################################

clear


$downloads = $env:USERPROFILE+"\Downloads"
sl $downloads
$csvFile = ls trans*.csv
$csv = @(Import-Csv $csvFile)

# Groceries
$shops = '*woolworths*',
  '*iga*',
  '*costco*',
  '*coles*',  
  '*cbd meats*',  
  '*supabarn*',
  '*aldi*',
  '*shop-rite*',
  '*chemist*',
  '*target*',
  '*news*'

# Steve's stuff
$steve = '*mocan green grout*',
  '*acton*it*holdings*',
  '*action*it*holdings*',
  '*spar*',
  '*dan*murphy*',
  '*liquor*',
  '*jaycar*',
  '*sportingbet*',
  '*mocca*',
  '*pj*',
  '*bottler*',
  '*uni*pub*',
  '*newacton*',
  '*sony*ent*'
  
# Take Away/Food
$takeaway = '*turkish*',
  '*subway*',
  '*snag*',
  '*chicken*spot*',
  '*delissio*',
  '*5*Senses*',
  '*cafe*',
  '*bistro*',
  'mcdonalds*',
  '*sumo*salad*',
  '*little*istanbul*',
  '*lemon pepper*',
  '*caffe*',
  '*walt & burley*',
  '*coffee*',
  '*guzman*y*gomez*',
  '*common*grounds*'
  

# Hardware
$hardware = '*masters*',
  '*bunnings*',
  '*ern*smith*'

# Online
$online = '*paypal*',
  '*amazon*',
  '*cydia*',
  '*itunes*',
  '*allbids*',
  '*msy*',
  '*ebay*'

# Bike & Car
$vehicles = '*act road user*',
  '*motorcycle*',
  '*cardirect*',
  '*car*care*',
  '*auto*',
  '*caltex*',
  '*petrol*',
  '*parking*'

# Holidays
$holidays = '*perisher*',
  '*hotel*',
  '*guthega*',
  '*seddon*deadly*sins*',
  '*jindabyne*'

# Taxis
$taxis = '*aerial*',
  '*taxi*',
  '*cabs*'

[datetime]$startDate = Get-Date
[datetime]$endDate = Get-Date

foreach ($record in $csv) {
  [decimal]$debitTotal += $record.Debit
  [decimal]$creditTotal += $record.Credit
  
    
  foreach ($shop in $shops) {    
    if (($record.Description).ToLower() -like $shop) {
      Clear-Variable [datetime]$tempDate
      if (!($record.Credit)) {
        [datetime]$tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
        if ($tempDate -lt $startDate) {
          $startDate = $tempDate          
        }
        $filteredShops += @($record)
        $added += @($record)        
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($sf in $steve) {    
      if (($record.Description).ToLower() -like $sf) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          [datetime]$tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredSteve += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($ta in $takeaway) {    
      if (($record.Description).ToLower() -like $ta) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          [datetime]$tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredTakeaway += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    if (($record.Description).ToLower() -like '*jb hi fi*') {
      Clear-Variable [datetime]$tempDate
      if (!($record.Credit)) {
        [datetime]$tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
        if ($tempDate -lt $startDate) {
          $startDate = $tempDate
        }
        $filteredJb += @($record)
        $added += @($record)
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($hwshop in $hardware) {    
      if (($record.Description).ToLower() -like $hwshop) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          [datetime]$tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredHardware += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($olshop in $online) {    
      if (($record.Description).ToLower() -like $olshop) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredOnline += @($record)
          $added += @($record)
        }
      }
    }  
  }
  if ($added -notcontains $record) {
    foreach ($vehicle in $vehicles) {    
      if (($record.Description).ToLower() -like $vehicle) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredVehicles += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($days in $holidays) {    
      if (($record.Description).ToLower() -like $days) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredHolidays += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    foreach ($taxi in $taxis) {    
      if (($record.Description).ToLower() -like $taxi) {
        Clear-Variable [datetime]$tempDate
        if (!($record.Credit)) {
          $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
          if ($tempDate -lt $startDate) {
            $startDate = $tempDate
          }
          $filteredTaxis += @($record)
          $added += @($record)
        }
      }
    }
  }
  if ($added -notcontains $record) {
    if (($record.Description).ToLower() -like '*interest*') {
      Clear-Variable [datetime]$tempDate
      if (!($record.Credit)) {
        $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
        if ($tempDate -lt $startDate) {
          $startDate = $tempDate
        }
        $filteredInterest += @($record)
        $added += @($record)
      }
    }
  }
  if ($added -notcontains $record) {
    Clear-Variable [datetime]$tempDate
    if (!($record.Credit)) {
      $tempDate = Get-Date ([datetime]::ParseExact($record.Date, "dd/MM/yyyy", $null))
      if ($tempDate -lt $startDate) {
        $startDate = $tempDate
      }
      if ($filteredMisc -notcontains $record) {
        $filteredMisc += @($record)
      }
    }
  }
}

$weeks = (($endDate - $startDate).Days)/7
$totalSpendShops = ($filteredShops | Measure-Object 'Debit' -Sum).Sum
$totalSpendSteve = ($filteredSteve | Measure-Object 'Debit' -Sum).Sum
$totalSpendJb = ($filteredJb | Measure-Object 'Debit' -Sum).Sum
$totalSpendHardware = ($filteredHardware | Measure-Object 'Debit' -Sum).Sum
$totalSpendOnline = ($filteredOnline | Measure-Object 'Debit' -Sum).Sum
$totalSpendVehicles = ($filteredVehicles | Measure-Object 'Debit' -Sum).Sum
$totalSpendHolidays = ($filteredHolidays | Measure-Object 'Debit' -Sum).Sum
$totalSpendTaxis = ($filteredTaxis | Measure-Object 'Debit' -Sum).Sum
$totalSpendMisc = ($filteredMisc | Measure-Object 'Debit' -Sum).Sum
$totalSpendTakeaway = ($filteredTakeaway | Measure-Object 'Debit' -Sum).Sum
$totalSpendInterest = ($filteredInterest | Measure-Object 'Debit' -Sum).Sum

$fornightAveShops = ($totalSpendShops/$weeks)*2
$fornightAveSteve = ($totalSpendSteve/$weeks)*2
$fornightAveJb = ($totalSpendJb/$weeks)*2
$fornightAveHardware = ($totalSpendHardware/$weeks)*2
$fornightAveOnline = ($totalSpendOnline/$weeks)*2
$fornightAveVehicles = ($totalSpendVehicles/$weeks)*2
$fornightAveHolidays = ($totalSpendHolidays/$weeks)*2
$fornightAveTaxis = ($totalSpendTaxis/$weeks)*2
$fornightAveMisc = ($totalSpendMisc/$weeks)*2
$fornightAveTakeaway = ($totalSpendTakeaway/$weeks)*2
$fornightAveInterest = ($totalSpendInterest/$weeks)*2
$debitAve = ($debitTotal/$weeks)*2
$creditAve = ($creditTotal/$weeks)*2
$totalAve = (($debitTotal-$creditTotal)/$weeks)*2

$filteredShops
$filteredSteve
$filteredJb
$filteredHardware
$filteredOnline
$filteredVehicles
$filteredHolidays
$filteredTaxis
$filteredTakeaway
$filteredInterest
$filteredMisc

Write-Host `
"`n`nFornightly Averages"`
"`n_______`n`nGroceries:`t"$fornightAveShops.ToString("c")`
"`nTake Away:`t"$fornightAveTakeaway.ToString("c")`
"`nJB Hi-Fi:`t"$fornightAveJb.ToString("c")`
"`nHardware:`t"$fornightAveHardware.ToString("c")`
"`nOnline:`t"$fornightAveOnline.ToString("c")`
"`nVehicles:`t"$fornightAveVehicles.ToString("c")`
"`nHolidays:`t"$fornightAveHolidays.ToString("c")`
"`nTaxis:`t`t"$fornightAveTaxis.ToString("c")`
"`nSteve:`t`t"$fornightAveSteve.ToString("c")`
"`nMisc:`t`t"$fornightAveMisc.ToString("c")`
"`nInterest:`t"$fornightAveInterest.ToString("c")`
"`n_______`n`nTOTAL CREDIT:`t"$creditTotal.ToString("c")`
"`nTOTAL DEBIT:`t"$debitTotal.ToString("c")`
"`nDIFFERENCE:`t"($debitTotal-$creditTotal).ToString("c")