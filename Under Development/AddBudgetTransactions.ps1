################################################################################
#
# Looks for entries in google spreadsheet, and adds them from bank export if 
# it's not already entered.
#
# Created Date: 04 DEC 2014
# Version: 0.1
#
################################################################################

################################################################################
#
# FUNCTIONS
#
################################################################################

Function Get-BudgetSpreadSheet {
  
  $clientId = $env:GOOGLE_CLIENT_ID
  $email = $env:GOOGLE_DEV_EMAIL
  $pubKeyFingerprint = $env:GOOGLE_PUB_KEY

  Add-Type -Path "C:\Program Files (x86)\Google\Google Data API SDK\Redist\Google.GData.Client.dll"
  Add-Type -Path "C:\Program Files (x86)\Google\Google Data API SDK\Redist\Google.GData.Extensions.dll"
  Add-Type -Path "C:\Program Files (x86)\Google\Google Data API SDK\Redist\Google.GData.Spreadsheets.dll"

  $service = New-Object Google.GData.Spreadsheets.SpreadsheetsService("TestGoogleDocs")
  $service.setUserCredentials($userName, $PWord)
  $query = New-Object Google.GData.Spreadsheets.SpreadsheetQuery
  $feeds = $service.Query($query).Entries
  $title = New-Object Google.GData.Client.AtomTextConstruct
  
}

Get-BudgetSpreadSheet