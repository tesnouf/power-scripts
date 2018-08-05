#!/usr/bin/env powershell
###############################################################################

### Aim to query SQL Server and insert the results into predefined tables
### want to ask for the pay period date in a message box
##3

###############################################################################




###############################################################################


cls
$dte = Get-Date -UFormat "%Y%m%d"
$usr = $env:UserName
## requires SSI Integrated security.
# $SQLInstance = "Rpostestdb3\lhotse"
# $db = "rtpx2"
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$serverFile = $scriptPath + '\servers.csv'

Write-Host "serverFile $serverFile"

$savedFilePath = $scriptPath + "\" + $dte + "_WBPayRateAudit.xlsx"
Write-Host "Save Location: $savedFilePath"
Write-Host "Date is : $dte"
# Create a Excel workbook - don't  forget to save and close it in the end!
$excel = New-Object -comobject excel.application
$excel.visible = $true
$workbook = $excel.workbooks.add()



$servers = @(Import-CSV $serverFile)

foreach($item in $servers){
          # Write-Host "Server $item.SERVER"
          $sqlServer = $item.Server
          $sqlDB = $item.NAME
          $folder = $item.FOLDER
          Write-Host "server: $sqlServer"
          Write-Host "db: $sqlDB"
          Write-Host "Folder: $FOLDER"

          $connectionString = "Data Source=$sqlServer; " +
              "Integrated Security=True;" +
              "Initial Catalog=$sqlDB"


          $sqlLocation = $scriptPath + '\sql'+ $folder +'\'
          $sentLocation = $scriptPath + '\sent\'
          Write-Host "Beginning Query"
          Write-Host "Running from   $scriptPath"
          Write-Host "SQL Location   $sqlLocation"

          $qFileNames = Get-ChildItem -Path $sqlLocation -Filter *.sql -Name | Sort-Object

            ## Loop through .sql files, insert them into their own listObjects and tabs on a excel sheet
          ForEach($sqlFile in $qFileNames) {
              Write-Host " $sqlFile"

                $sqlFile = $sqlLocation + $sqlFile

                ## New verions of SQL connectionString
                $connection = new-object system.data.SqlClient.SQLConnection $connectionString
                $command = $connection.CreateCommand()
                $command.CommandText = [IO.File]::ReadAllText($sqlFile).replace("datehere",$pp)

                # Create DataSet to store table(s)
                $ds = new-object System.Data.DataSet

                # Create data adapter to fill dataset
                $da = new-object System.Data.SqlClient.SqlDataAdapter
                $da.SelectCommand = $command
                $da.Fill($ds) | Out-Null

                Write-host "Connection sucessufully run on $sqlFile"



                ForEach ($table in $ds.Tables) {  ## this is not ideal need t orework based ona single result form each .sql file
                  ## Add in some logic to check that here are some actual results - don't bother running if resultset is empty
                      $csvResults = $table | ConvertTo-CSV -Delimiter "`t" -NoTypeInformation
                      $csvResults | Set-Clipboard
                      $Worksheet = $workbook.Sheets.Add()
                      # $Worksheet.Name = # work out a tab naming style
                      $Range = $Worksheet.Range("a1")

                    # $Range = $Worksheet.Range("a1","$Columns$($csvResults.count)")
                      $Worksheet.Paste($Range, $false)
                      $sheetName = (Get-Item $sqlFile).Basename  ##name of Query becomes the Tab
                      $tableName = "_" + $sheetName
                      $Excel.DisplayAlerts = $false
                      ## Add in a list object
                      $Worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.xlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes) | Out-Null
                      $Worksheet.ListObjects(1).TableStyle = "TableStyleMedium16"
                      $Worksheet.ListObjects(1).Name = $tableName
                      $Worksheet.Name = "$sheetName"
                      $Worksheet.Columns.AutoFit()
                }


            }



}





$workbook.SaveAs($savedFilePath)


#
#   $qFileNames = Get-ChildItem -Path $sqlLocation -Filter *.sql -Name | Sort-Object
#
#
#
#   ## Loop through .sql files, insert them into their own listObjects and tabs on a excel sheet
#   ForEach($sqlFile in $qFileNames) {
#     Write-Host " $sqlFile"
#
#       $sqlFile = $sqlLocation + $sqlFile
#
#       ## New verions of SQL connectionString
#       $connection = new-object system.data.SqlClient.SQLConnection $connectionString
#       $command = $connection.CreateCommand()
#       $command.CommandText = [IO.File]::ReadAllText($sqlFile).replace("datehere",$pp)
#
#       # Create DataSet to store table(s)
#       $ds = new-object System.Data.DataSet
#
#       # Create data adapter to fill dataset
#       $da = new-object System.Data.SqlClient.SqlDataAdapter
#       $da.SelectCommand = $command
#       $da.Fill($ds) | Out-Null
#
#       Write-host "Connection sucessufully run on $sqlFile"
#
#
#
#       ForEach ($table in $ds.Tables) {  ## this is not ideal need t orework based ona single result form each .sql file
#         ## Add in some logic to check that here are some actual results - don't bother running if resultset is empty
#             $csvResults = $table | ConvertTo-CSV -Delimiter "`t" -NoTypeInformation
#             $csvResults | Set-Clipboard
#             $Worksheet = $workbook.Sheets.Add()
#             # $Worksheet.Name = # work out a tab naming style
#             $Range = $Worksheet.Range("a1")
#
#           # $Range = $Worksheet.Range("a1","$Columns$($csvResults.count)")
#             $Worksheet.Paste($Range, $false)
#             $sheetName = (Get-Item $sqlFile).Basename  ##name of Query becomes the Tab
#             $tableName = "_" + $sheetName
#             $Excel.DisplayAlerts = $false
#             ## Add in a list object
#             $Worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.xlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes) | Out-Null
#             $Worksheet.ListObjects(1).TableStyle = "TableStyleMedium16"
#             $Worksheet.ListObjects(1).Name = $tableName
#             $Worksheet.Name = "$sheetName"
#             $Worksheet.Columns.AutoFit()
#       }
#
#
#   }
#
#   ## Clean up some stuff
#   # Add in generic conneciton details to the default cheet and rename it, do some pretifying
#   ## Set up a bunch of info for the whole workbook to record some Archive Data
#   $version = "V1 - 5/6/2018"
#   $bodyFont = "Verdana"
#   $headingFont = "Arial Rounded MT Bold"
#
#
#   $workbook.Title = "Generic SQL report compliation "  + $dte
#   $workbook.Author = "Tim Esnouf"
#   $ActiveWorksheet = $workbook.worksheets.Item("Sheet1")
#   $ActiveWorksheet.Name = "Query Details"
#   $ActiveWorksheet.cells.item(1,1) = "Server"
#   $ActiveWorksheet.cells.item(1,2) = "DB"
#   $ActiveWorksheet.cells.item(1,3) = "Date Excecuted"
#   $ActiveWorksheet.cells.item(1,4) = "User"
#   $wholeRange = $ActiveWorksheet.UsedRange
#
#   $wholeRange.Font.Name = $headingFont
#   $ActiveWorksheet.cells.item(2, 1) = $SQLInstance
#   $ActiveWorksheet.cells.item(2, 2) = $db
#   $ActiveWorksheet.cells.item(2, 3) = $dte  ## likely this is too large for a cell -> look to insert a text box
#   $ActiveWorksheet.cells.item(2, 4) = $usr
#   $wholeRange = $ActiveWorksheet.UsedRange
#   $wholeRange.EntireColumn.AutoFit() | Out-Null
#
#
