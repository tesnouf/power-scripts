## Start of a rename script. TE
# get the name of a folder that is NOT 6 digits
# query RPOS for an employee ID that matches (join EmployeeID and Email tables)
# if there is an employee ID for that person and the folder exists move all the contents of the folder to it
# !! Move attachments first, delete the source attachment fie then all other files
# finally delete the folder






$BaseDir = "C:\temp"
$NameToFind = "\*.com"

########################################################################################
## Get a Datatable set up so we can reference it
## requires SSI Integrated security.
$SQLInstance = "Rpostestdb3\lhotse"
$db = "rtpx2"

$connectionString = "Data Source=$SQLInstance; " +
    "Integrated Security=True;" +
    "Initial Catalog=$db"


## SQL query
$sqlQuery = "SELECT
	RIGHT(m.EmailAddress , LEN(m.EmailAddress) - 3) ,
	p.EmployeeID
FROM dbo.EmailProfile m
JOIN dbo.EmployeeProfile p
	ON m.IPCode = p.IPCode
	AND p.StatusCode = 1
	AND p.EmployeeResortCode = 85
WHERE m.StatusCode = 1"

## New verions of SQL connectionString
$connection = new-object system.data.SqlClient.SQLConnection $connectionString
$command = $connection.CreateCommand()
$command.CommandText = $sqlQuery

# Create DataTable to store DataTable# we could use a adataset in the future and wrap all this
# into one email but that doent suit the use case at the moment dt it is.
$ds = new-object System.Data.DataTable

# Create data adapter to fill dataset
$da = new-object System.Data.SqlClient.SqlDataAdapter
$da.SelectCommand = $command
$da.Fill($ds) | Out-Null

$connection.Close()
########################################################################################

Get-ChildItem -Path $BaseDir -Force | Where-Object {$_.name -like "*@*"} | Select-Object  fullname, name | format-table

$foldersToMerge = Get-ChildItem -Path $BaseDir -Force | Where-Object {$_.name -like "*@*"}

foreach($folder in $foldersToMerge){
    # get an object so that we can do a lookup in RPOS for an employee ID
    $proEmail = $folder.Name
    # set up the locations to copy and move to / from
    $sourceLocation = $folder.FullName
    $sourceAttachments = Join-Path $sourceLocation attachments

    Write-Host $proEmail $sourceLocation $sourceAttachments $sqlResult

    # SQL lookup here
    ## Check Dataset to see if senderEmail count is = 1

    $emp = $ds.Select("Column1 like '*$proEmail*'")

    if ($emp.count -eq 1) {
        $sqlResult = $emp.EmployeeID

    } else {
        continue
      }

    Write-Host $proEmail $sourceLocation $sourceAttachments $sqlResult
    # If results = EmployeeID with length 6 numbers then
    # $sqlResult = $proEmail + 234567

    # Look for folder containing those numbers if it does not exist create it.
    $destinationLocation = Join-Path $BaseDir $sqlResult

    If(!(test-path $destinationLocation)){
            Write-Host "Folder Path " $destinationLocation
            New-Item -ItemType Directory -Force -Path $destinationLocation
    }

    $destinationLocationAttachments = Join-Path $BaseDir $sqlResult attachments
        If(!(test-path $destinationLocationAttachments)){
            Write-Host "Folder Path " $destinationLocationAttachments
            New-Item -ItemType Directory -Force -Path $destinationLocationAttachments
    }

    Get-ChildItem -Path $sourceAttachments -Recurse | Move-Item -Destination $destinationLocationAttachments | Out-Null
    Get-ChildItem -Path $sourceLocation -Recurse | Move-Item -Destination $destinationLocation
    # Move all Atachemnts into this folder
    # move all emails into the folder
    # delete the email folder.
    Remove-Item $sourceLocation -Recurse
}
