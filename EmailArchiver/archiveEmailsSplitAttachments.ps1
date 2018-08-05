## Extract all Attachments from the WBSSPay inbox
#initial run took 25 minutes and created 609 folders for 2.3 GB data


## TODO
# save file with first part of email name or other identifier and into a specific folder on the network - done TE
#

Clear Host

# Get Start Time
$startDTM = (Get-Date)

## To Do:
# Create a sqlcmd lookup and save into DataSet/DataTable all WB employee ID's and email accounts.

####################### FUNCTIONS GO HERE ##############################################
#Currently set to replace all illegal characters with an underscore (_)
#Used to create a sensible string for saving email messages
Function Remove-InvalidFileNameChars {

    param(
        [Parameter(Mandatory=$true, Position=0)]
        [String]$Name
    )

    return [RegEx]::Replace($Name, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '_')
}

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




########################################################################################
$testPath = "C:\temp\"
$testPathEX = Join-Path $testPath WBStaff
#$testPath ="N:\P-T\SkiSchoolTraining-WB\SSPay\CertificationArchive"
#$Datestr = [datetime]::Today.ToString('yyyyMMdd') + "_" Currently commented out to handle first run may leave in so that emails are filed by when they were sent.
$INBOXToCheck = "WBSSPay"
$folderToCheck = "_Processed" # _Processed
#$folderToCheck = "_testingProcessed" # _Processed
$folderToMove = "_testingArchive" # _Processed

Write-Host  "The Date is: " $DateStr

Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$olSaveType = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
# This confirms the accounts we are accessing based off Outlooks set up
$namespace.Accounts | Select-Object DisplayName, SmtpAddress, UserName, AccountType, ExchangeConnectionMode | Sort-Object  -Property SmtpAddress | Format-Table
# This confirms the Inboxes set up under that account
$namespace.Folders | Select-Object Name | Format-Table

$oFolder = $namespace.Folders.Item($INBOXToCheck).Folders.Item($folderToCheck).Items
$moveToFolder = $namespace.Folders.Item($INBOXToCheck).Folders.Item($folderToMove)

#$moveToFolder | Select-Object *

foreach ($item in $oFolder) {
    $colAttachments = $item.attachments
    $countAttachments = $colAttachments.Count
    $senderEmail = $item.SenderEmailAddress


    $senderName = $item.SenderName
        #Get email's subject and date
    [string]$subject = $item.Subject
    [string]$sentOn = $item.SentOn
    Write-Host $sentOn
    $Datestr = Remove-InvalidFileNameChars -Name ($item.SentOn.ToString('yyyyMMdd')) ## this is here for First run only
    $Datestr = $Datestr

    ## Check Dataset to see if senderEmail count is = 1
    $emp = $ds.Select("Column1 like '*$senderEmail*'")

    if ($emp.count -eq 1) {
        $senderEmail = $emp.EmployeeID

    }

       # Check to see if a folder with the sender name (SMTP) or Email (EX)  exists
       # TODO:
        #Wrap this section in another IF statement that looks up the EmployeeID associated with an email form LHOTSE
        # If EmpID found then create a folder and save everything to that folder.  If not then use one of the two methods below
        # to temporarily file
        ## Will need some logic to handle moving or renaming email folders to employee ID's before launch
       # Two different ways to save based on the email source
       if($item.SenderEmailType -eq "SMTP") {
            $folderPath = Join-Path $testPath $senderEmail
            Write-Host "Folder Path SMTP " $folderPath
       }
       elseif($item.SenderEmailType -eq "EX"){
            $folderPath = Join-Path $testPathEX $senderName
            Write-Host "Folder Path EX" $folderPath
       }    else {
                continue ## Move to the next Item if it is not an Exchange or SMTP Message
                # expectation is that these will either not exist or can be manually filed.
            }



       # Check to see if the Employees Folder exists and create if not
       If(!(test-path $folderPath)){
            Write-Host "Folder Path " $folderPath
            New-Item -ItemType Directory -Force -Path $folderPath
        }

       $folderAttachments = Join-Path $folderPath attachments

        # Do the same for an atachment subfolder (just in case someone created something manually!
        If(!(test-path $folderAttachments)){
            Write-Host "Folder Path " $folderAttachments
            New-Item -ItemType Directory -Force -Path $folderAttachments
        }

        # add a trailing \ to the destination string
        if ($folderPath[-1] -ne "\") {
            $folderPath += "\"
        }
        if ($folderAttachments[-1] -ne "\") {
            $folderAttachments += "\"
        }

        #Create valid names fof the email to be saved and a filePath
        Write-Host "Folder Path for email: " $folderPath
        #Strip subject and date of illegal characters, add .msg extension, and combine
        $emailName = Remove-InvalidFileNameChars -Name ($Datestr + "_" + $subject + ".msg")
        $emailPath = Join-Path $folderPath $emailName
        Write-Host "Email path and Name: " $emailPath

        # save the email to the network share
        $item.SaveAs($emailPath, $olSaveType::olMSG)

        # save Attachments to the folder
        foreach ($attachment in $colAttachments) {
        Write-Host $attachment.filename
                If ($attachment.fileName -ne $null) {
                    $attachment.saveasfile((Join-Path $folderAttachments $Datestr"_$($attachment.filename)")) ## Save with todays date at the start

                }
        }
        Start-Sleep -Milliseconds 500 # just in case
        #$item.Move($moveToFolder) | out-null
        #Start-Sleep -Milliseconds 10 # just in case
        $item.Unread=$true
        #Start-Sleep -Milliseconds 500 # just in case
#        $item | Get-Member
#        $moveToFolder | Get-Member
        Write-Host "Email Saved going to folder " $moveToFolder
        }





# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed


Write-Host "Completed Task Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
