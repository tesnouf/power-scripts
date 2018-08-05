Two scripts work in Tandem

Archive:
- Navigates to a particular Shared inbox folder through Outlook Exchange Application
- For each email in that Mailbox it
  - Determines the sender
  - looks to the Archive Network folder
    - Determines the Employee ID
    - Creates a folder with the ID or if none exists yet the email address
    - saves all emails into that folder
    - creates an Attachment folder
    - saves all attachments to that folder
    - updates all Emails and Attachments with the DATE sent
    - marks emails as UNREAD  
        - We do this because moving a large number of email via PowerShell has proven buggy with the network lag we have and it's easy to go in and manually move them

  Migrate and Rename
  - There are situations where it is possible that an employee first submits info without either:
    - An email address saved into the host Application
    - A valid Employee ID

  - This will go through the list of folders that have email addresses and IF there is a Employee ID associated with the email address merge the data into that folder.

  TODO's
  Reinvestigate the move portion of files in the Archive script
  Currently set to mare as UNREAD 
