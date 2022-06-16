# Export Graph User Calendar Events

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER ClientID
This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

### PARAMETER TenantID
This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

### PARAMETER CertificateThumbprint
This is an optional parameter. String parameter with the Certificate thumbprint which is uploaded to the AzureAD App.

### PARAMETER Mailboxes
This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.

### PARAMETER StartDate
This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider 1 year backwards from the current date.  

### PARAMETER EndDate
This is an optional parameter. The script will search for meeting items ending based on this EndDate backwards. If this parameter is ommitted, by default will consider 1 year forwards from the current date.

### PARAMETER ExportFolderPath
Insert target folder path named like "C:\Temp". By default this will be "$home\desktop"

### PARAMETER DisableTranscript
This is an optional parameter. Transcript is enabled by default. Use this parameter to not write the powershell Transcript.

### PARAMETER DisconnectMgGraph
This is an optional parameter. Use this parameter to disconnect from MgGraph when it finishes.


## Examples  
### Example 1  
```powershell
PS C:\> .\Export-GraphUserCalendarEvents.ps1 -StartDate 06/20/2022 -Verbose
```  
The script will install required modules if not already installed.  
Later it will request the user credential, and ask for permissions consent if not granted already.  
Then it will search for all meeting items matching the startDate on 06/20/2022 forward.  
It will export the items found to the default ExportFolderPath in files "_alias_-CalendaritemsReport.csv".  

### Example 2  
```powershell
# Following line requires to be connected to Exchange Online
PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
PS C:\> .\Export-GraphUserCalendarEvents.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -Verbose
```
The script will install required modules if not already installed.  
Later it will connect to MgGraph using AzureAD App details (requires ClientID, TenantID and CertificateThumbprint).  
Then it will search for all meeting items matching default StartDate and EndDate, for all mailboxes belonging to the "Staff" Office.  
It will export the items found to the default ExportFolderPath in files "_alias_-CalendaritemsReport.csv".  

## Version History:
### 1.00 - 06/16/2022
 - First Release.
### 1.00 - 06/14/2022
 - Project start.