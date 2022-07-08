# Update Graph User Calendar Events End Date

## Authors:  
Agustin Gallegos  

## Required App permissions  
In case you are using Application permissions (not delegated), your Azure App needs to be granted "Calendars.ReadWrite" permission.  

## Parameters list  

### PARAMETER ClientID
This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

### PARAMETER TenantID
This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

### PARAMETER CertificateThumbprint
This is an optional parameter. Certificate thumbprint which is uploaded to the AzureAD App.

### PARAMETER Subject
This is a mandatory parameter. The exact subject text to filter meeting items. This parameter cannot be used together with the "FromAddress" parameter.

### PARAMETER Mailboxes
This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.

### PARAMETER EventEndDate
This is a required parameter. This is the end date we will update on the Recurrent meeting items. By Default will be 1 year forward from the current date.  

### PARAMETER StartDate
This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider the current date.  

### PARAMETER EndDate
This is an optional parameter. The script will search for meeting items ending based on this EndDate backwards. If this parameter is ommitted, by default will consider 1 year forward from the current date.  

### PARAMETER DisableTranscript
This is an optional parameter. Transcript is enabled by default. Use this parameter to not write the powershell Transcript.

### PARAMETER ListOnly
This is an optional parameter. Use this parameter to list the calendar events found, without deleting them. This is a good parameter to use, to actually see the current found events and double check these are the ones to be deleted.  

### PARAMETER DisconnectMgGraph
This is an optional parameter. Use this parameter to disconnect from MgGraph when it finishes.


## Examples  
### Example 1  
```powershell
PS C:\> .\Update-GraphUserCalendarEventsEndDate.ps1 -Subject "Yearly Team Meeting" -StartDate 06/20/2022 -Verbose
```  
The script will install required modules if not already installed.  
Later it will request the user credential, and ask for permissions consent if not granted already.  
Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on 06/20/2022 forward.  
It will set the end date on the recurrent meetings on 1 year forward from the current date.  

### Example 2  
```powershell
# Following line requires to be connected to Exchange Online
PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
PS C:\> .\Update-GraphUserCalendarEventsEndDate.ps1 -Subject "Yearly Team Meeting" -Mailboxes $mailboxes.PrimarySMTPAddress  -ClientID "12345678" -TenantId "abcdefg" -CertificateThumbprint "a1b2c3d4" -Verbose
```
The script will install required modules if not already installed.  
Later it will connect to MgGraph using AzureAD App details (requires ClientID, TenantID and CertificateThumbprint).  
Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on the current date forward, for all mailboxes belonging to the "Staff" Office.  
It will set the end date on the recurrent meetings on 1 year forward from the current date.  

## Version History:
### 1.00 - 07/08/2022
 - First Release.
### 1.00 - 07/08/2022
 - Project start.