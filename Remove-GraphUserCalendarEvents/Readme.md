# Remove Graph User Calendar Events

## Authors:  
Agustin Gallegos  

## Parameters list  

### PARAMETER ClientID
This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.

### PARAMETER TenantID
This is an optional parameter. String parameter with the TenantID your AzureAD tenant.

### PARAMETER CertificateThumbprint
This is an optional parameter. Certificate thumbprint which is uploaded to the AzureAD App.

### PARAMETER Subject
This is an mandatory parameter. The exact subject text to filter meeting items. This parameter cannot be used together with the "FromAddress" parameter.

### PARAMETER FromAddress
This is an mandatory parameter. The sender address to filter meeting items. This parameter cannot be used together with the "Subject" parameter.

### PARAMETER Mailboxes
This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.

### PARAMETER StartDate
This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider the current date.

### PARAMETER DisableTranscript
This is an optional parameter. Transcript is enabled by default. Use this parameter to not write the powershell Transcript.

### PARAMETER DisconnectMgGraph
This is an optional parameter. Use this parameter to disconnect from MgGraph when it finishes.


## Examples  
### Example 1  
```powershell
PS C:\> .\Remove-GraphUserCalendarEvents.ps1 -Subject "Yearly Team Meeting" -StartDate 06/20/2022 -Verbose
```  
The script will install required modules if not already installed.
Later it will request the user credential, and ask for permissions consent if not granted already.
Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on 06/20/2022 forward.
It will display the items found and proceed to remove them.
All data in the powershell console will be extracted to the Powershell Transcript.

### Example 2  
```powershell
PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} | Select-Object PrimarySMTPAddress
PS C:\> .\Remove-GraphUserCalendarEvents.ps1 -Subject "Yearly Team Meeting" -Mailboxes $mailboxes -Verbose
```
The script will install required modules if not already installed.
Later it will connect to MgGraph using AzureAD App details (requires ClientID, TenantID and CertificateThumbprint).
Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on the current date forward, for all mailboxes belonging to the "Staff" Office.
It will display the items found and proceed to remove them.
All data in the powershell console will be extracted to the Powershell Transcript.

## Version History:
### 1.03 - 06/09/2022  
- Added: Parameter to search based on a StartTime.
- Added: Parameter to disconnect from MgGraph after the script finishes.
- Update: Added logic to authenticate as a user, or as an AzureAD App.
### 1.00 - 05/12/2022
 - First Release.
### 1.00 - 05/12/2022
 - Project start.