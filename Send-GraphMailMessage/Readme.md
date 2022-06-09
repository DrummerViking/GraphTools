# Send Graph Mail Message

## Authors:  
Agustin Gallegos  
Vlad Radu  

## Parameters list  

### PARAMETER ToRecipients
List of recipients in the "To" list. This is a Mandatory parameter.

### PARAMETER CCRecipients
List of recipients in the "CC" list. This is an optional parameter.

### PARAMETER BccRecipients
List of recipients in the "Bcc" list. This is an optional parameter.

### PARAMETER Subject
Use this parameter to set the subject's text. By default will have: "Test message sent via Graph".

### PARAMETER Body
Use this parameter to set the body's text. By default will have: "Test message sent via Graph using Powershell".

### PARAMETER DisconnectMgGraph
Use this optional parameter to disconnect from MgGraph after sending the email message.


## Examples  
### Example 1  
```powershell
PS C:\> .\Send-GraphMailMessage.ps1 -ToRecipients "john@contoso.com"
```  
The script will download (if not already installed) the required Graph powershell modules.
Will authenticate to MS Graph (if not already), and will prompt for consent if it was not already granted.
Then will send the email message to "john@contoso.com" from the user previously authenticated.

### Example 2  
```powershell
PS C:\> .\Send-GraphMailMessage.ps1 -ToRecipients "julia@contoso.com","carlos@contoso.com" -BccRecipients "mark@contoso.com" -Subject "Lets meet!"
```
The script will download (if not already installed) the required Graph powershell modules.
Will authenticate to MS Graph (if not already), and will prompt for consent if it was not already granted.
Then will send the email message to "julia@contoso.com" and "carlos@contoso.com" and bcc to "mark@contoso.com", from the user previously authenticated.

## Version History:
### 1.01 - 05/06/2022  
- Added: Added 'Bcc' capability as well.
### 1.00 - 05/05/2022
 - First Release.
### 1.00 - 05/05/2022
 - Project start.