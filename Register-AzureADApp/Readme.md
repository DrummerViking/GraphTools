# Register AzureAD App

## Authors:  
Agustin Gallegos  

## Info  
Sample taken from https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0&tabs=powershell.  
Added some additional functions to pass a list of scope permissions. Added logic to create a self-signed certificate if not passed a cert file as a parameter, or eventually use a ClientSecret.

## Parameters list  

### PARAMETER AppName
The friendly name of the app registration.

### PARAMETER CertPath
The file path to your .CER public key file. If this parameter is ommitted, and the "UseClientSecret" is not used, we will be creating a new self-signed certificate (with a validity period of 1 year) for the app.

### PARAMETER TenantId
Optional parameter to set the TenantID GUID.

### PARAMETER scopes
Mandatory parameter. This is the list of scope permissions that you want to assign to the app. You can add multiple values comma separated.

### PARAMETER UseClientSecret
Use this optional parameter, to configure a ClientSecret (with a validity period of 1 year) instead of a certificate.

### PARAMETER StayConnected
Use this optional parameter to not disconnect from Graph after the script execution.


## Examples  
### Example 1  
```powershell
PS C:\> .\Register-AzureADApp.ps1 -AppName "Graph DemoApp" -UseClientSecret -StayConnected -scopes "Calendars.ReadWrite","Group.Read.All","mail.Send"
```

The script will create a new AzureAD App Registration.  
The name of the app will be "Graph DemoApp".  
It will add the following API Permissions: "Calendars.ReadWrite","Group.Read.All","mail.Send".  
It will not use a certificate for authentication, it will use a ClientSecret (later will be exposed).  

Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.  
Later it will expose a sample connection method.  

### Example 2  
```powershell
PS C:\> .\Register-AzureADApp.ps1 -AppName "DemoApp" -CertPath "C:\Temp\MyCert.cer" -scopes "User.Read.All","mail.Send"
```

The script will create a new AzureAD App Registration.  
The name of the app will be "DemoApp".  
It will add the following API Permissions: "User.Read.All","mail.Send".  
As the "UseClientSecret" parameter is not used, the app will be using a certificate for authentication.  
As the "CertPath" parameter is used, we will use an existing certificate from "C:\Temp\MyCert.cer".  

Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.  
Later it will expose a sample connection method.  

At the end of the script execution, it will disconnect the current Graph connection.  

### Example 3
```powershell
PS C:\> .\Register-AzureADApp.ps1 -AppName "DemoApp2" -scopes "User.Read.All","mail.Send"
```

The script will create a new AzureAD App Registration.  
The name of the app will be "DemoApp2".  
It will add the following API Permissions: "User.Read.All","mail.Send".  
As the "UseClientSecret" parameter is not used, the app will be using a certificate for authentication.  
As the "CertPath" parameter is not used either, we will create a brand new self-signed certificate.  

Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.  
Later it will expose a sample connection method.  

At the end of the script execution, it will disconnect the current Graph connection.  

## Version History:
### 1.00 - 10/13/2022
 - First Release.
### 1.00 - 10/13/2022
 - Project start.