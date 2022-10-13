<#
    .SYNOPSIS
    Script to create an Azure App Registration.

    .DESCRIPTION
    Script to create an Azure App Registration.
    It will require an additional PS module "Microsoft.Graph.Applications", if not already installed it will download it.
    You have to pass the list of app permissions you want to grant.
    You can use the "UseClientSecret" switch parameter to configure a new ClientSecret for the app. If this parameter is ommitted, we will use a Certificate.
    You can pass a certificate path if you have an existing certificate, or leave the parameter blank and a new self-signed certificate will be created.

    .PARAMETER AppName
    The friendly name of the app registration.

    .PARAMETER CertPath
    The file path to your .CER public key file. If this parameter is ommitted, and the "UseClientSecret" is not used, we will be creating a new self-signed certificate (with a validity period of 1 year) for the app.

    .PARAMETER TenantId
    Optional parameter to set the TenantID GUID.

    .PARAMETER scopes
    Mandatory parameter. This is the list of scope permissions that you want to assign to the app. You can add multiple values comma separated.

    .PARAMETER UseClientSecret
    Use this optional parameter, to configure a ClientSecret (with a validity period of 1 year) instead of a certificate.

    .PARAMETER StayConnected
    Use this optional parameter to not disconnect from Graph after the script execution.

    .EXAMPLE
    PS C:\> .\Register-AzureADApp.ps1 -AppName "Graph DemoApp" -UseClientSecret -StayConnected -scopes "Calendars.ReadWrite","Group.Read.All","mail.Send"

    The script will create a new AzureAD App Registration.
    The name of the app will be "Graph DemoApp".
    It will add the following API Permissions: "Calendars.ReadWrite","Group.Read.All","mail.Send".
    It will not use a certificate for authentication, it will use a ClientSecret (later will be exposed).

    Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.
    Later it will expose a sample connection method.

    .EXAMPLE
    PS C:\> .\Register-AzureADApp.ps1 -AppName "DemoApp" -CertPath "C:\Temp\MyCert.cer" -scopes "User.Read.All","mail.Send"

    The script will create a new AzureAD App Registration.
    The name of the app will be "DemoApp".
    It will add the following API Permissions: "User.Read.All","mail.Send".
    As the "UseClientSecret" parameter is not used, the app will be using a certificate for authentication.
    As the "CertPath" parameter is used, we will use an existing certificate from "C:\Temp\MyCert.cer".

    Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.
    Later it will expose a sample connection method.

    At the end of the script execution, it will disconnect the current Graph connection.

    .EXAMPLE
    PS C:\> .\Register-AzureADApp.ps1 -AppName "DemoApp2" -scopes "User.Read.All","mail.Send"

    The script will create a new AzureAD App Registration.
    The name of the app will be "DemoApp2".
    It will add the following API Permissions: "User.Read.All","mail.Send".
    As the "UseClientSecret" parameter is not used, the app will be using a certificate for authentication.
    As the "CertPath" parameter is not used either, we will create a brand new self-signed certificate.

    Once the app is created, the script will expose the link to grant "Admin consent" for the permissions requested.
    Later it will expose a sample connection method.

    At the end of the script execution, it will disconnect the current Graph connection.

    .NOTES
    General notes
#>
[Cmdletbinding()]
param(
    [Parameter(Mandatory = $true)]
    [String]
    $AppName,

    [Parameter(Mandatory = $false)]
    [String]
    $CertPath,

    [Parameter(Mandatory = $false)]
    [String]
    $TenantId,

    [Parameter(Mandatory = $true)]
    [String[]] $scopes,

    [Parameter(Mandatory = $false)]
    [Switch]
    $UseClientSecret,

    [Parameter(Mandatory = $false)]
    [Switch]
    $StayConnected = $false
)
# Required modules
if ( -not(Get-module "Microsoft.Graph.Applications" -ListAvailable) ) {
    Install-Module "Microsoft.Graph.Applications" -Scope CurrentUser -Force
}
Import-Module "Microsoft.Graph.Applications"

# Graph permissions variables
$graphResourceId = "00000003-0000-0000-c000-000000000000"
$scopesArray = New-Object System.Collections.ArrayList
# Looking for each scope listed, and finding its permission ID to add to the array.
foreach ($sc in $scopes) {
    New-Variable perm -Value @{
        Id   = (Find-MgGraphPermission -SearchString $sc -PermissionType Application -ExactMatch).id 
        Type = "Role"
    }
    $null = $scopesArray.add($perm)
    remove-variable perm
}

# Requires an admin
if ($TenantId) {
    Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read" -TenantId $TenantId
}
else {
    Connect-MgGraph -Scopes "Application.ReadWrite.All User.Read"
}

# Get context for access to tenant ID
$context = Get-MgContext

# Load cert
if ( $CertPath ) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertPath)
    Write-Host -ForegroundColor Cyan "Certificate loaded"
}
elseif ( -not($UseClientSecret) ) {
    # Create certificate
    $mycert = New-SelfSignedCertificate -DnsName $context.Account.Split("@")[1] -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(1) -KeySpec KeyExchange

    # Export certificate to .pfx file
    $mycert | Export-PfxCertificate -FilePath mycert.pfx -Password (ConvertTo-SecureString -String "LS1setup!" -AsPlainText -Force ) -Force

    # Export certificate to .cer file
    $mycert | Export-Certificate -FilePath mycert.cer -Force
    $cerPath = Get-ChildItem -Path ".\mycert.cer"
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath.FullName)
}

# Create app registration
if ( -not($UseClientSecret) ) {
    $appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
        -Web @{ RedirectUris = "http://localhost"; } `
        -RequiredResourceAccess @{ ResourceAppId = $graphResourceId; ResourceAccess = $scopesArray.ToArray() } `
        -AdditionalProperties @{} -KeyCredentials @(@{ Type = "AsymmetricX509Cert"; Usage = "Verify"; Key = $cert.RawData })
}
else {
    $appRegistration = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" `
        -Web @{ RedirectUris = "http://localhost"; } `
        -RequiredResourceAccess @{ ResourceAppId = $graphResourceId; ResourceAccess = $scopesArray.ToArray() } `
        -AdditionalProperties @{}

    $appObjId = Get-MgApplication -Filter "AppId eq '$($appRegistration.Appid)'"
    $passwordCred = @{
        displayName = 'Secret created in PowerShell'
        endDateTime = (Get-Date).Addyears(1)
    }
    $secret = Add-MgApplicationPassword -applicationId $appObjId.Id -PasswordCredential $passwordCred
}
Write-Host -ForegroundColor Cyan "App registration created with app ID" $appRegistration.AppId

# Create corresponding service principal
New-MgServicePrincipal -AppId $appRegistration.AppId -AdditionalProperties @{} | Out-Null
Write-Host -ForegroundColor Cyan "Service principal created"
Write-Host
Write-Host -ForegroundColor Green "Success"
Write-Host

# Generate admin consent URL
$adminConsentUrl = "https://login.microsoftonline.com/" + $context.TenantId + "/adminconsent?client_id=" + $appRegistration.AppId
Write-Host -ForeGroundColor Yellow "Please go to the following URL in your browser to provide admin consent"
Write-Host $adminConsentUrl
Write-Host

# Generate Connect-MgGraph command
if ( -not($UseClientSecret)) {
    $connectGraph = "Connect-MgGraph -ClientId """ + $appRegistration.AppId + """ -TenantId """`
        + $context.TenantId + """ -CertificateThumbprint """ + $cert.Thumbprint + """"
}
else {
    $connectGraph = @"
###########################################################################
if ( -not(Get-module "MSAL.PS" -ListAvailable) ) {
    Install-Module "MSAL.PS" -Scope CurrentUser -Force
}
Import-Module MSAL.PS

`$token = get-MsalToken -ClientID $($appRegistration.AppId) -TenantID $($context.TenantId) -ClientSecret (ConvertTo-SecureString -string "$($secret.SecretText)" -AsPlainText -Force)

Connect-MgGraph -AccessToken `$Token.AccessToken
###########################################################################
"@
    }
Write-Host -ForeGroundColor Cyan "After providing admin consent, you can use the following command to Connect to MgGraph using app-only:"
Write-Host $connectGraph

if ($StayConnected -eq $false) {
    Write-Host
    $null = Disconnect-MgGraph
    Write-Host "Disconnected from Microsoft Graph"
}
else {
    Write-Host
    Write-Host -ForegroundColor Yellow "The connection to Microsoft Graph is still active. To disconnect, use Disconnect-MgGraph"
}