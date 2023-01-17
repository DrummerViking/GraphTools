<#
    .SYNOPSIS
    Script to send email messages through MS Graph using Powershell.
    
    .DESCRIPTION
    Script to send email messages through MS Graph using Powershell.
    
    .PARAMETER ToRecipients
    List of recipients in the "To" list. This is a Mandatory parameter.
    
    .PARAMETER CCRecipients
    List of recipients in the "CC" list. This is an optional parameter.

    .PARAMETER BccRecipients
    List of recipients in the "Bcc" list. This is an optional parameter.

    .PARAMETER Subject
    Use this parameter to set the subject's text. By default will have: "Test message sent via Graph".

    .PARAMETER Body
    Use this parameter to set the body's text. By default will have: "Test message sent via Graph using Powershell".
    
    .PARAMETER DisconnectMgGraph
    Use this optional parameter to disconnect from MgGraph after sending the email message.
    
    .EXAMPLE
    PS C:\> .\Send-GraphMailMessage.ps1 -ToRecipients "john@contoso.com"

    The script will download (if not already installed) the required Graph powershell modules.
    Will authenticate to MS Graph (if not already), and will prompt for consent if it was not already granted.
    Then will send the email message to "john@contoso.com" from the user previously authenticated.

    .EXAMPLE
    PS C:\> .\Send-GraphMailMessage.ps1 -ToRecipients "julia@contoso.com","carlos@contoso.com" -BccRecipients "mark@contoso.com" -Subject "Lets meet!"

    The script will download (if not already installed) the required Graph powershell modules.
    Will authenticate to MS Graph (if not already), and will prompt for consent if it was not already granted.
    Then will send the email message to "julia@contoso.com" and "carlos@contoso.com" and bcc to "mark@contoso.com", from the user previously authenticated.

    .NOTES
    Author: Vlad Radu
    Collaborator: Agustin Gallegos
#>
[Cmdletbinding()]
Param (
    [parameter(Mandatory = $true)]
    [String[]] $ToRecipients,

    [String[]] $CCRecipients,

    [String[]] $BccRecipients,

    [String] $Subject = "Test message sent via Graph",

    [String] $Body = "Test message sent via Graph using Powershell",

    [Switch] $DisconnectMgGraph
)
Begin {
    # Downloading required Graph modules
    @(
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Users.Actions'
        'Microsoft.Graph.Authentication'
    ) | ForEach-Object {
        if ( -not(Get-Module $_ -ListAvailable)) {
            Write-Verbose "'$_' Module not found. Installing it..."
            Install-Module $_ -Scope CurrentUser -Force
        }
    }
    Import-Module Microsoft.Graph.Users, Microsoft.Graph.Users.Actions

    # Connect to Graph if there is no current context
    $conn = Get-MgContext
    if ( $null -eq $conn -or $conn.Scopes -notcontains "Mail.Send" ) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] There is currently no active connection to MgGraph or current connection is missing required 'Mail.Send' Scope."
        Connect-MgGraph -Scopes "Mail.Send"        
    }
}
Process {
    # Base mail body Hashtable
    $Global:MailBody = @{
        Message         = @{
            Subject = $Subject;
            Body    = @{
                Content     = $Body; 
                ContentType = "HTML"
            }
        }
        savetoSentItems = "true"
    }

    # looping through each recipient in the list, and adding it in the hash table
    $recipientsList = New-Object System.Collections.ArrayList
    foreach ( $recipient in $ToRecipients) {
        $null = $recipientsList.add(
            @{
                EmailAddress = @{
                    Address = $recipient
                }
            }
        )
    }
    $MailBody.Message.Add("ToRecipients", $recipientsList)

    # looping through each recipient in the CC list, and adding it in the hash table
    if ( $CCRecipients.Count -gt 0 ) {
        $ccRecipientsList = New-Object System.Collections.ArrayList
        foreach ( $cc in $CCRecipients) {
            $null = $ccRecipientsList.add(
                @{
                    EmailAddress = @{
                        Address = $cc
                    }
                }
            )
        }
        $MailBody.Message.Add("CcRecipients", $ccRecipientsList)
    }

    # looping through each recipient in the Bcc list, and adding it in the hash table
    if ( $BccRecipients.Count -gt 0 ) {
        $BccRecipientsList = New-Object System.Collections.ArrayList
        foreach ( $bcc in $BccRecipients) {
            $null = $BccRecipientsList.add(
                @{
                    EmailAddress = @{
                        Address = $bcc
                    }
                }
            )
        }
        $MailBody.Message.Add("BccRecipients", $BccRecipientsList)
    }

    # Making Graph call to send email message
    Send-MgUserMail -UserId (Get-MgContext).Account -BodyParameter $MailBody
    if ( $? ) {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Successfully sent the email message using graph."
    }
    else {
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Something failed to send the email message using graph. Error message: $($Error[-1].exception.message)"
    }
}
End {
    if ( $DisconnectMgGraph ) {
        Disconnect-MgGraph
        Write-Host "[$((Get-Date).ToString("HH:mm:ss"))] Disconneting from MS Graph."
    }
}