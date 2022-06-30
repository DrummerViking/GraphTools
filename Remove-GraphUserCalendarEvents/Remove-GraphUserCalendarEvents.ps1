<#
    .SYNOPSIS
    Script to delete meeting items using Graph via Powershell.
    
    .DESCRIPTION
    Script to delete meeting items using Graph via Powershell.
    It can run on a single mailbox, or multiple mailboxes.

    If it runs on a single mailbox, the module can pop-up and request the authenticated user to consent Graph permissions. The script will run against the authenticated mailbox.
    If it runs against multiple mailboxes, an AzureAD Registered App is needed, with the appropriate Application permissions (requires 'Calendars.ReadWrite' API permission granted).

    If the event is a meeting, deleting the event on the organizer's calendar sends a cancellation message to the meeting attendees.
    
    .PARAMETER ClientID
    This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.
    
    .PARAMETER TenantID
    This is an optional parameter. String parameter with the TenantID your AzureAD tenant.
    
    .PARAMETER CertificateThumbprint
    This is an optional parameter. Certificate thumbprint which is uploaded to the AzureAD App.
    
    .PARAMETER Subject
    This is an mandatory parameter. The exact subject text to filter meeting items. This parameter cannot be used together with the "FromAddress" parameter.
    
    .PARAMETER FromAddress
    This is an mandatory parameter. The sender address to filter meeting items. This parameter cannot be used together with the "Subject" parameter.
    
    .PARAMETER Mailboxes
    This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.
    
    .PARAMETER StartDate
    This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider the current date.
    
    .PARAMETER EndDate
    This is an optional parameter. The script will search for meeting items ending based on this EndDate backwards. If this parameter is ommitted, by default will consider 1 year forward from the current date.

    .PARAMETER DisableTranscript
    This is an optional parameter. Transcript is enabled by default. Use this parameter to not write the powershell Transcript.

    .PARAMETER ListOnly
    This is an optional parameter. Use this parameter to list the calendar events found, without deleting them. This is a good parameter to use, to actually see the current found events and double check these are the ones to be deleted.
    
    .PARAMETER DisconnectMgGraph
    This is an optional parameter. Use this parameter to disconnect from MgGraph when it finishes.
    
    .EXAMPLE
    PS C:\> .\Remove-GraphUserCalendarEvents.ps1 -Subject "Yearly Team Meeting" -StartDate 06/20/2022 -Verbose
    
    The script will install required modules if not already installed.
    Later it will request the user credential, and ask for permissions consent if not granted already.
    Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on 06/20/2022 forward.
    It will display the items found and proceed to remove them.

    .EXAMPLE
    PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
    PS C:\> .\Remove-GraphUserCalendarEvents.ps1 -Subject "Yearly Team Meeting" -Mailboxes $mailboxes.PrimarySMTPAddress -ClientID "12345678" -TenantId "abcdefg" -CertificateThumbprint "a1b2c3d4" -Verbose
    
    The script will install required modules if not already installed.
    Later it will connect to MgGraph using AzureAD App details (requires ClientID, TenantID and CertificateThumbprint).
    Then it will search for all meeting items matching exact subject "Yearly Team Meeting" starting on the current date forward, for all mailboxes belonging to the "Staff" Office.
    It will display the items found and proceed to remove them.

    .NOTES
    Author: Agustin Gallegos
    #>
[CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
param (
    [String] $ClientID,

    [String] $TenantID,

    [String] $CertificateThumbprint,

    [parameter(ParameterSetName="Subject")]
    [String] $Subject,

    [parameter(ParameterSetName="FromAddress")]
    [String] $FromAddress,

    [String[]] $Mailboxes,

    [DateTime] $StartDate = (Get-date),

    [DateTime] $EndDate = (Get-date).AddYears(1),

    [Switch] $DisableTranscript,

    [Switch] $ListOnly,

    [Switch] $DisconnectMgGraph
)
    
begin {
    if ( -not($DisableTranscript) ) {
        Start-Transcript
    }
    # Downloading required Graph modules
    if ( -not(Get-Module Microsoft.Graph.Users -ListAvailable)) {
        Write-Verbose "'Microsoft.Graph.Users' Module not found. Installing it..."
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    }
    if ( -not(Get-Module Microsoft.Graph.Calendar -ListAvailable)) {
        Write-Verbose "'Microsoft.Graph.Calendar' Module not found. Installing it..."
        Install-Module Microsoft.Graph.Calendar -Scope CurrentUser -Force
    }
    if ( -not(Get-Module Microsoft.Graph.Authentication -ListAvailable)) {
        Write-Verbose "'Microsoft.Graph.Authentication' Module not found. Installing it..."
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph.Users, Microsoft.Graph.Calendar -Verbose:$false

    # Connect to Graph if there is no current context
    $conn = Get-MgContext
    if ( $null -eq $conn -or $conn.Scopes -notcontains "Calendars.ReadWrite" ) {
        Write-Verbose "There is currently no active connection to MgGraph or current connection is missing required 'Calendars.ReadWrite' Scope."
        if ( -not($PSBoundParameters.ContainsKey('Mailboxes')) ) {
            # Connecting to graph with the user account
            Write-Verbose "Connecting to graph with the user account"
            Connect-MgGraph -Scopes "Calendars.ReadWrite"
        }
        else {
            # Connecting to graph using Azure App
            if ( $clientID -eq '' -or $TenantID -eq '' -or $CertificateThumbprint -eq '' ) {
                Write-Host "ERROR: Required 'ClientID', 'TenantID' and 'CertificateThumbprint' parameters are missing to connect using App Authentication." -ForegroundColor Red
                Exit
            }
            Write-Verbose "Connecting to graph with Azure AppId: $ClientID"
            Connect-MgGraph -ClientId $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
        }
    }
    else {
        if ( $null -eq $conn.Account ){
            Write-Verbose "Currently connect with App Account: $($conn.AppName)"
        }
        else {
            Write-Verbose "Currently connected with User Account: $($conn.Account)"
        }
    }
    $mbxs = (Get-MgContext).Account
    if ( $PSBoundParameters.ContainsKey('Mailboxes') ) {
        $mbxs = $Mailboxes
    }
}

process {

    $i = 0
    foreach ( $mb in $mbxs ) {
        #$events = New-Object System.Collections.ArrayList
        $i++
        Write-Progress -activity "Scanning Users: $i out of $($mbxs.Count)" -status "Percent scanned: " -PercentComplete ($i * 100 / $($mbxs.Count)) -ErrorAction SilentlyContinue
        Write-Verbose "Working on mailbox $mb"
        switch ($PSBoundParameters.Keys) {
            Subject {
                Write-Verbose "Collecting events based on exact subject: '$Subject' between $startDate and $endDate."
                $events = Get-MgUserCalendarView -UserId $mb -Filter "Subject eq '$subject'" -StartDateTime $StartDate -EndDateTime $EndDate -All
            }
            FromAddress {
                Write-Verbose "Collecting events based on sender: '$FromAddress' between $startDate and $endDate."
                $events = Get-MgUserCalendarView -UserId $mb -StartDateTime $StartDate -EndDateTime $EndDate -all | Where-Object { $_.Organizer.EmailAddress.Address -eq "$FromAddress" } 
            }
        }
        if ( $events.Count -eq 0 ) {
            Write-Verbose "No events found based on parameters criteria. Please double check and try again."
            Continue
        }
        # Exporting found events to Verbose deleting
        if ( $PSBoundParameters.ContainsKey('Verbose') ) {
            Write-Verbose "Displaying events details:"
            $events | Select-Object subject,@{N="Mailbox";E={$mb}},@{N="organizer";E={$_.Organizer.EmailAddress.Address}},@{N="Attendees";E={$_.Attendees | ForEach-Object {$_.EmailAddress.Address -join ";"}}},@{N="StartTime";E={$_.Start.DateTime}},@{N="EndTime";E={$_.End.DateTime}},id
        }
        if ( -not($ListOnly) ) {
            foreach ( $event in $events ) {
                Write-Host "Removing event item from '$($event.Organizer.EmailAddress.Address)' with subject '$($event.Subject)' and item ID '$($event.id)'"
                Remove-MgUserEvent -UserId $mb -EventId $event.id
            }
        }
    }
}

end {
    if ( -not($DisableTranscript) ) {
        Stop-Transcript
    }

    if ( $DisconnectMgGraph ) {
        Disconnect-MgGraph
    }
}