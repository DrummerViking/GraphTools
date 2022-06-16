<#
    .SYNOPSIS
    Script to export meeting items using Graph via Powershell.
    
    .DESCRIPTION
    Script to export meeting items using Graph via Powershell.
    It can run on a single mailbox, or multiple mailboxes.

    If it runs on a single mailbox, the module can pop-up and request the authenticated user to consent Graph permissions. The script will run against the authenticated mailbox.
    If it runs against multiple mailboxes, an AzureAD Registered App is needed, with the appropriate Application permissions (requires 'Calendars.Read' API permission granted).

    .PARAMETER ClientID
    This is an optional parameter. String parameter with the ClientID (or AppId) of your AzureAD Registered App.
    
    .PARAMETER TenantID
    This is an optional parameter. String parameter with the TenantID your AzureAD tenant.
    
    .PARAMETER CertificateThumbprint
    This is an optional parameter. Certificate thumbprint which is uploaded to the AzureAD App.
    
    .PARAMETER Mailboxes
    This is an optional parameter. This is a list of SMTP Addresses. If this parameter is ommitted, the script will run against the authenticated user mailbox.
    
    .PARAMETER ExportFolderPath
    Insert target folder path named like "C:\Temp". By default this will be "$home\desktop"

    .PARAMETER StartDate
    This is an optional parameter. The script will search for meeting items starting based on this StartDate onwards. If this parameter is ommitted, by default will consider 1 year backwards from the current date.
    
    .PARAMETER EndDate
    This is an optional parameter. The script will search for meeting items ending based on this EndDate backwards. If this parameter is ommitted, by default will consider 1 year forward from the current date.
    
    .PARAMETER DisableTranscript
    This is an optional parameter. Transcript is enabled by default. Use this parameter to not write the powershell Transcript.
    
    .PARAMETER DisconnectMgGraph
    This is an optional parameter. Use this parameter to disconnect from MgGraph when it finishes.
    
    .EXAMPLE
    PS C:\> .\Export-GraphUserCalendarEvents.ps1 -StartDate 06/20/2022 -Verbose
    The script will install required modules if not already installed.
    Later it will request the user credential, and ask for permissions consent if not granted already.
    Then it will search for all meeting items matching the startDate on 06/20/2022 forward.
    It will export the items found to the default ExportFolderPath in files "<alias>-CalendaritemsReport.csv".

    .EXAMPLE
    PS C:\> $mailboxes = Get-EXOMailbox -Filter {Office -eq "Staff"} -Properties PrimarySMTPAddress | Select-Object PrimarySMTPAddress
    PS C:\> .\Export-GraphUserCalendarEvents.ps1 -Mailboxes $mailboxes.PrimarySMTPAddress -ClientID "12345678" -TenantId "abcdefg" -CertificateThumbprint "a1b2c3d4" -Verbose
    The script will install required modules if not already installed.
    Later it will connect to MgGraph using AzureAD App details (requires ClientID, TenantID and CertificateThumbprint).
    Then it will search for all meeting items matching default StartDate and EndDate, for all mailboxes belonging to the "Staff" Office.
    It will export the items found to the default ExportFolderPath in files "<alias>-CalendaritemsReport.csv".

    .NOTES
    Author: Agustin Gallegos
    #>
    [CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
    param (
        [String] $ClientID,
    
        [String] $TenantID,
    
        [String] $CertificateThumbprint,
    
        [String[]] $Mailboxes,
    
        [DateTime] $StartDate = (Get-date).AddYears(-1),
    
        [DateTime] $EndDate = (Get-date).AddYears(1),

        [String] $ExportFolderPath = "$home\Desktop",

        [Switch] $DisableTranscript,
    
        [Switch] $DisconnectMgGraph
    )
 
    begin {
        $disclaimer = @"
#################################################################################
#
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss 
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
# 
#################################################################################
"@
        Write-Host $disclaimer -foregroundColor Yellow
        Write-Host " "

        if ( -not($DisableTranscript) ) {
            Start-Transcript
        }

        # creating folder path if it doesn't exists
        if ( $ExportFolderPath -ne "$home\Desktop\" ) {
            if ( -not (Test-Path $ExportFolderPath) ) {
                write-host "Folder '$ExportFolderPath' does not exists. Creating folder." -foregroundColor Green
                $null = New-Item -Path $ExportFolderPath -ItemType Directory -Force
            }
        }
        else {
            # Checking if Desktop folder is located in the user's profile folder, or synched to OneDrive
            if ( -not(Test-Path $ExportFolderPath) ) {
                $ExportFolderPath = "$env:OneDriveCommercial\Desktop"
            }
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
        if ( $null -eq $conn -or $conn.Scopes -notcontains "Calendars.Read" ) {
            Write-Verbose "There is currently no active connection to MgGraph or current connection is missing required 'Calendars.Read' Scope."
            if ( -not($PSBoundParameters.ContainsKey('Mailboxes')) ) {
                # Connecting to graph with the user account
                Write-Verbose "Connecting to graph with the user account"
                Connect-MgGraph -Scopes "Calendars.Read"
            }
            else {
                # Connecting to graph using Azure App Application flow
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
            $i++
            Write-Progress -Id 0 -activity "Processing Users: $i out of $($mbxs.Count)" -status "Percent scanned: " -PercentComplete ($i * 100 / $($mbxs.Count)) -ErrorAction SilentlyContinue
            Write-Verbose "Working on mailbox $mb"

            Write-Verbose "Collecting events between $StartDate and $EndDate"
            $eventsFound = Get-MgUserCalendarView -UserId $mb -EndDateTime $EndDate -StartDateTime $StartDate -All

            if ( $eventsFound.Count -eq 0 ) {
                Write-Verbose "No events found based on parameters criteria. Please double check and try again."
                Continue
            }
            # Exporting found events to Verbose deleting
            Write-Verbose "Exporting events to $ExportFolderPath"
            $eventsFound | Select-Object subject,@{N="organizer";E={$_.Organizer.EmailAddress.Address}},@{N="Attendees";E={$_.Attendees | ForEach-Object {$_.EmailAddress.Address -join ";"}}},@{N="location";E={$_.location.DisplayName}},@{N="StartTime";E={$_.Start.DateTime}},@{N="EndTime";E={$_.End.DateTime}},type,id | Export-csv "$ExportFolderPath\$($mb.split("@")[0])-CalendaritemsReport.csv" -NoTypeInformation -Force
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