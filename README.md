# Powershell Graph tools
Powershell Graph tools mostly use in Exchange Online (Office 365) and AzureAD.  
Most of these apps will require to register an AzureAD App Registration in order to work in bulk. If that is the case, you can use the 1st script to register an App and later use this app for the rest of the scripts.  
These scripts will require the AzureAD App with a registered Certificate to authenticate. You can still rely on the 1st script offered here to create the app and create the certificate for yourself, or you can create your app and certificate manually following [this doc](https://learn.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate)  

1. [Register-AzureADApp](#register-azureadapp)
2. [Remove-GraphUserCalendarEvents](#remove-graphusercalendarevents)
3. [Export-GraphUserCalendarEvents](#export-graphusercalendarevents)
4. [Send-GraphMailMessage](#send-graphmailmessage)
5. [Update-GraphUserCalendarEventsEndDate](#update-graphusercalendareventsenddate)

## Register-AzureADApp
Script to create a new AzureAD App registration.  

[More Info](/Register-AzureADApp/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/Graphtools/master/Register-AzureADApp/Register-AzureADApp.ps1)  

----

## Remove-GraphUserCalendarEvents
Script to delete meeting items using Graph via Powershell.  
It can run on a single mailbox, or multiple mailboxes.  

[More Info](/Remove-GraphUserCalendarEvents/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/Graphtools/master/Remove-GraphUserCalendarEvents/Remove-GraphUserCalendarEvents.ps1)  

----

## Export-GraphUserCalendarEvents
Script to export calendar items to CSV using Graph via Powershell.  
It can run on a single mailbox, or multiple mailboxes.  

Take into account that Graph Api does not support connecting to Archive mailboxes, so if you need to access and export archive items, you can rely on [EWS](https://github.com/agallego-css/tools#export-calendar-items-exo). More info [here](https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)  

The report exports the following columns:  
> Subject, Organizer, Attendees, Location, Start Time, End Time, Type, ItemId  

[More info](/Export-GraphUserCalendarEvents/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/Graphtools/master/Export-GraphUserCalendarEvents/Export-GraphUserCalendarEvents.ps1)  

----

## Send-GraphMailMessage

Script to send email messages through MS Graph using Powershell.

[More Info](/Send-GraphMailMessage/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/Graphtools/master/send-GraphMailMessage/Send-GraphMailMessage.ps1)

----

## Update-GraphUserCalendarEventsEndDate  

Script to update recurring meeting with no end date to have an end date using Graph via Powershell.  
[More Info](/Update-GraphUserCalendarEvents/) - [Download (Right click and select 'Save link as')](https://raw.githubusercontent.com/agallego-css/Graphtools/master/Update-GraphUserCalendarEvents/Update-GraphUserCalendarEventsEndDate.ps1)