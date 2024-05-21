# M365cde - GraphMail
This function can be used for sending e-mails over Microsoft Graph. It is using the Microsoft.Graph module with the cmdlet Send-MgUserMail to send e-mails over Microsoft Exchange Online. It could be used within a script or a Azure Automation Runbook.
There are multiple parameters available to customize the mail. The sender, recipient, CC, BCC, subject, content, attachment, importance, and reply to address can be customized. 

_Feel free to share any feature requests with me!_

# Dependencies
The module requires this modules to be installed and imported:
- Microsoft.Graph.Authentication
- Microsoft.Graph.Users.Actions

# Installation
The module is published on [PowerShell Gallery](https://www.powershellgallery.com/packages/M365cde.GraphMail/) and can be installed with this command within a powershell console:

    Install-Module -Name M365cde.GraphMail -Scope CurrentUser

# Usage
To use the module u can use the function
```
Send-M365cdeMail -Sender [string] -Recipient [psobject] -CC [psobject] -BCC [psobject] -Subject [string] -Content [string] -AttachmentPath [psobject] -SaveToSentItems [switch] -Importance [string]
```

There are multiple parameters available to customize the mail:

The mandatory parameters are:
- Sender (the sender of the mail - must be a single e-mail address provided  as a string from a existing user-mailbox or shared-mailbox)
- Recipient (the recipient of the mail - can be multiple recipients as a semicolon-separated string or an array/object with a single property)
- Subject (the subject of the mail)
- Content (the content of the mail - HTML content is supported)

The optional parameters are:
- CC (the CC recipient of the mail - can be multiple recipients as a semicolon-separated string or an array/object with a single property)
- BCC (the BCC recipient of the mail - can be multiple recipients as a semicolon-separated string or an array/object with a single property)
- AttachmentPath (the path to the attachment - can be multiple attachments as a semicolon-separated string or an array/object with a single property)
- SaveToSentItems (boolean switch to save the mail to the sent items folder of the sender's mailbox, default is false)
- Importance (the importance of the mail - can be 'Normal', 'High', or 'Low', default is 'Normal')

Sample usage:
```
Send-M365cdeMail -Sender 'john.doe@contoso.com' -Recipient $mailRecipient -Subject 'Report' -Content "<b>Attached you'll find the report.</b>" -AttachmentPath 'C:\Temp\Report.pdf' -SaveToSentItems -Importance 'High'
```

# Changelog
- v1.0.1 Changed Module-Check to
  - Change onto Microsoft.Graph.Authentication and Microsoft.Graph.Users.Actions, so only those are required to run the function.
- v1.0.0 First final release
  - Release after testing the functions with multiple inputs
  - With this release the module is available on PSGallery API 2.0
- v0.0.1 First release
  - First release of this script
