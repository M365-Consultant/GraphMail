# Send-M365cdeMail function
function Send-M365cdeMail(){
<#
.SYNOPSIS
This function can be used for sending e-mails over Microsoft Graph.

.DESCRIPTION
This function can be used for sending e-mails over Microsoft Graph. It is using the Microsoft.Graph module with the cmdlet Send-MgUserMail to send e-mails over Microsoft Exchange Online. It could be used within a script or a Azure Automation Runbook.
There are multiple parameters available to customize the mail. The sender, recipient, CC, BCC, subject, content, attachment, importance, and reply to address can be customized.
The script requires the Microsoft.Graph module to be imported and an active connection to Microsoft Graph. The identity used to send the mail must have the necessary permission 'Mail.Send' to send the mail.
#>

# Define the input parameters
[CmdletBinding(SupportsShouldProcess)]
Param
(
    [Parameter (Mandatory= $true)]
    [String] $Sender,   # the sender of the mail
    [Parameter (Mandatory= $true)]
    [psobject] $Recipient,    # the recipient of the mail - multiple recipients can be separated by a semicolon
    [Parameter (Mandatory= $false)]
    [psobject] $CC,   # the CC recipient of the mail - multiple recipients can be separated by a semicolon
    [Parameter (Mandatory= $false)]
    [psobject] $BCC,  # the BCC recipient of the mail - multiple recipients can be separated by a semicolon
    [Parameter (Mandatory= $true)]
    [String] $Subject,  # the subject of the mail
    [Parameter (Mandatory= $true)]
    [String] $Content,  # the content of the mail (HTML content is supported)
    [Parameter (Mandatory= $false)]
    [psobject] $AttachmentPath, # the name of the attachment
    [Parameter (Mandatory= $false)]
    [switch] $SaveToSentItems  = $false,   # if the mail should be saved to the sent items folder of the senders mailbox
    [Parameter (Mandatory= $false)]
    [ValidateSet("Normal", "High", "Low")] $Importance = "Normal"  # the importance of the mail
)

#Check if the Microsoft Graph module is imported
If (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Error "Microsoft.Graph.Authentication module is not imported. Please import the Microsoft.Graph.Authentication module first."
    break
}
If (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions)) {
    Write-Error "Microsoft.Graph.Users.Actions module is not imported. Please import the Microsoft.Graph.Users.Actions module first."
    break
}

# Check if there is an active connection to Microsoft Graph
If (-not (Get-MgContext)) {
    Write-Error "No active connection to Microsoft Graph. Please connect to Microsoft Graph first."
    break
}

# Create the mail recipients array
if ($null -eq $Recipient) {
    # Handle null value
    $mailRecipientArray = $null
} elseif  (($Recipient | Get-Member -MemberType Properties).Count -gt 1) {
    # Handle multiple properties (object)
    Write-Error "The input on parameter Recipient is an object with multiple properties. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
} elseif ($Recipient -is [array]) {
    # Handle multiple recipients (array)
    $mailRecipientArray = $Recipient
} elseif ($Recipient -is [string]) {
    # Handle single recipient (string)
    $mailRecipientArray = $Recipient.Split(";")
} else {
    # Handle other cases (e.g., object)
    Write-Error "Unexpected data type for recipients. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
}

# Create the mail CC recipients array
if ($null -eq $CC) {
    # Handle null value
    $mailCCArray = $null
} elseif  (($CC | Get-Member -MemberType Properties).Count -gt 1) {
    # Handle multiple properties (object)
    Write-Error "The input on parameter CC is an object with multiple properties. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
} elseif ($CC -is [array]) {
    # Handle multiple recipients (array)
    $mailCCArray = $CC
} elseif ($CC -is [string]) {
    # Handle single recipient (string)
    $mailCCArray = $CC.Split(";")
} else {
    # Handle other cases (e.g., object)
    Write-Error "Unexpected data type for CC recipients. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
}

# Create the mail BCC recipients array
if ($null -eq $BCC) {
    # Handle null value
    $mailBCCArray = $null
} elseif  (($BCC | Get-Member -MemberType Properties).Count -gt 1) {
    # Handle multiple properties (object)
    Write-Error "The input on parameter BCC is an object with multiple properties. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
} elseif ($BCC -is [array]) {
    # Handle multiple recipients (array)
    $mailBCCArray = $BCC
} elseif ($BCC -is [string]) {
    # Handle single recipient (string)
    $mailBCCArray = $BCC.Split(";")
} else {
    # Handle other cases (e.g., object)
    Write-Error "Unexpected data type for BCC recipients. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
}

# Create the attachment data for each provided attachment path
if ($null -eq $AttachmentPath) {
    # Handle null value
    $AttachmentArray = $null
} elseif  (($AttachmentPath | Get-Member -MemberType Properties).Count -gt 1) {
    # Handle multiple properties (object)
    Write-Error "The input on parameter AttachmentPath is an object with multiple properties. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
} elseif ($AttachmentPath -is [array]) {
    # Handle multiple attachments (array)
    $AttachmentArray = $AttachmentPath
} elseif ($AttachmentPath -is [string]) {
    # Handle single attachment (string)
    $AttachmentArray = $AttachmentPath.Split(";")
} else {
    # Handle other cases (e.g., object)
    Write-Error "Unexpected data type for attachment path. Use string (semicolon-separated), a object with a single propertie for the mail or array."
    break
}

# Create the parameters for the mail
$params = @{
    Message = @{
        Subject = $Subject
        Body = @{
            ContentType = "html"
            Content = $Content
        }
        ToRecipients = @(
            foreach ($recipient in $mailRecipientArray) {
                @{
                    EmailAddress = @{
                        Address = $recipient
                    }
                }
            }
        )
    }
    SaveToSentItems = If ($SaveToSentItems) { "true" } Else { "false" }
}

# Add CcRecipients if the array exists
if ($mailCCArray) {
    $params.Message.CcRecipients = @(
        foreach ($cc in $mailCCArray) {
            @{
                EmailAddress = @{
                    Address = $cc
                }
            }
        }
    )
}

# Add BccRecipients if the array exists
if ($mailBCCArray) {
    $params.Message.BccRecipients = @(
        foreach ($bcc in $mailBCCArray) {
            @{
                EmailAddress = @{
                    Address = $bcc
                }
            }
        }
    )
}

# Add the attachment if the attachment data is provided
if ($AttachmentArray) {
    $params.Message.Attachments = @(
        foreach ($file in $AttachmentArray){
            $fileName = (Get-Item -Path $file).Name
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name = $fileName
                ContentType = "text/plain"
                ContentBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file))
            }
        }
    )
}

# Add the importance of the mail
$params.Message.Importance = $Importance

# Send the mail
if ($PSCmdlet.ShouldProcess("Mail send operation: `n" +
                            "From $Sender `n" +
                            "To: $mailRecipientArray `n" +
                            "CC: $mailCCArray `n" +
                            "BCC: $mailBCCArray `n" +
                            "Subject: $Subject")) {
    $result = Send-MgUserMail -UserId $Sender -BodyParameter $params -PassThru
    # Check if the mail has been sent successfully
    If ($result) {
        Write-Output "Mail has been sent successfully."
    }
    else {
        Write-Error "Mail could not be sent. Please check the parameters and try again."
    }
}


}


Export-ModuleMember -Function Send-M365cdeMail