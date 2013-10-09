<#
.SYNOPSIS
    Sends an email to users whos passwords soon expire (or already have)
.DESCRIPTION
    This script searches through AD to find users whos passwords will soon expire
    (or already have) by comparing when their password was set to the password age
    policy.  The script then sends that user an email notifying them when their
    password will expire and how to change their password, as well as a report
    to an administrator
.NOTES
    File Name : adpen.ps1
    Author    : Richard Fleming - me@richardfleming.me
    Requires  : ActiveDirectory Module
    Resources : Email template part of https://github.com/mailchimp/Email-Blueprints and modified
    Release   : 1
.LINK
    http://richardfleming.me/active-directory-password-expiry-notifification/
    http://github.com/richardfleming/adpen/
#>

# --- Editable Variables, later assigned into $ghtSettings ---

# This variable sets how many days in advance (of password expiration) that email warnings will be sent
$pwdNotificationStartInDays = 15

# Enable or disable reporting
$reportsEnabled = $True

# Mail Server
$smtpServer                 = 'mailserver.example.com'

# These three variables define the contact details in the HTML email for help
$emailContactAddress        = 'support@example.com'
$emailContactDisplayName    = 'IS Support'
$emailContactSubject        = 'Help Me!'

# These three variables define reporting
$emailReportAddress         = 'helpdesk@example.com'
$emailReportDisplayName     = 'IS Helpdesk'
$emailReportSubject         = 'Report of expired/expiring users for ' + (Get-Date -Format d)

# These three variables set who the email comes from and who is exempt (SAM user account name format)
$emailFromAddress           = 'no-reply@example.com'
$emailFromDisplayName       = 'Automated Password Expiry Notifier'
$emailExempt                = @(
                                "",
                                "",
                                ""
                              )

# Test mode, limiting queries to your Display Name and sending results to your email.
$testMode = $True
$testAddress = 'jexample@example.com'
$testDisplayName = 'John Example'

# Set this to $False after configuring your variables
$preventScriptExecution = $True

# --- Do not edit below this line --- #

<##
 # Simple Template Engine for PowerShell
 # Complements Brice Lambson
 # http://brice-lambson.blogspot.ca/2012/09/simple-template-engine-for-powershell.html
 #>
function Merge-Tokens($template, $tokens)
{
    return [regex]::Replace(
        $template,
        '\$(?<tokenName>\w+)\$',
        {
            param($match)
             
            $tokenName = $match.Groups['tokenName'].Value
             
            return $tokens[$tokenName]
        })    
}

# Probe AD for the MaxPwdAge value and calculate its value in days
function Get-MaxPwdAge {
    $strDNSRoot = $ghtSettings.Domain.DNSRoot
    $strDN = $ghtSettings.Domain.DistinguishedName
    $strConnection = "LDAP://"+$strDN
    $objAD = [ADSI]$strConnection
    $intMaxPwdAge = -($objAD.ConvertLargeIntegerToInt64($objAD.MaxPwdAge.Value))/(600000000 * 1440)

    Write-Output $intMaxPwdAge
}

function Populate-Report($strDisplayName, $objPwdTimeLeft, $now) {
    $objPwdExpires = $now.Add($objPwdTimeLeft)

    if ($objPwdTimeLeft.TotalHours -le 0) {
        $ghtReports.Expired.Add($strDisplayName, (Get-Date $objPwdExpires -Format 'dd-MMM-yyyy') )
    }
    elseif (($objPwdTimeLeft.TotalHours -gt 0) -and ($objPwdTimeLeft.TotalHours -le 24)) {
        $ghtReports.Today.Add( $strDisplayName, (Get-Date $objPwdExpires -Format 'hh:mm:ss tt zzz') )
    }

}

# Format and Send out notification email to users
function Send-PwdExpiryEmailToUsers ([string]$strUserMail, [string]$strUserDisplayName, $objPwdTimeLeft, $now)  {

    $objPwdExpires = $now.Add($objPwdTimeLeft)
    $strExpireText = "is set to expire in"

    if ($objPwdTimeLeft.TotalHours -gt 24) {
        $strExpireTime = [string]$objPwdTimeLeft.Days + " days"
    }
    elseif (($objPwdTimeLeft.TotalHours -gt 0) -and ($objPwdTimeLeft.TotalHours -le 24)) {
        $strExpireTime = [string]$objPwdTimeLeft.Hours + " hours at " + (Get-Date $objPwdExpires -Format 'hh:mm:ss tt zzz') + " today"
    }
    else {
        $strExpireText = "expired on"
        $strExpireTime = (Get-Date $objPwdExpires -Format 'dddd d MMMM yyyy hh:mm:ss tt zzz')
    }
    
    $strMessageBody = Merge-Tokens (Get-Content .\Templates\notificationEmail.html) @{
        EMAIL_SUBJECT = $ghtSettings.Email.From.Subject; 
        CONTACT_DISPLAYNAME = $ghtSettings.Email.Contact.DisplayName;
        CONTACT_EMAIL = $ghtSettings.Email.Contact.Address;
        CONTACT_SUBJECT = $ghtSettings.Email.Contact.Subject;
        EXPIRE_TEXT = $strExpireText;
        EXPIRE_TIME = $strExpireTime;
        DISPLAY_NAME = $strUserDisplayName;
        BOT_NAME = $ghtSettings.Email.From.DisplayName        
    }

    # Prepare Email

    # MailAddress object and receiver config
    $objRecipient = New-Object System.Net.Mail.MailAddress($strUserMail, $strUserDisplayName)

    # Create Logo Attachment
    $objAttachment = New-Object System.Net.Mail.Attachment ( $ghtSettings.Script.Path + '\Templates\logo.png')
    $objAttachment.ContentDisposition.Inline = $True
    $objAttachment.ContentDisposition.DispositionType = "Inline"
    $objAttachment.ContentType.MediaType = "image/png"
    $objAttachment.ContentId = "logo"
    
    # Create SMTP Server Object
    $objSmtpClient = New-Object System.Net.Mail.SmtpClient
    $objSmtpClient.Host = $ghtSettings.SmtpServer

    # Create MailMessage object and email structure
    $objMailMessage = New-Object System.Net.Mail.MailMessage
    $objMailMessage.Sender = $ghtSettings.Email.Sender
    $objMailMessage.From = $objMailMessage.Sender
    $objMailMessage.To.Add($objRecipient)
    $objMailMessage.Subject = "Your password " + $strExpireText + " " + $strExpireTime
    $objMailMessage.IsBodyHtml = $True
    $objMailMessage.Body = $strMessageBody
    $objMailMessage.Attachments.Add($objAttachment)

    # Send the message
    $objSmtpClient.Send($objMailMessage);

    # Cleanup
    $objAttachment.Dispose()
    $objMailMessage.Dispose()
}

function Send-Report {
    $strExpired = ""
    $strToday = ""

    # If there aren't any expired or expiring users to report, then break
    if (($ghtReports.Expired.Count -eq 0) -and ($ghtReports.Today.Count -eq 0)) {
        break
    }

    if ($ghtReports.Expired.Count -eq 0) {
        $strExpired = "<tr><td>None!</td></tr>"
    } else {
        $ghtReports.Expired.GetEnumerator() | ForEach-Object {
            $strExpired += "<tr><td>" + $_.Key + "</td><td>" + $_.Value + "</td></tr>`n"
        }
    }
                        
    if ($ghtReports.Today.Count -eq 0) {
        $strToday = "<tr><td>None!</td></tr>"
    } else {
        $ghtReports.Today.GetEnumerator() | ForEach-Object {
            $strToday += "<tr><td>" + $_.Key + "</td><td>" + $_.Value + "</td></tr>`n"
        }
    }

    $strMessageBody = Merge-Tokens (Get-Content .\Templates\reportEmail.html) @{
        EMAIL_SUBJECT = $ghtSettings.Email.Report.Subject; 
        EXPIRED_PASSWORDS = $strExpired;
        EXPIRING_TODAY = $strToday;
        BOT_NAME = $ghtSettings.Email.From.DisplayName        
    }

    # Prepare Email

    # MailAddress object and receiver config
    $objRecipient = New-Object System.Net.Mail.MailAddress($ghtSettings.Email.Report.Address, $ghtSettings.Email.Report.DisplayName)

    # Create Logo Attachment
    $objAttachment = New-Object System.Net.Mail.Attachment ( $ghtSettings.Script.Path + '\Templates\logo.png')
    $objAttachment.ContentDisposition.Inline = $True
    $objAttachment.ContentDisposition.DispositionType = "Inline"
    $objAttachment.ContentType.MediaType = "image/png"
    $objAttachment.ContentId = "logo"
    
    # Create SMTP Server Object
    $objSmtpClient = New-Object System.Net.Mail.SmtpClient
    $objSmtpClient.Host = $ghtSettings.SmtpServer

    # Create MailMessage object and email structure
    $objMailMessage = New-Object System.Net.Mail.MailMessage
    $objMailMessage.Sender = $ghtSettings.Email.Sender
    $objMailMessage.From = $objMailMessage.Sender
    $objMailMessage.To.Add($objRecipient)
    $objMailMessage.Subject = $ghtSettings.Email.Report.Subject
    $objMailMessage.IsBodyHtml = $True
    $objMailMessage.Body = $strMessageBody
    $objMailMessage.Attachments.Add($objAttachment)

    # Send the message
    $objSmtpClient.Send($objMailMessage);

    # Cleanup
    $objAttachment.Dispose()
    $objMailMessage.Dispose()
}


# Gets a list of users then checks to see if their passwords are near expiration or expired, then fires off an email.
# If the hashtable $ghtSettings.Email.Exempt is set, then any SamAccountNames listed in there will not receive an email.
function Get-PasswordExpiredUsers {
    $queryAddOn = ""

    # Test Mode query addition
    if ($ghtSettings.Test.Mode -eq $True) {
        $queryAddOn = "-and (DisplayName -like '" + $ghtSettings.Test.DisplayName + "') " 
    }

    $strFilter = "(ObjectClass -eq 'user') -and (l -like '*') " + $queryAddOn + "-and (Enabled -eq 'True')"
    $objUsers = Get-ADUser -Filter $strFilter -Properties SamAccountName, DisplayName, mail, PasswordLastSet
    $intMaxPwdAge = Get-MaxPwdAge

    # Interate through each user object and check if the password is expired, or nearing expiration limit
    ForEach ($objUser in $objUsers) {
        $objPwdExpires = $objUser.PasswordLastSet.AddDays($intMaxPwdAge)
        $now = (Get-Date)
        $objPwdTimeLeft = $objPwdExpires - $now
                    
        # Ignore exempt users
        if ($ghtSettings.Email.Exempt -contains $objUser.SamAccountName) { break }

        # Test mode email redirection
        if ($ghtSettings.Test.Mode) {
            $strSendTo = $ghtSettings.Test.Address
        } else {
            $strSendTo = $objUser.mail
        }

        # Prepare-SMTPMessage if password expires less than or on a certain day
        if ( $objPwdTimeLeft.Days -le $ghtSettings.Pwd.NotificationStartInDays ) {
            Send-PwdExpiryEmailToUsers $strSendTo $objUser.DisplayName $objPwdTimeLeft $now
            Populate-Report -strDisplayName $objUser.DisplayName -objPwdTimeLeft $objPwdTimeLeft -now $now
        }
    }

    # If reporting enabled, send report
    if ($ghtReports.Enabled -eq $True) {
        Send-Report
    }
}


# --- Script execution starts here --- #

if($preventScriptExecution -eq $true) {
    Write-Output 'You must configure this script for your environment and ensure you set $preventScriptExecution to $False'
    Break
} 

# Import required module.
Import-Module ActiveDirectory

# Consolidate user set variables (from the top) with a few others to create a global hashtable of all settings
Set-Variable -Name ghtSettings -Value @{
    'Domain' = (Get-ADDomain);
    'Email' = @{
        'Contact' = @{
            'Address' = $emailContactAddress;
            'DisplayName' = $emailContactDisplayName;
            'Subject' = $emailContactSubject };
        'Exempt' = $emailExempt;
        'From' = @{
            'Address' = $emailFromAddress;
            'DisplayName' = $emailFromDisplayName };
        'Report' = @{
            'Address' = $emailReportAddress;
            'DisplayName' = $emailReportDisplayName;
            'Subject' = $emailReportSubject; };
        'Sender' = New-Object System.Net.Mail.MailAddress($emailFromAddress, $emailFromDisplayName);
        }
    'Pwd' = @{
        'NotificationStartInDays' = $pwdNotificationStartInDays; };
    'Script' = @{
        'Path' = (Get-Item .).FullName };
    'SmtpServer' = $smtpServer;
    'Test' = @{
        'Mode' = $testMode;
        'Address' = $testAddress;
        'DisplayName' = $testDisplayName }
} -Scope global

Set-Variable -Name ghtReports -Value @{
    'Enabled' = $reportsEnabled;
    'Expired' = @{};
    'Today' = @{};
} -Scope global

# Start the process
Get-PasswordExpiredUsers

# Cleanup
Remove-Variable -Name ghtSettings -Scope global
Remove-Variable -Name ghtReports -Scope global


# For best results, please sign this PowerShell Script
# See http://bit.ly/14LdRdm on how to accomplish this
