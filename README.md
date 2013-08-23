Active Directory Password Expiry Notifier
=========================================

ADPEN searches through an Active Directory domain using an LDAP query to find users who's passwords will soon expire or have already expired.  It then sends an email notification to that user and (optionally) a report to an administrator.

## Usage
1. Clone [this repository](https://github.com/richardfleming/adpen.git "this repository")
2. Modify adpen.ps1 to match your environment
3. Replace .\Templates\logo.png to use your own (600 x any height)
4. Test, test and test.
5. Copy over to an Active Directory Domain Controller
6. Create a batch file that runs this script
7. Open Task Scheduler and create a task that points to your batch file.
8. Run your task and see if it works.
9. Ensure the $testMode variable is set to $False
8. You're done.

Ensure that your domain controller is set to run remote signed scripts, or this will fail.

### Configurables
Specifies the number of days remaining (before password expires) that email notification emails start

- **$pwdNotificationStartInDays** (Default = 15)  

Enable or disable the distribution of daily administrator report on who's password expires today and who's password has already expired.

- **$reportsEnabled** (Default = $True)

The fully qualified domain name of the SMTP server this script should use
   
- **$smtpServer**


The following three $emailContactX variables are specifically related to the email notification email users receive.  That email contains a contact link that users can click on to send an email requesting help.  Suggestions are your helpdesk ticketing system, or a helpdesk distribution list.

- **$emailContactAddress** (E-Mail Address)
- **$emailContactDisplayName** (E-Mail nice display name)
- **$emailContactSubject** (E-Mail Subject line)

The following three $emailReportX variables are similar to the $emailContactX variables, but are specific to whom reports go to.

- **$emailReportAddress**  
- **$emailReportDisplayName**  
- **$emailReportSubject**

The following two $emailFromX variables define what email address sends the notification and reporting emails.  Example: no-reply@example.com "Automated Password Expiry Notifier"
  
- **$emailFromAddress**  
- **$emailFromDisplayName**

If you have users who knowingly have the 'Password Never Expires' flag set on their account you can exempt them from receiving password expired emails by adding their SAM account name to this array

- **$emailExempt** = @( "BExample", "JSmith" )

There is a Test-Mode that is enabled by default to avoid flooding your users with emails while you configure this script for your site.  The three $testX variables defined below turn on/off test mode, specify which email address to send test emails to, and which user (via DisplayName) the LDAP query uses

- **$testMode** (Default = $True)
- **$testAddress**
- **$testDisplayName**

The last is a safety feature that prevents script execution.  By default it is set to true to ensure you edit this script and configure it to your specific site.  You **must** change this value for the script to execute.

- **$preventScriptExecution** (Default = $True)

## References
1. [Brice Lambson](http://brice-lambson.blogspot.ca/2012/09/simple-template-engine-for-powershell.html "Brice Lambson's"): Simple Template Engine for PowerShell.  
2. [MailChimp/Email-Blueprints](https://github.com/mailchimp/Email-Blueprints "MailChimp - Email-Blueprints"): HTML E-Mail template [base\\_boxed\\_basic\\_query.html](https://github.com/mailchimp/email-blueprints/blob/master/responsive-templates/base_boxed_basic_query.html "base_boxed_basic_query.html") modified and used as basis for .\Templates\notificationEmail.html and .\Templates\reportEmail.html  
