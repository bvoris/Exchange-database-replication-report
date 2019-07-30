#Created by Brad Voris 4/8/15
#Powershell 4.0
#Grants a user full access to a mailbox

#Get SA credentials
$LiveCred = Get-Credential
#Enter Powershell Session to Exchange server and authenticate
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<EXHANGESERVERNAME>/PowerShell/ -Credential $LiveCred -Authentication Kerberos
#import the session/module
Import-PSSession $Session
# Get variable for mailbox from user input
$results = Get-MailboxDatabaseCopyStatus *  | ConvertTo-Html
#close remote session
Remove-PSSession $Session
#HTML Heading
$htmlhead = @"
<HEAD>
<TITLE>Exchange Database Copy Report</TITLE>
<style>
table {
    border-collapse: collapse;
}

table, td, th {
    border: 1px solid black;
}
</style>
</HEAD>
"@
#HTML Body for report
$htmlbody = @"
<CENTER><Font size=5><B>Exchange Database Copy Report</B></font></BR><Font size=3>$dated<BR /><TABLE cellpadding="10"><TR bgcolor= #FEF7D6><TD>Report</TD></TR><TR bgcolor= #D9E3EA><TD>$results</TD></TR></TABLE></CENTER></font>"@

#Date for file name variable
$fileDate = get-date -uformat %Y-%m-%d
#Report output & location
ConvertTo-HTML -head $htmlhead -body $htmlbody | Out-File C:\DatabaseCopyStatus-$fileDate.html