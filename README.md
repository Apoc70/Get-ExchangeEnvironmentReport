# Get-ExchangeEnvironmentReport

Creates an HTML report describing the On-Premises Exchange environment.

Based on the original 1.6.2 version by Steve Goodman

## Description

This script creates an HTML report showing the following information about an Exchange 2019, 2016, 2013, 2010, and, to a lesser extent, 2007 and 2003 environment.

The report shows the following:

* As summary
  * Total number of servers per Exchange Server version
  * Total number of mailboxes per On-Premises Exchange Server version, Office 365, and Exchange Organisation
  * Total number of Exchange Server functional roles

* Per Active Directory Site
  * Total number of mailboxes
  * Internal, External, and CAS Array names
  * Exchange Server computers
    * Product version
    * Service Pack, Update Rollup, and/or Cumulative Update
    * Number of preferred and maximum active databases
    * Functional Roles
    * Operating System with Service Pack

* Per Database Availability Group
  * Total number of member servers
  * List of member servers
  * DAG databases
    * Number of mailboxes and average mailbox size
    * Number of archive mailboxes and average archive mailbox size
    * Database size
    * Database whitespace
    * Disk space available for database and log file volume
    * Last full backup timestamp
    * Circular logging enabled
    * Mailbox server hosting an active copy
    * List of mailbox servers hosting database copies

* Per Database (Non-DAG, pre-DAG Exchange Server)
  * Storage group and database name
  * Server name hosting the database
  * Number of mailboxes and average mailbox size
  * Number of archive mailboxes and average archive mailbox size
  * Database size
  * Database whitespace
  * Disk space available for database and log file volume
  * Last full backup timestamp
  * Circular logging enabled

The PowerShell script does not gather information on public folders or analyzes Exchange cluster technologies like Exchange Server 2007/2003 CCR/SCR.

## Requirements

* Exchange Server Management Shell 2010 or newer
* WMI and Remote Registry access from the computer running the script to all internal Exchange Servers

## Release

* 2.0 : Initial Community Release of the updated original script
* 2.1 : Table header label updated for a more consistent labeling
* 2.2 : Bug fixes and enhancements
  * CCS fixes for Html header tags (issue #5)
  * New script parameter _ShowDriveNames_ added to optionally show drive names for EDB/LOG file paths in database table (issue #4)
  * Exchange organization name added to report header  

## Example Report

![Example Report](/images/screenshot.png)

## Parameters

### HTMLReport

File name to write HTML Report to

### SendMail

Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory

### MailFrom

Email address to send from. Passed directly to Send-MailMessage as -From

### MailTo

Email address to send to. Passed directly to Send-MailMessage as -To

### MailServer

SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer

### ViewEntireForest

By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.

### ServerFilter

Use a text based string to filter Exchange Servers by, e.g., NL-*
Note the use of the wildcard (*) character to allow for multiple matches.

### ShowDriveNames

Include drive names of EDB file path and LOG file folder in database report table

## Examples

### Example 1

Generate an HTML report and send the result as HTML email with attachment to the specified recipient using a dedicated smart host

``` PowerShell
.\Get-ExchangeEnvironmentReport.ps1 -HTMReport ExchangeEnvironment.html -SendMail -ViewEntireForet $true -MailFrom roaster@mcsmemail.de -MailTo grillmaster@mcsmemail.de -MailServer relay.mcsmemail.de
```

### Example 2

Generate an HTML report and save the report as 'report.html'

``` PowerShell
.\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html
```

### Example 3

Generate the HTML report including EDB and LOG drive names

``` PowerShell
.\Get-ExchangeEnvironmentReport.ps1 -ShowDriveNames -HTMLReport .\report.html
```

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Based on the original 1.6.2 version by Steve Goodman.

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)