# offboarding-application
Powershell Script Module used to offboard users from an environment that utilises Active Directory, Exchange, and JIRA

## Usage
Used to offboard users.

It will complete the following:

Active Directory:
- Disable the account
- Reset the password
- Strip permissions from the account
- Move to disabled users OU

Exchange:
- Hide the email address from the address lists
- Forward the users emails
- Set an out of office message
- Send a reminder to export the mailbox

Home Drive:
- Move the home drive to the backup server
- Move the home drive to a manager/team folder

JIRA:
- Comment on the original offboarding ticket with steps completed in the program

Slack:
- Send a notification in the automation slack channel linking the JIRA ticket number

## Dependencies
Download and install [Remote Server Administration Tools for Windows 10 (RSAT)](https://www.microsoft.com/en-us/download/details.aspx?id=45520)
This is required for the script to work with Active Directory

## Installing the PowerShell Module
Copy the module to ```C:\Program Files\WindowsPowerShell\Modules\offboard_script```

## Using the PowerShell Module
1. Open PowerShell as Administrator
2. Enter ```offboard``` followed by a username to offboard
3. Information will be displayed as the user is offboarded
