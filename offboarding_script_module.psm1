# Offboard users
# Created by Christofer Astuto

# Download and install RSAT for the script to work with Active Directory - https://www.microsoft.com/en-us/download/details.aspx?id=45520

Function offboard ($sam) {
    
    # Checks and installs modules if not already installed
    chkModules

    # Checks if a username was input
    If ($sam){

        # Imports relevant modules
        impModules

        # Checks if the username exists in the domain
        try {
            if (Get-ADuser $sam) {
                $userExists = $true
            }
        }
        catch {
            $userExists = $false
        }

        # Runs code if the user exists
        If ($userExists -ne $false) {
            Write-Host "`r`nOffboarding the user: $sam
            "
            
            # Sets the variables
            setVars

            # Offboards Active Directory Account
            adDispension

            # Offboards the users email account
            emailDispension

            # Offboards the users home drive
            driveDispension

            # Comments on the JIRA ticket
            jiraTicket

            # Clear session variables if the user does another offboarding in the same session
            Remove-Variable * -ErrorAction SilentlyContinue

        } Else {
            "The username '$sam' does not exist."
        }
    } Else {
        "Please enter a username.`r`nFormat: offboard <username>"
    }
}

Function setVars {
    $script:date = [datetime]::Today.ToString('dd-MM-yyyy')

    # Get the properties of the account and set variables
    $script:user = Get-ADuser $sam -properties canonicalName, distinguishedName, displayName, mailNickname, name, MemberOf, mail, homeDirectory
    $script:dn = $user.distinguishedName
    $script:cn = $user.canonicalName
    $script:un = $user.SamAccountName
    $script:din = $user.displayName
    $script:UserAlias = $user.mailNickname
    $script:name = $user.name
    $script:name_nospace = $name -replace " ", "_"
    $script:email = $user.mail
    $script:homeDir = $user.homeDirectory

    # Set email variables
    $script:emailTo = '<Helpdesk Email Address>'
    $script:emailFrom = '<Automation Email Address>'
    $script:emailSMTP = '<SMTP Server>'

    # Set domain controller variable
    $script:domainController = '<Domain Controller>'

    # Set slack URI variable for automation slack channel
    $script:slackURI = '<Slack URI>'

    # Set Home drive backup location variable
    $script:homeDirBkp = '<Backup Server Root Location for Home Drives>' + $un
}

Function adDispension {
    # Offboards Active Directory Account

    # Disable the account
    Disable-ADAccount $sam -Server $domainController
    Write-Host ("* $din's Active Directory account is disabled.")

    # Reset password
    Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "<Random Password>" -Force) $sam -Server $domainController
    Write-Host ("* $din's Active Directory password has been changed.")

    # Strip the permissions from the account (Except 'Domain Users' and 'Tunnel_Users')
    foreach($group in ($user | Select-Object -ExpandProperty MemberOf)) {
        If ($group -ne '<AD Group to Retain>') {
            Remove-ADGroupMember -Identity $group -Members $sam -Confirm: $false -Server $domainController
            }
        }
    Write-Host ("* $din's Active Directory group memberships (permissions) stripped from account.")

    # Move the account to the Disabled Users OU
    Move-ADObject -Identity $dn -TargetPath '<Disabled Users OU>' -Server $domainController
    Write-Host ("* $din's Active Directory account moved to 'Disabled Users' OU.")
}

Function emailDispension {
    # Offboards the users email account

    # Import the Exchange snapin (assumes desktop PowerShell)
    if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.SnapIn"})) { 
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI http://'<Exchange Server>'/PowerShell/ -Authentication kerberos
        import-PSSession $session -DisableNameChecking -AllowClobber -WarningAction:SilentlyContinue | Out-Null
    }

    # Set export date as a variable
    $exportDate = [datetime]::Today.adddays(90).ToString('dd-MM-yyyy')

    # Hides the email from address lists
    Set-Mailbox -Identity $UserAlias -HiddenFromAddressListsEnabled $true
    Write-Host ("* $din's Email has been hidden from address lists.")

    # Loop flag variables
    $Go1 = 0
    $GoDone = 0

    While ($GoDone -ne 1) {
        # Ask admin what they want to do with user's email
        $EmailWhat = Read-Host "
        What would you like to do with the user's email account? (1, 2, 3, or 4)
        
        Choices: 
        1) Forward the user's emails to another user and send a reminder to export the mailbox.
        2) Set an out of office message and send a reminder to export the mailbox.
        3) Send a reminder to export the mailbox.
        4) Leave it alone or exit.
        
        "
        If ($EmailWhat -eq "1") {
            # Forward email to manager or other user
            Write-Host "
        Selection Chosen: Forward the user's email to a manager or other user and send a reminder to export the mailbox."

            $ForEm = Read-Host "
        Email account to forward emails to? (example: John.Smith@domain.com.au)
        "
            Set-Mailbox $UserAlias -ForwardingAddress $ForEm

            Write-Host ("`r`n* The user's email, $email has been forwarded to $ForEm.")

            # Sends a reminder to export the mailbox and archive
            $emailSubject = 'Export Mailbox - ' + $User.Name + ' - ' + $exportDate
            $emailBody = $User.Name + "'s mailbox to be exported on " + $exportDate
            Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSMTP
            
            Write-Host ("* Export Mailbox ticket for $din has been created for HelpDesk.")

            #Set JIRA Ticket variables
            $script:jiraEmail = "Emails have been forwarded to $ForEm"
            $script:jiraEmailExport = "Export mailbox ticket has been created for action on $exportDate"

            $GoDone = 1
            
        } ElseIf ($EmailWhat -eq "2") {
            
            # Set an out of office message
            Write-Host "
        Selection Chosen: Set an out of office message for the user and send a reminder to export the mailbox."
            While ($Go1 -ne "1") {
                # Ask the admin how they want to set the out of office message
                $AutoReplyOption = Read-Host "
        What would you like to do with the out of office message? (1 or 2)
        
        Choices: 
        1) Set the default out of office message and specify an email address to reference (example: Jane Do is no longer with the company please contact John.Smith@domain.com.au)
        2) Specify the out of office message
        "
                If ($AutoReplyOption -eq "1") {
                    $ForEm = Read-Host "
        Specify an Email account to reference in the out of office message (example: John.Smith@domain.com.au)
        "
                    $message = $User.Name + " is no longer with the company please contact " + $ForEm
                    $Go1 = 1
                    
                } ElseIf ($AutoReplyOption -eq "2") {
                    $message = Read-Host "What would you like the out of office message to be?
        "
                    $Go1 = 1
                    
                } Else{
                    Clear-Host
                    Write-Host "
        I'm sorry, I didn't understand. You typed '$AutoReplyOption'. Please only input '1', or '2'.
                    
        "
                }
            }
            # Sets the out of office message
            Set-MailboxAutoReplyConfiguration $UserAlias -AutoReplyState Enabled -ExternalMessage $message -InternalMessage $message
            Write-Host ("`r`n* The out of office message for $email has been set to '$message'.")

            # Sends a reminder to export the mailbox and archive
            $emailSubject = 'Export Mailbox - ' + $User.Name + ' - ' + $exportDate
            $emailBody = $User.Name + "'s mailbox to be exported on " + $exportDate
            Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSMTP
            
            Write-Host ("* Export Mailbox ticket for $din has been created for HelpDesk.")

            # Set JIRA Ticket variables
            $script:jiraEmail = "Email out of office message has been set to $message"
            $script:jiraEmailExport = "Export mailbox ticket has been created for action on $exportDate"

            $GoDone = 1

        } ElseIf ($EmailWhat -eq "3") {
            # Sends a reminder to export the mailbox and archive
            Write-Host "
        Selection Chosen: Send a reminder to export the mailbox and archive.
        "

            $emailSubject = 'Export Mailbox - ' + $User.Name + ' - ' + $exportDate
            $emailBody = $User.Name + "'s mailbox to be exported on " + $exportDate
            Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSMTP
            
            Write-Host ("* Export Mailbox ticket for $din has been created for HelpDesk.")
            
            # Set JIRA Ticket variables
            $script:jiraEmail = "Email has not been actioned (not re-directed and no out of office message set)"
            $script:jiraEmailExport = "Export mailbox ticket has been created for action on $exportDate"

            $GoDone = 1
            
        } ElseIf ($EmailWhat -eq "4") {
            # Leave Alone
            Write-Host "
        Selection Chosen: Leave email alone. Exiting.
        "

            Write-Host ("* The user's email account, $email, has NOT been actioned and has only been hidden from address lists if it exists.")
            
            # Set JIRA Ticket variables
            $script:jiraEmail = "Email has not been actioned (not re-directed and no out of office message set)"
            $script:jiraEmailExport = "No export mailbox ticket has been created"

            $GoDone = 1
        
        } Else {
            Write-Host "
        I'm sorry, I didn't understand. You typed '$EmailWhat'. Please only input '1', '2', '3', or '4'.
        "
        }
    }

    # Exits the exchange session
    Get-PSSession | Remove-PSSession
}

Function driveDispension {
    # Offboards the users home drive

    # Loop flag variable
    $GoDoneDrive = 0

    While ($GoDoneDrive -ne 1) {
        $DriveWhat = Read-Host "
        What would you like to do with the user's home drive? (1, 2, 3, or 4)
        
        Choices: 
        1) Backup and delete
        2) Backup and grant access to a manager
        3) Backup and grant access to a team
        4) Leave it alone or exit.

        "

        If ($DriveWhat -eq "1") {
            # Backup and delete
            Write-Host "
        Selection Chosen: Backup and delete the home drive.

        Please wait while the home drive is backed up (depending of the drive size, this may take some time).
        "

            robocopy $homeDir $homeDirBkp /E /MOVE /NFL /NDL /R:10 /W:5

            Write-Host ("* The user's Home Drive has been backed up to " + $script:homeDirBkp + ".")

            # Set JIRA Ticket variable
            $script:jiraDrive = "Home Drive has been backed up to " + $script:homeDirBkp

            $GoDoneDrive = 1

        } ElseIf ($DriveWhat -eq "2") {
            # Backup and grant access to a manager
            Write-Host "
        Selection Chosen: Backup home drive and grant access to a manager.
            "

            $driveMoveUser = Read-Host "
        Manager username to move the Home Drive to? (example: JSmith)
        "
            $driveMove = (Get-ADuser $driveMoveUser -properties homeDirectory).homeDirectory + "\" + $un

            Write-Host "
        Please wait while the home drive is being copied and backed up (depending of the drive size, this may take some time).
            "

            robocopy $homeDir $driveMove /E /COPYALL /NFL /NDL /R:10 /W:5
            robocopy $homeDir $homeDirBkp /E /MOVE /NFL /NDL /R:10 /W:5
            
            Write-Host ("
    * The user's Home Drive has been copied to $driveMove and also backed up to " + $script:homeDirBkp + ".")

            # Set JIRA Ticket variable
            $script:jiraDrive = "Home Drive has been copied to $driveMove and also backed up to " + $script:homeDirBkp

            $GoDoneDrive = 1

        } ElseIf ($DriveWhat -eq "3") {
            # Backup and grant access to a team
            Write-Host "
        Selection Chosen: Backup home drive and grant access to a team.
        "

            $driveMove = Read-Host "
        Team drive location to move the Home Drive to? (example: \\server\folder\site)
        "
            $driveMove = $driveMove + "\" + $un
            
            Write-Host "
        Please wait while the home drive is being copied and backed up (depending of the drive size, this may take some time).
        "

            robocopy $homeDir $driveMove /E /COPYALL /NFL /NDL /R:10 /W:5
            robocopy $homeDir $homeDirBkp /E /MOVE /NFL /NDL /R:10 /W:5
            
            Write-Host ("
    * The user's Home Drive has been copied to $driveMove and also backed up to " + $script:homeDirBkp + ".")

            # Set JIRA Ticket variable
            $script:jiraDrive = "Home Drive has been copied to $driveMove and also backed up to " + $script:homeDirBkp

            $GoDoneDrive = 1

        } ElseIf ($DriveWhat -eq "4") {
            # Leave Alone
            Write-Host "
        Selection Chosen: Leave Home Drive alone. Exiting.
        "

            Write-Host ("* The user's Home Drive has NOT been actioned and is still in place.")
            
            # Set JIRA Ticket variable
            $script:jiraDrive = "Home Drive has not been actioned"

            $GoDoneDrive = 1

        } Else {
            Write-Host "
        I'm sorry, I didn't understand. You typed '$DriveWhat'. Please only input '1', '2', '3', or '4'.
        "
        }
    }
}

Function jiraTicket {
    # Comments on the JIRA ticket

    $jiraTicket = Read-Host "
    Enter the JIRA Ticket number to comment on the ticket (Leave this blank to skip).
    "

    If ($jiraTicket -ne "") {
        $emailSubject = "RE: $jiraTicket"
        $emailBody = "Offboarding Automation Application has completed the below:
        - Disabled User Account in Active Directory
        - User Password has been reset in Active Directory
        - Moved to 'Disabled Users' OU in Active Directory
        - Removed from all groups except 'Domain Users' and '<AD Group to Retain>' in Active Directory
        - Hidden from address lists in Exchange (if email exists)
        - $jiraEmail
        - $jiraEmailExport
        - $jiraDrive
        "
        Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $emailSMTP
        
        Write-Host ("`r`n* Ticket notes sent to HelpDesk ticket $jiraTicket.")

    # Sends a slack confirmation message with the ticket number linked
    New-SlackMessageAttachment -Color $([System.Drawing.Color]::blue) ` -Title $jiraTicket ` -TitleLink https://jira.'<Company Domain>'.com.au/browse/$jiraTicket ` -Pretext "User $un has been off-boarded" ` -Fallback 'Your client is bad' | New-SlackMessage | Send-SlackMessage -Uri $slackURI | out-null
    Write-Host ("* Slack notification sent in Automation Channel.
    ")

    } else {
        Write-Host ("`r`n* Ticket notes not sent to HelpDesk.")

    # Sends a slack confirmation message
    Send-SlackMessage -Uri $slackURI -Text "User $un has been off-boarded" | out-null
    Write-Host ("* Slack notification sent in Automation Channel.
    ")
    }
}

Function chkModules {
    # Checks for NuGet Package Provider and installs if it isn't already
    If (-not (Get-PackageProvider -Name NuGet)) {
        Install-PackageProvider NuGet -Force
    }
    
    # Checks for PSSlack Module and installs if it isn't already
    If (-not (Get-Module PSSlack -ListAvailable)) {
        Install-Module PSSlack -Force
    }
}

Function impModules {
    # Imports AD Module
    Import-Module ActiveDirectory

    # Imports PSSlack Module
    Import-Module PSSlack
}

# Makes only the offboard function available to users
Export-ModuleMember -Function offboard