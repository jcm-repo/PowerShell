<#
    .AUTHOR
        https://github.com/jcm-repo/PowerShell
    
    .DESCRIPTION
        This script will send email notifications to users reminding them to change their passwords before it expires.

    .USAGE
        Run, or schedule, the report from any domain member on which RSAT is installed.
    
    .PARAMETERS
        -MailFrom
            Specify the email address from which the email notifications will be sent. It is recommended to use your
            helpdesk's email address, but this should only be done if your helpdesk system can be configured to ignore
            read receipts and auto replies.

        -MailServer
            Specify an IP address or DNS name for your SMTP server.
#>

#region Set Initial Variables
    
    [CmdletBinding()]
    Param
    (
        [string] $MailFrom = 'helpdesk@company.com', # Read usage!
        
        [string] $MailServer = 'smtp.company.com'
    )

    $DateFormat = 'dddd, dd MMMM, "at" HH:mm'
    
    $MaxPasswordAge = ( Get-ADDefaultDomainPasswordPolicy ).MaxPasswordAge.Days
    
    $NotificationDays = 1, 2, 3, 5, 10, 15

#endregion

#region Set Initial Functions
    
    Function Generate-MailMessage
    {
        "Dear $( $User.GivenName ),`n`n" +
        "Please note your password will expire in $( $User.DaysRemaining ) day(s) on $( $User.ExpiryDateTime ).`n`n" +
        "To change your password now, press the `"CTRL`" + `"ALT`" + `"DELETE`" keys on your keyboard and select the `"Change a password`" option.`n`n" +
        "Please note this email is an automated message for information purposes only - there is no need to reply to this message.`n`n" +
        "Kind regards,`n" +
        "Helpdesk`n" +
        "$( $MailFrom )"
    }

    Function Get-DaysRemaining
    {
        ( New-TimeSpan -Start ( Get-Date ) -End $_.PasswordLastSet.AddDays( $MaxPasswordAge ) ).Days
    }

#endregion

#region Get User Inforamtion From Active Directory
    
    $ADUsersParams = @{
        LDAPFilter = “(!LogonHours=$( '\00' * 21 ))”
        Properties = `
            'GivenName',`
            'Mail',`
            'PasswordLastSet',`
            'PasswordNeverExpires' }

    $ADUsers = Get-ADUser @ADUsersParams |
        Where-Object -FilterScript {
            $_.Enabled -and
            $_.Mail -and
            $_.PasswordLastSet -and
            !$_.PasswordNeverExpires
        } |
        Select-Object -Property @(
            'GivenName'
            'Mail'
            @{ Name       = 'ExpiryDateTime'
               Expression = { $_.PasswordLastSet.AddDays( $MaxPasswordAge ).ToString( $DateFormat ) }
            }
            @{ Name       = 'DaysRemaining'
               Expression = { Get-DaysRemaining }
            }
        )

#endregion

#region Send Email Notifications
    
    ForEach ( $User in $ADUsers )
    {
        If ( $NotificationDays -contains $User.DaysRemaining )
        {
            # Don't send email to users during testing (PowerShell ISE detected)
            If ( $psISE )
            {
                Write-Output "`n=============================================`n"
                Write-Output "Mail To: $( $User.Mail )`n"

                Generate-MailMessage
            }
            Else
            {
                $MailParams = @{
                    From       = $MailFrom
                    To         = $User.Mail
                    Subject    = "Password Expiring Within $( $User.DaysRemaining ) days"
                    Body       = ( Generate-MailMessage )
                    SmtpServer = $MailServer
                }
                
                Send-MailMessage @MailServer
            }
        }
    }

#endregion
