<#
    .AUTHOR
        https://github.com/jcm-repo/PowerShell/
    
    .SYNOPSIS
        Remote Desktop license report.
    
    .DESCRIPTION
        Report on installed license keypacks and issued licenses.
    
    .USAGE
        Run, or schedule, the report on a Remote Desktop license server.
#>

#region Set Initial Variables

    $MailFrom    = 'RD Licensing <support@company.com>'
    $MailTo      = 'support@company.com'
    $MailSubject = 'Remote Desktop Licensing Report'
    $MailServer  = 'smtp.company.co.za'

    $HTMLStylePath = 'C:\Scripts\HTMLStyling.css'

    # Send email to alternate address during testing (PowerShell ISE detected)
    If ($psISE)
    {
        $MailTo      = 'admin.name@company.com'
        $MailSubject = 'Remote Desktop Licensing Report (Test)'
    }

#endregion
#region Import Cascading Style Sheet
    
    If ( Test-Path -Path $HTMLStylePath )
    {
        $HTMLStyle   = Get-Content -Path $HTMLStylePath
    }
    Else
    {
        Write-Warning -Message "The CSS file located at $( $HTMLStylePath ) was not found!"
        
        $HTMLStyle   = ''
    }

#endregion
#region Get License Information

    $LicenseKeyPacks = Get-WmiObject -Class 'Win32_TSLicenseKeyPack'

    $IssuedLicenses = Get-WmiObject -Class 'Win32_TSIssuedLicense' |
        Sort-Object -Property @(
            'sIssuedToUser'
        )

#endregion
#region Create License Version Array (Needed for lookups)
    
    $LicenseVersions = @{}
    
    $LicenseKeyPacks |
        Select-Object -Property @(
            'KeyPackId'
            'ProductVersion'
        ) |
        ForEach-Object -Process {
            $LicenseVersions[$_.KeyPackId] = $_.ProductVersion
        }

#endregion
#region Process License KeyPack Information
    
    $HTMLParamters = @{
            Fragment = $true
            PreContent = "<h1>License KeyPack(s) on $( $env:COMPUTERNAME ) ($( ( $LicenseKeyPacks | Measure ).Count )):</h1>"
        }
    
    $LicenseKeyPacksHTML = $LicenseKeyPacks |
        Select-Object -Property @(
            @{ name       = 'KeyPack ID'
               expression = { $_.KeyPackId }
            }
            @{ name       = 'Product Version'
               expression = { $_.ProductVersion }
            }
            @{ name       = 'Type and Model'
               expression = { $_.TypeAndModel }
            }
            @{ name       = 'Total'
               expression = { $_.TotalLicenses }
            }
            @{ name       = 'Issued'
               expression = { $_.IssuedLicenses }
            }
            @{ name       = 'Available'
               expression = { $_.AvailableLicenses }
            }
            @{ name       = 'Expiration Date'
               expression = { $_.ConvertToDateTime( $_.ExpirationDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
        ) |
        ConvertTo-Html @HTMLParamters

#endregion
#region Process Issuance Information
    
    $HTMLParamters = @{
            Fragment = $true
            PreContent = "<h1>Issued Licenses ($( ( $IssuedLicenses | Measure ).Count )):</h1>"
        }

    $IssuedLicensesHTML = $IssuedLicenses |
        Select-Object -Property @(
            @{ name       = 'Issued To'
               expression = { $_.sIssuedToUser.Split( '\' )[1].ToUpper() }
            }
            @{ name       = 'Product Version'
               expression = { $LicenseVersions[$_.KeyPackID] }
            }
            @{ name       = 'Issue Date'
               expression = { $_.ConvertToDateTime( $_.IssueDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
            @{ name       = 'Expiration Date'
               expression = { $_.ConvertToDateTime( $_.ExpirationDate ).ToString( 'dd MMM yyyy @ HH:mm' ) }
            }
        ) |
        ConvertTo-Html @HTMLParamters

#endregion
#region Prepare Email Body
    
    $MailBody = $HTMLStyle + $LicenseKeyPacksHTML + $IssuedLicensesHTML |
        Out-String

#endregion
#region Send Email
    
    $SendMailParameters = @{
        From       = $MailFrom
        To         = $MailTo
        Subject    = $MailSubject
        Body       = $MailBody
        BodyAsHtml = $true
        SmtpServer = $MailServer
    }

    Send-MailMessage @SendMailParameters

#endregion
