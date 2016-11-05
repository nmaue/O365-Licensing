# Must have Azure, AD, and MSOL Sign-on Assistant installed on computer
# Fluidly manage licensing within Active Directory Users and Computers
# Disabled users' licenses are revoked and new users get licenses according to their OU
# Moving an account to a defined OU can also revoke license
# Logfile is stored in user's AppData\Roaming\O365License.log



# Logfile
$logfile1 = ($env:APPDATA + "\O365License.log")
$logfile2 = ($env:APPDATA + "\FullAccess.log")

# License Skus
$E3Sku = "DomainName:ENTERPRISEPACK"
$K1Sku = "DomainName:DESKLESSPACK"

# License options
$UsageLocation = "US"

# Credentials
$AdminUsername = "admin@DomainName.onmicrosoft.com"
$AdminPassword = "AdminPassword"

# Start Script

ac $logfile1 "-----$(Get-Date) O365License v0.1 $($env:COMPUTERNAME) Session log-----`n"

$SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminUsername,$SecurePassword

# Load modules

$env:PSModulePath += ";C:\Windows\System32\WindowsPowerShell\v1.0\Modules\"

try{
    Import-Module msonline
}catch{
    ac $logfile1 "Critical error, unable to load Azure AD module, did you install it?"
    ac $logfile1 $error[0]
    Exit
}
try{
    Import-Module activedirectory
}catch{
    ac $logfile1 "Critical error, unable to load ActiveDirectory module, did you install it?"
    ac $logfile1 $error[0]
    Exit
}


try{
    Connect-MsolService -Credential $cred -ErrorAction Stop
}catch{
    ac $logfile1 "Critical error, unable to connect to O365, check the credentials"
    ac $logfile1 $error[0]
    Exit
}

#store available licenses
$accountlicenses = Get-MsolAccountSku

$e3licenses = $accountlicenses | where-object {$_.AccountSkuId -like "*$($e3Sku)*"}
$k1licenses = $accountlicenses | where-object {$_.AccountSkuId -like "*$($k1Sku)*"}

ac $logfile1 "Current licensing status:"
ac $logfile1 "K1 Licences used: $($k1licenses.ConsumedUnits) / $($k1licenses.ActiveUnits)"
ac $logfile1 "E3 Licences used: $($e3licenses.ConsumedUnits) / $($e3licenses.ActiveUnits)"


# Get OU Containers
"OU=TheOUYouWant, DC=DOMAIN, DC=NET","OU=AnotherOUYouWant, DC=DOMAIN, DC=NET" | ForEach {
    Get-ADUser -Filter * -SearchBase $_  |

    # Assign Licenses
    ForEach-Object {
        # Check if user is enabled
        If ($_.enabled -eq $true) {
            # Assign E3 License
            try {
                $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
            } catch {
                ac $logfile1 "Unable to find O365 User $($_.UserPrincipalName)"
                ac $logfile1 $error[0]
            }

            If( -Not $user.isLicensed) {
                try {
                    Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $UsageLocation
                } catch {
                    ac $logfile1 "Unable to assign usage location to $($_.UserPrincipalName)"
                    ac $logfile1 $error[0]
                }
                
                try {
                    Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $E3Sku
                    ac $logfile1 "E3 license assigned to $($_.UserPrincipalName)"
                } catch {
                    ac $logfile1 "Unable to assign E3 license to $($_.UserPrincipalName)"
                    ac $logfile1 $error[0]
                }
            }
        }
    }
}

# Get OU For K1
"OU=TheOUYouWant, DC=DOMAIN, DC=NET","OU=AnotherOUYouWant, DC=DOMAIN, DC=NET" | ForEach {
    Get-ADUser -Filter * -SearchBase $_  |

    # Assign Licenses
    ForEach-Object {
        # Check if user is enabled
        If ($_.enabled -eq $true) {
            # Assign K1 License
            try {
                $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
            } catch {
                ac $logfile1 "Unable to find O365 User $($_.UserPrincipalName)"
                ac $logfile1 $error[0]
            }

            If( -Not $user.isLicensed) {
                try {
                    Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $UsageLocation
                } catch {
                    ac $logfile1 "Unable to assign usage location to $($_.UserPrincipalName)"
                    ac $logfile1 $error[0]
                }
                
                try {
                    Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $K1Sku
                    ac $logfile1 "K1 license assigned to $($_.UserPrincipalName)"
                } catch {
                    ac $logfile1 "Unable to assign K1 license to $($_.UserPrincipalName)"
                    ac $logfile1 $error[0]
                }
            }
        }
        Elseif ($_.enabled -eq $false) {
            
            # Check if disabled user has a license
            try {
                $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
            } catch {
                ac $logfile1 "Unable to find O365 User $($_.UserPrincipalName)"
                ac $logfile1 $error[0]
            }

            If( $user.isLicensed) {
                # Remove all licenses
                $license = Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select-Object licenses
                $licensearray = $license.Licenses
                For ($i=0; $i -lt $licensearray.count; $i++) {
                    try {
                        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $licensearray[$i].AccountSkuId
                        ac $logfile1 "License $($licensearray[$i].AccountSkuId) removed from $($user.UserPrincipalName)"
                    } catch {
                        ac $logfile1 "Unable to remove license $($licensearray[$i].AccountSkuId) from $($user.UserPrincipalName)"
                        ac $logfile1 $error[0]
                    }
                }
            }
        }
    }
}

# Remove Licenses for Old Accounts
# Get OU Containers Old Accounts
"OU=TheOUYouWant, DC=DOMAIN, DC=NET","OU=AnotherOUYouWant, DC=DOMAIN, DC=NET" | ForEach {
    Get-ADUser -Filter * -SearchBase $_ |

    # Remove Licenses
    ForEach-Object {
            
        try {
            $user = Get-MsolUser -UserPrincipalName $_.UserPrincipalName
        } catch {
            ac $logfile1 "Unable to find O365 User $($_.UserPrincipalName)"
            ac $logfile1 $error[0]
        }
            
        $license = Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select-Object licenses
        $licensearray = $license.Licenses
        For ($i=0; $i -lt $licensearray.count; $i++) {
            try {
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $licensearray[$i].AccountSkuId
                ac $logfile1 "License $($licensearray[$i].AccountSkuId) removed from $($user.UserPrincipalName)"
            } catch {
                ac $logfile1 "Unable to remove license $($licensearray[$i].AccountSkuId) from $($user.UserPrincipalName)"
                ac $logfile1 $error[0]
            }
        }  
    }
}

# end licensing script
ac $logfile1 "-----$(Get-Date) O365License v0.1 $($env:COMPUTERNAME) Session END-----`n"
ac $logfile1 "`n"

# start full access
# import microsoft exchange

ac $logfile2 "-----$(Get-Date) FullAccess v0.1 $($env:COMPUTERNAME) Session log-----`n"

try {
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
} catch {
    ac $logfile2 "Unable to connect to Exchange Online, check credentials"
    ac $logfile2 $error[0]
}

try {
    Import-PSSession $session
} catch {
    ac $logfile2 "Unable to import Exchange Online commands"
    ac $logfile2 $error[0]
}


$mailboxes = Get-Mailbox -resultsize unlimited

$mailboxes | ForEach-Object {
    $permissions = Get-MailboxPermission -Identity $_.PrimarySMTPAddress -User admin@DomainName.onmicrosoft.com
    If (-not $permissions.AccessRights -contains "FullAccess") {
        try {
            Add-MailboxPermission -Identity $_.PrimarySMTPAddress -User admin@DomainName.onmicrosoft.com -AccessRights FullAccess -AutoMapping:$false
            ac $logfile2 "Added Full Access permission to $($_.PrimarySMTPAddress)"
        } catch {
            ac $logfile2 "Unable to add Full Access permission to $($_.PrimarySMTPAddress)"
            ac $logfile2 $error[0]
        }
    }
}

Remove-PSSession $session
# end full access

