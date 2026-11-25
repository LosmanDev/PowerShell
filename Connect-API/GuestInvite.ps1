<#
########################################################
Bulk guest invite from CSV (set attributes, manager, group)
########################################################
#>

Connect-MgGraph -Scopes "User.Invite.All","User.ReadWrite.All","GroupMember.ReadWrite.All"

$csvPath = "$HOME/Documents/Guest Accounts/Guest_Account_List.csv"
$groupId = "00000000-0000-0000-0000-000000000000"   # Replace with target group's ID
$logPath = "$HOME/Documents/Guest Accounts/guest-invite-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    $line | Tee-Object -FilePath $logPath -Append
}

function Get-TrimmedValue {
    param($Value)
    return ([string]$Value).Trim()
}

Write-Log "Starting from $csvPath"

Import-Csv $csvPath | ForEach-Object {

    $firstName   = Get-TrimmedValue $_.'First Name'
    $lastName    = Get-TrimmedValue $_.'Last Name'
    $email       = Get-TrimmedValue $_.'Email'
    $employeeId  = Get-TrimmedValue $_.'Employee ID'
    $title       = Get-TrimmedValue $_.'Job Title'
    $state       = Get-TrimmedValue $_.'State'
    $country     = Get-TrimmedValue $_.'Country'
    $fteManager  = Get-TrimmedValue $_.'FTE Manager'
    $companyName = Get-TrimmedValue $_.'Company name'

    if ([System.String]::IsNullOrWhiteSpace($email)) {
        Write-Log "Skipping empty email" "WARNING"
        return
    }

    $displayName = ("$firstName $lastName").Trim()
    if ([System.String]::IsNullOrWhiteSpace($displayName)) { $displayName = $email }

    Write-Log "Processing $displayName <$email>"

    try {
        $guestUser = Get-MgUser -Filter "mail eq '$email'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue

        if (-not $guestUser) {
            $inv = New-MgInvitation `
                -InvitedUserDisplayName $displayName `
                -InvitedUserEmailAddress $email `
                -InviteRedirectUrl "https://myapps.microsoft.com" `
                -SendInvitationMessage:$true `
                -ErrorAction Stop

            Write-Log "Invitation: $($inv.Id) $($inv.Status)"

            $guestUser = $null
            for ($i = 1; $i -le 12; $i++) {
                Start-Sleep 10
                $guestUser = Get-MgUser -Filter "mail eq '$email'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
                if ($guestUser) { break }
            }

            if (-not $guestUser) { throw "Guest not found after retries" }
        }

        $updateParams = @{}
        if ($firstName)   { $updateParams.GivenName   = $firstName }
        if ($lastName)    { $updateParams.Surname     = $lastName }
        if ($title)       { $updateParams.JobTitle    = $title }
        if ($companyName) { $updateParams.CompanyName = $companyName }
        if ($email)       { $updateParams.OtherMails  = @($email) }

        if ($employeeId -and $employeeId.Length -le 16) {
            $updateParams.EmployeeId = $employeeId
        }

        if ($state)   { $updateParams.State   = $state }
        if ($country) { $updateParams.Country = $country }

        if ($updateParams.Count -gt 0) {
            Update-MgUser -UserId $guestUser.Id @updateParams -ErrorAction Stop
        }

        if ($fteManager) {
            $managerUser = Get-MgUser -Filter "mail eq '$fteManager' or userPrincipalName eq '$fteManager'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue

            if ($managerUser) {
                try {
                    Set-MgUserManagerByRef -UserId $guestUser.Id -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($managerUser.Id)"
                    } -ErrorAction Stop
                    Write-Log "Manager set: $fteManager"
                }
                catch {
                    if ($_.Exception.Message -match "already|Conflict") {
                        Write-Log "Manager already set" "INFO"
                    }
                }
            }
            else {
                Write-Log "Manager not found: $fteManager" "WARNING"
            }
        }

        try {
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($guestUser.Id)"
            } -ErrorAction Stop
            Write-Log "Added to group"
        }
        catch {
            if ($_.Exception.Message -match "already exist") {
                Write-Log "Already in group" "INFO"
            }
        }

        Write-Log "SUCCESS: $email" "SUCCESS"
    }
    catch {
        Write-Log "ERROR ${email}: $($_.Exception.Message)" "ERROR"
    }
}

Write-Log "Complete. Log: $logPath"

Disconnect-MgGraph

<#
########################################################
Re-invite user
########################################################
#>

Connect-MgGraph -Scopes "User.Invite.All","User.ReadWrite.All"

$email = "guest@external.com"
$safeEmail = $email.Replace("'","''")
$user = @(Get-MgUser -Filter "mail eq '$safeEmail'" -Property Id,Mail,UserType)

if ($user.Count -ne 1) {
    throw "Expected one user for $email; found $($user.Count)."
}

if ($user[0].UserType -ne "Guest") {
    throw "$email is not a guest account."
}

New-MgInvitation `
    -InvitedUserEmailAddress $email `
    -InviteRedirectUrl "https://myapps.microsoft.com" `
    -ResetRedemption `
    -SendInvitationMessage `
    -InvitedUser $user[0]

<#
########################################################
Invite guest
########################################################
#>

# Groups every new guest is added to
$MfaGroupId             = "00000000-0000-0000-0000-000000000001"   # Replace with MFA enrollment group ID
$SharePointGuestGroupId = "00000000-0000-0000-0000-000000000002"   # Replace with SharePoint guest group ID

$email     = "guest@external.com"
$firstName = "First"
$lastName  = "Last"

if ((Get-MgUser -Filter "usertype eq 'Guest' and mail eq '$email'" -All)) {
    Write-Host "$email Guest user already exists" -BackgroundColor Blue
} else {
    New-MgInvitation -InvitedUserEmailAddress $email -InvitedUserDisplayName "$firstName $lastName" -InviteRedirectUrl 'https://myapplications.microsoft.com' -SendInvitationMessage:$false
}

Start-Sleep -Seconds 20

$usr = Get-MgUser -Filter "usertype eq 'Guest' and mail eq '$email'" -All

if ($usr) {
    # Add user to the MFA enrollment group
    Write-Host "Adding $firstName $lastName to the group"
    New-MgGroupMember -GroupId $MfaGroupId -DirectoryObjectId $usr.Id

    # Add user to the SharePoint guest group
    New-MgGroupMember -GroupId $SharePointGuestGroupId -DirectoryObjectId $usr.Id

    Write-Host "Setting Name: $firstName $lastName"
    Update-MgUser -UserId $usr.Id -GivenName $firstName -Surname $lastName
}

Read-Host

<#
########################################################
CSV guest invite
########################################################
#>

Start-Transcript -Path .\logs\Log-GuestInvitation.log -Append

Connect-MgGraph -Scopes "User.Invite.All","UserAuthenticationMethod.Read.All","User.Read.All"

$template = Import-Csv .\GuestTemplate.txt

foreach ($i in $template) {
    if ((Get-MgUser -Filter "usertype eq 'Guest' and mail eq '$($i.email)'" -All)) {
        Write-Host ($i.email.ToString() + ' Guest user already exists') -BackgroundColor Blue
    } else {
        New-MgInvitation -InvitedUserEmailAddress $i.email -InvitedUserDisplayName ($i.First.ToString() + ' ' + $i.Last.ToString()) -InviteRedirectUrl 'https://myapplications.microsoft.com' -SendInvitationMessage:$false
    }
}

Start-Sleep -Seconds 20

# To figure out Strong Authentication Requirements
foreach ($i in $template) {

    $usr = Get-MgUser -Filter "usertype eq 'Guest' and mail eq '$($i.email)'" -All

    # Add user to the MFA enrollment group
    Write-Host 'Adding '$i.First' '$i.Last' to the group'
    New-MgGroupMember -GroupId $MfaGroupId -DirectoryObjectId $usr.Id

    # Add user to the SharePoint guest group
    New-MgGroupMember -GroupId $SharePointGuestGroupId -DirectoryObjectId $usr.Id

    Write-Host 'Setting Name: '$i.First' '$i.Last
    Update-MgUser -UserId $usr.Id -GivenName $i.First -Surname $i.Last
}

Stop-Transcript

Write-Host ''
Write-Host 'EOF, Push Enter.' -BackgroundColor DarkGreen
Read-Host
