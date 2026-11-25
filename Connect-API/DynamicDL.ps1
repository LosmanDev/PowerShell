<#
#############################################################################
Pull DDL members
#############################################################################
#>

Connect-ExchangeOnline -Device

Get-DynamicDistributionGroupMember -Identity "ExampleDDL" | Select-Object Name, PrimarySmtpAddress | Export-Csv -Path "/Downloads/ExampleDDL_Members.csv" -NoTypeInformation

Get-DynamicDistributionGroup -Identity "ExampleDDL" | Select-Object RecipientFilter, RecipientContainer

<#
#############################################################################
Pull all DDLs' sender permissions
#############################################################################
#>

$groups = @(
    "ExampleDDL"
)

foreach ($group in $groups) {
    $recip = Get-Recipient -Identity $group -ErrorAction SilentlyContinue
    if ($recip.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        Get-DynamicDistributionGroup -Identity $group | Select-Object Name, @{Name = "AcceptMessagesOnlyFrom"; Expression = { $_.AcceptMessagesOnlyFrom -join ", " } }, @{Name = "AcceptMessagesOnlyFromSendersOrMembers"; Expression = { $_.AcceptMessagesOnlyFromSendersOrMembers -join ", " } } | Format-Table -Wrap
    }
    elseif ($recip.RecipientTypeDetails -match "MailUniversal") {
        Get-DistributionGroup -Identity $group | Select-Object Name, @{Name = "AcceptMessagesOnlyFrom"; Expression = { $_.AcceptMessagesOnlyFrom -join ", " } }, @{Name = "AcceptMessagesOnlyFromSendersOrMembers"; Expression = { $_.AcceptMessagesOnlyFromSendersOrMembers -join ", " } } | Format-Table -Wrap
    }
    else {
        Write-Warning "Object '$group' not found."
    }
}

foreach ($group in $groups) {
    $recip = Get-Recipient -Identity $group -ErrorAction SilentlyContinue

    if ($recip.RecipientTypeDetails -eq "DynamicDistributionGroup") { $g = Get-DynamicDistributionGroup -Identity $group }
    elseif ($recip.RecipientTypeDetails -match "MailUniversal")     { $g = Get-DistributionGroup -Identity $group }
    else { Write-Warning "Object '$group' not found."; continue }

    "Group: $($g.DisplayName) <$($g.PrimarySmtpAddress)>  (Name: $($g.Name))"

    $g.AcceptMessagesOnlyFromSendersOrMembers |
        ForEach-Object { Get-Recipient $_ -ErrorAction SilentlyContinue } |
        Sort-Object DisplayName |
        Format-Table DisplayName, PrimarySmtpAddress, RecipientTypeDetails -AutoSize
}

<#
#############################################################################
Add $senders
#############################################################################
#>

$senders = @(
    "sender@company.com"
)

foreach ($group in $groups) {
    $recip = Get-Recipient -Identity $group -ErrorAction SilentlyContinue
    if ($recip.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        Set-DynamicDistributionGroup -Identity $group -AcceptMessagesOnlyFromSendersOrMembers @{Add = $senders }
        Write-Host "Updated: $group"
    }
    elseif ($recip.RecipientTypeDetails -match "MailUniversal") {
        Set-DistributionGroup -Identity $group -AcceptMessagesOnlyFromSendersOrMembers @{Add = $senders }
        Write-Host "Updated: $group"
    }
    else {
        Write-Warning "Failed: '$group' not found."
    }
}

<#
#############################################################################
Check if a specific user can send to a DDL
#############################################################################
#>

$ddl      = "ExampleDDL"
$sendUser = "user@company.com"

$g = Get-DynamicDistributionGroup $ddl
$s = Get-Recipient $sendUser

function Test-InList($list) {
    foreach ($e in $list) {
        $r = Get-Recipient $e
        if ($r.DistinguishedName -eq $s.DistinguishedName) { return "Direct: $($r.DisplayName) <$($r.PrimarySmtpAddress)>" }
        if ($r.RecipientTypeDetails -match 'Group' -and
            (Get-DistributionGroupMember $r.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue | Where-Object DistinguishedName -eq $s.DistinguishedName)) {
            return "Via group: $($r.DisplayName) <$($r.PrimarySmtpAddress)>"
        }
    }
    return "Not found"
}

"Group: $($g.DisplayName) <$($g.PrimarySmtpAddress)>"
"Sender: $($s.DisplayName) <$($s.PrimarySmtpAddress)> [$($s.RecipientType)]"
"Allowed list : " + $(if ($g.AcceptMessagesOnlyFromSendersOrMembers) { Test-InList $g.AcceptMessagesOnlyFromSendersOrMembers } else { "No restriction (open)" })
"Blocked list : " + (Test-InList $g.RejectMessagesFromSendersOrMembers)
"Bypass moderation : " + (Test-InList $g.BypassModerationFromSendersOrMembers)
"Sender type  : " + $s.RecipientType
