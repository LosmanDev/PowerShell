<#
#############################################################################
Bulk remove users listed in a CSV from a distribution / mail-enabled security group
#############################################################################
#>

$GroupId     = "00000000-0000-0000-0000-000000000000"   # Replace with target group's ID
$InputCsv    = "$HOME/Downloads/Users_To_Remove.csv"
$EmailColumn = "Email"

# Primary and alternate email domains (used to match users across both domains)
$PrimaryDomain   = "company.com"
$AlternateDomain = "company-alt.com"

$UseBypassSecurityGroupManagerCheck = $true
$PreviewOnly = $false # Set to $true to preview without removing anyone.
$LookupBatchSize = 25
$WriteChunkSize = 250

if (-not (Test-Path -LiteralPath $InputCsv -PathType Leaf)) {
    throw "Input CSV was not found: $InputCsv"
}

$Users = @(Import-Csv -LiteralPath $InputCsv -ErrorAction Stop)
if ($Users.Count -eq 0) {
    throw "The input CSV contains no data rows."
}

$EmailProperty = $Users[0].PSObject.Properties.Name |
    Where-Object { $_ -ieq $EmailColumn } |
    Select-Object -First 1

if (-not $EmailProperty) {
    throw "The CSV does not contain the column '$EmailColumn'."
}

$Comparer = [System.StringComparer]::OrdinalIgnoreCase

function Get-MemberIdentity {
    param([object]$Member)

    $Dn = ([string]$Member.DistinguishedName).Trim()
    $GuidStr = ([string]$Member.Guid).Trim()

    if ($GuidStr -and $GuidStr -ne "00000000-0000-0000-0000-000000000000") {
        return $GuidStr
    }
    if ($Dn) { return $Dn }

    $Primary = ([string]$Member.PrimarySmtpAddress).Trim()
    if ($Primary) { return $Primary }

    return $null
}

Write-Host "Retrieving existing group members..." -ForegroundColor Cyan
$CurrentMembers = @(
    Get-DistributionGroupMember `
        -Identity $GroupId `
        -ResultSize Unlimited `
        -ErrorAction Stop
)

Write-Host "Existing members retrieved: $($CurrentMembers.Count)" -ForegroundColor Cyan

if ($CurrentMembers.Count -eq 0) {
    Write-Host "Group already has 0 members. Exiting." -ForegroundColor Yellow
    return
}

$ExistingKeys = [System.Collections.Generic.HashSet[string]]::new($Comparer)
$ExistingDnToKey = [System.Collections.Generic.Dictionary[string, string]]::new($Comparer)
$MembersByPrimary = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new($Comparer)

foreach ($Member in $CurrentMembers) {
    $Key = Get-MemberIdentity -Member $Member
    if (-not $Key) {
        throw "Encountered an existing group member without a valid Guid, DistinguishedName, or PrimarySmtpAddress. Aborting to prevent accidental membership loss."
    }

    [void]$ExistingKeys.Add($Key)

    $Dn = ([string]$Member.DistinguishedName).Trim()
    if ($Dn) {
        $ExistingDnToKey[$Dn] = $Key
    }

    $Primary = ([string]$Member.PrimarySmtpAddress).Trim()
    if ($Primary) {
        $List = $null
        if (-not $MembersByPrimary.TryGetValue($Primary, [ref]$List)) {
            $List = [System.Collections.Generic.List[string]]::new()
            $MembersByPrimary[$Primary] = $List
        }
        if (-not $List.Contains($Key)) {
            $List.Add($Key)
        }
    }
}

$Counts = [ordered]@{
    "Removed"                           = 0
    "Would Remove (Preview)"            = 0
    "Blank Email"                       = 0
    "Duplicate in CSV"                  = 0
    "Duplicate Target (Alias)"          = 0
    "No Match in Group"                 = 0
    "Ambiguous Group Members (Skipped)" = 0
    "Failed"                            = 0
}

$SeenInput = [System.Collections.Generic.HashSet[string]]::new($Comparer)
$RemoveKeySet = [System.Collections.Generic.HashSet[string]]::new($Comparer)
$RemoveQueue = [System.Collections.Generic.List[string]]::new()
$UnmatchedEmails = [System.Collections.Generic.List[string]]::new()

# Pass 1: Match directly against existing members' PrimarySmtpAddress (+ primary/alternate domain swap)
foreach ($User in $Users) {
    $Email = (([string]$User.$EmailProperty) -replace '[\u00A0\u200B\uFEFF]', '').Trim()

    if ([string]::IsNullOrWhiteSpace($Email)) {
        $Counts["Blank Email"]++
        continue
    }

    if (-not $SeenInput.Add($Email)) {
        $Counts["Duplicate in CSV"]++
        continue
    }

    $CandidateAddrs = [System.Collections.Generic.List[string]]::new()
    $CandidateAddrs.Add($Email)

    if ($Email -match '^([^@]+)@') {
        $Prefix = $Matches[1]
        $CandidateAddrs.Add("$Prefix@$PrimaryDomain")
        $CandidateAddrs.Add("$Prefix@$AlternateDomain")
    }

    $MatchedList = $null
    foreach ($Addr in $CandidateAddrs) {
        $List = $null
        if ($MembersByPrimary.TryGetValue($Addr, [ref]$List)) {
            $MatchedList = $List
            break
        }
    }

    if ($null -eq $MatchedList) {
        $UnmatchedEmails.Add($Email)
    }
    elseif ($MatchedList.Count -gt 1) {
        $Counts["Ambiguous Group Members (Skipped)"]++
    }
    else {
        $TargetKey = $MatchedList[0]
        if ($RemoveKeySet.Add($TargetKey)) {
            $RemoveQueue.Add($TargetKey)
        }
        else {
            $Counts["Duplicate Target (Alias)"]++
        }
    }
}

# Pass 2: Batched alias resolution only for CSV emails not matched by primary SMTP
if ($UnmatchedEmails.Count -gt 0) {
    $AliasToKeys = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]::new($Comparer)

    function Add-MatchedMemberAlias {
        param([object]$Recipient)

        $ResolvedKey = $null
        $RecGuid = ([string]$Recipient.Guid).Trim()
        $RecDn = ([string]$Recipient.DistinguishedName).Trim()

        if ($RecGuid -and $ExistingKeys.Contains($RecGuid)) {
            $ResolvedKey = $RecGuid
        }
        elseif ($RecDn -and $ExistingDnToKey.ContainsKey($RecDn)) {
            $ResolvedKey = $ExistingDnToKey[$RecDn]
        }

        if (-not $ResolvedKey) { return }

        $LocalAddrs = [System.Collections.Generic.HashSet[string]]::new($Comparer)
        if ($Recipient.PrimarySmtpAddress) {
            [void]$LocalAddrs.Add(([string]$Recipient.PrimarySmtpAddress).Trim())
        }
        foreach ($RawAddr in $Recipient.EmailAddresses) {
            if ($RawAddr -match '^(?i)smtp:(.+)$') {
                [void]$LocalAddrs.Add($Matches[1].Trim())
            }
        }

        foreach ($Addr in $LocalAddrs) {
            $KeySet = $null
            if (-not $AliasToKeys.TryGetValue($Addr, [ref]$KeySet)) {
                $KeySet = [System.Collections.Generic.HashSet[string]]::new($Comparer)
                $AliasToKeys[$Addr] = $KeySet
            }
            [void]$KeySet.Add($ResolvedKey)
        }
    }

    $TotalBatches = [math]::Ceiling($UnmatchedEmails.Count / $LookupBatchSize)
    Write-Host "Checking $($UnmatchedEmails.Count) unmatched addresses for secondary aliases in $TotalBatches batches..." -ForegroundColor Cyan

    for ($i = 0; $i -lt $UnmatchedEmails.Count; $i += $LookupBatchSize) {
        $BatchNum = [int]($i / $LookupBatchSize) + 1
        $EndIndex = [math]::Min($i + $LookupBatchSize - 1, $UnmatchedEmails.Count - 1)
        $Chunk = $UnmatchedEmails[$i..$EndIndex]

        $FilterClauses = foreach ($Addr in $Chunk) {
            $Escaped = $Addr -replace "'", "''"
            "EmailAddresses -eq 'smtp:$Escaped'"
        }
        $Filter = $FilterClauses -join ' -or '

        try {
            Get-EXORecipient `
                -Filter $Filter `
                -ResultSize Unlimited `
                -Properties PrimarySmtpAddress, EmailAddresses, Guid, DistinguishedName `
                -ErrorAction Stop |
            ForEach-Object { Add-MatchedMemberAlias -Recipient $_ }
        }
        catch {
            foreach ($Addr in $Chunk) {
                try {
                    $Escaped = $Addr -replace "'", "''"
                    Get-EXORecipient `
                        -Filter "EmailAddresses -eq 'smtp:$Escaped'" `
                        -ResultSize Unlimited `
                        -Properties PrimarySmtpAddress, EmailAddresses, Guid, DistinguishedName `
                        -ErrorAction Stop |
                    ForEach-Object { Add-MatchedMemberAlias -Recipient $_ }
                }
                catch { }
            }
        }

        if ($BatchNum % 10 -eq 0 -or $BatchNum -eq $TotalBatches) {
            Write-Host "  Batch $BatchNum / $TotalBatches complete ($($EndIndex + 1) / $($UnmatchedEmails.Count) checked)" -ForegroundColor DarkCyan
        }
    }

    foreach ($Email in $UnmatchedEmails) {
        $KeySet = $null
        if (-not $AliasToKeys.TryGetValue($Email, [ref]$KeySet) -or $KeySet.Count -eq 0) {
            $Counts["No Match in Group"]++
        }
        elseif ($KeySet.Count -gt 1) {
            $Counts["Ambiguous Group Members (Skipped)"]++
        }
        else {
            $TargetKey = [System.Linq.Enumerable]::First($KeySet)
            if ($RemoveKeySet.Add($TargetKey)) {
                $RemoveQueue.Add($TargetKey)
            }
            else {
                $Counts["Duplicate Target (Alias)"]++
            }
        }
    }
}

# Pass 3: Apply chunked bulk removal
if ($RemoveQueue.Count -eq 0) {
    Write-Host "No matching group members to remove." -ForegroundColor Yellow
}
elseif ($PreviewOnly) {
    $Counts["Would Remove (Preview)"] = $RemoveQueue.Count
    Write-Host "PREVIEW ONLY: Would remove $($RemoveQueue.Count) members (remaining: $($ExistingKeys.Count - $RemoveQueue.Count))." -ForegroundColor Yellow
}
elseif ($ExistingKeys.Count -gt 10000) {
    Write-Host "Group exceeds 10,000 members; removing $($RemoveQueue.Count) members individually..." -ForegroundColor Yellow
    foreach ($Key in $RemoveQueue) {
        try {
            Remove-DistributionGroupMember `
                -Identity $GroupId `
                -Member $Key `
                -BypassSecurityGroupManagerCheck:$UseBypassSecurityGroupManagerCheck `
                -Confirm:$false `
                -ErrorAction Stop
            $Counts["Removed"]++
        }
        catch {
            $Counts["Failed"]++
        }
    }
}
else {
    $CommittedKeys = [System.Collections.Generic.HashSet[string]]::new($ExistingKeys, $Comparer)
    $TotalWaves = [math]::Ceiling($RemoveQueue.Count / $WriteChunkSize)
    Write-Host "Removing $($RemoveQueue.Count) members in $TotalWaves waves of $WriteChunkSize..." -ForegroundColor Cyan

    for ($i = 0; $i -lt $RemoveQueue.Count; $i += $WriteChunkSize) {
        $WaveNum = [int]($i / $WriteChunkSize $RemoveQueue.Count - $i)
        $WaveKeys = $RemoveQueue.GetRange($i, $Count)

        $CandidateKeys = [System.Collections.Generic.HashSet[string]]::new($CommittedKeys, $Comparer)
        $CandidateKeys.ExceptWith($WaveKeys)

        try {
            $MembersPayload = if ($CandidateKeys.Count -gt 0) { [string[]]$CandidateKeys } else { @() }

            Update-DistributionGroupMember `
                -Identity $GroupId `
                -Members $MembersPayload `
                -BypassSecurityGroupManagerCheck:$UseBypassSecurityGroupManagerCheck `
                -Confirm:$false `
                -ErrorAction Stop

            $CommittedKeys.ExceptWith($WaveKeys)
            $Counts["Removed"] += $WaveKeys.Count
            Write-Host "  Wave $WaveNum / $TotalWaves committed (Remaining members: $($CommittedKeys.Count))" -ForegroundColor Green
        }
        catch {
            $WaveErr = if ($_.ErrorDetails.Message) { [string]$_.ErrorDetails.Message } else { [string]$_.Exception.Message }
            Write-Warning "Wave $WaveNum bulk write failed ($WaveErr). Resyncing state and falling back to individual removal..."

            $CommittedKeys.Clear()
            Get-DistributionGroupMember -Identity $GroupId -ResultSize Unlimited -ErrorAction Stop |
            ForEach-Object {
                $K = Get-MemberIdentity -Member $_
                if ($K) { [void]$CommittedKeys.Add($K) }
            }

            foreach ($Key in $WaveKeys) {
                if (-not $CommittedKeys.Contains($Key)) {
                    $Counts["Removed"]++
                    continue
                }

                try {
                    Remove-DistributionGroupMember `
                        -Identity $GroupId `
                        -Member $Key `
                        -BypassSecurityGroupManagerCheck:$UseBypassSecurityGroupManagerCheck `
                        -Confirm:$false `
                        -ErrorAction Stop

                    [void]$CommittedKeys.Remove($Key)
                    $Counts["Removed"]++
                }
                catch {
                    $ItemErr = if ($_.ErrorDetails.Message) { [string]$_.ErrorDetails.Message } else { [string]$_.Exception.Message }
                    if ($ItemErr -match '(?i)not a member|couldn.?t be found|could not be found') {
                        [void]$CommittedKeys.Remove($Key)
                        $Counts["Removed"]++
                    }
                    else {
                        $Counts["Failed"]++
                        Write-Warning "  Failed to remove ${Key}: $ItemErr"
                    }
                }
            }
            Write-Host "  Wave $WaveNum / $TotalWaves fallback complete (Remaining members: $($CommittedKeys.Count))" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "PROCESS COMPLETED | Preview: $PreviewOnly" -ForegroundColor Cyan
Write-Host "Total input rows: $($Users.Count)"

$Counts.GetEnumerator() |
    Where-Object { $_.Value -gt 0 } |
    Select-Object @{Name = "Status"; Expression = { $_.Key } }, @{Name = "Count"; Expression = { $_.Value } } |
    Format-Table -AutoSize
