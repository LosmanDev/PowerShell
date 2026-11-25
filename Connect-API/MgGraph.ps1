<#
#############################################################################
Extract all security groups a user is in
#############################################################################
#>

Connect-MgGraph -Scopes "User.ReadWrite.All"

$UserPrincipalName = "user@company.com"

Get-MgUserMemberOf -UserId $UserPrincipalName -All |
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' } |
    Select-Object @{Name="DisplayName"; Expression={$_.AdditionalProperties["displayName"]}},
                  Id |
    Sort-Object DisplayName |
    Format-Table -AutoSize

Disconnect-MgGraph

<#
#############################################################################
Pull users from CSV and check whether enabled/disabled/not found
#############################################################################
#>

Connect-MgGraph -Scopes "User.ReadWrite.All"

# Define paths
$importPath = "$HOME/Downloads/users.csv"
$exportPath = "$HOME/Downloads/Users_Updated.csv"

$csvData = Import-Csv -Path $importPath
$userLookup = @{}

# Chunk size reduced to 7 to prevent exceeding Graph's 15 OR clause limit
$chunkSize = 7
$totalRows = $csvData.Count

for ($i = 0; $i -lt $totalRows; $i += $chunkSize) {
    $chunkLimit = [math]::Min($i + $chunkSize - 1, $totalRows - 1)
    $chunk = $csvData[$i..$chunkLimit]

    # Construct OData filter checking both UPN and Mail (Max 14 clauses)
    $filterArray = foreach ($row in $chunk) {
        $email = $row.Email
        "userPrincipalName eq '$email' or mail eq '$email'"
    }
    $filterString = $filterArray -join " or "

    # Fetch chunk from Graph
    $users = Get-MgUser -Filter $filterString -Property Id, UserPrincipalName, Mail, AccountEnabled

    # Populate local hash table
    foreach ($user in $users) {
        if (![string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
            $userLookup[$user.UserPrincipalName] = $user.AccountEnabled
        }
        if (![string]::IsNullOrWhiteSpace($user.Mail)) {
            $userLookup[$user.Mail] = $user.AccountEnabled
        }
    }
}

foreach ($row in $csvData) {
    $targetEmail = $row.Email
    if ($userLookup.ContainsKey($targetEmail)) {
        $row.'Account enabled' l]
    } else {
        $row.'Account enabled' = "Not Found"
    }
}

$csvData | Export-Csv -Path $exportPath -NoTypeInformation
