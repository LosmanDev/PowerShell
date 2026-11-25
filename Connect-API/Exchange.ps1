<#
#############################################################################
Different commands to pull user/mailbox information from Exchange Online,
along with Set commands to change that same information.
#############################################################################
#>

Connect-ExchangeOnline -Device

<#
AccessRights {Editor}: User can read, create, modify, and delete standard calendar events.
None: Removes all flags. The user receives no meeting invites and cannot view private items.
Delegate: Routes calendar meeting invites to the user's inbox.
CanViewPrivateItems: Grants the user visibility into calendar events explicitly marked as "Private".
#>

$Identity = "user@company.com"
$Delegate = "delegate@company.com"
$Flags    = "None"

<#
Displays the specified user's calendar permissions and delegate flags in a formatted table.
Displays all permission entries on the mailbox's Calendar.
#>
Get-MailboxFolderPermission -Identity "${Identity}:\Calendar" -User $Delegate | Format-Table @{Name="Mailbox";Expression={$Identity}}, FolderName, User, AccessRights, SharingPermissionFlags -AutoSize -Wrap
Get-MailboxFolderPermission -Identity "${Identity}:\Calendar"

# Resets ALL calendar delegates to Editor without delegate status and suppresses confirmation.
Remove-MailboxFolderPermission -Identity "${Identity}:\Calendar" -ResetDelegateUserCollection -Confirm:$false

<#
Updates an existing calendar permission to Editor; "None" removes delegate status.
Adds Editor calendar access without delegate status; fails if the user's permission entry already exists.
#>
Set-MailboxFolderPermission -Identity "${Identity}:\Calendar" -User $Delegate -AccessRights Editor -SharingPermissionFlags $Flags
Add-MailboxFolderPermission -Identity "${Identity}:\Calendar" -User $Delegate -AccessRights Editor -SharingPermissionFlags $Flags

# Displays directly assigned mailbox permissions, excluding inherited and SELF entries.
Get-MailboxPermission -Identity "sharedmailbox@company.com" | Where-Object { -not $_.IsInherited -and $_.User -notlike "NT AUTHORITY\SELF" } | Select-Object User, AccessRights

# Displays directly assigned Send As permissions, excluding inherited and SELF entries.
Get-RecipientPermission -Identity "sharedmailbox@company.com" | Where-Object { -not $_.IsInherited -and $_.Trustee -notlike "NT AUTHORITY\SELF" } | Select-Object Trustee, AccessRights

# Displays selected mailbox properties, including addresses, visibility, Sent Items copying, and forwarding settings.
"user1@company.com","user2@company.com" | Get-Mailbox | Format-List DisplayName,Alias,UserPrincipalName,FirstName,LastName,PrimarySmtpAddress,EmailAddresses,HiddenFromAddressListsEnabled,IsMailboxEnabled,RecipientTypeDetails,MessageCopyForSentAsEnabled,ForwardingSmtpAddress,DeliverToMailboxAndForward,ForwardingAddress

# Filters the sender's last two days of trace results for matching subjects and includes trace IDs.
Get-MessageTraceV2 -SenderAddress "sender@company.com" -StartDate (Get-Date).AddDays(-2) -EndDate (Get-Date) | Where-Object { $_.Subject -match "Keyword1|Keyword2" } | Format-Table MessageTraceId, Received, RecipientAddress, Subject, Status

# Filters the sender's last seven days of trace results for recipient addresses matching a wildcard pattern.
Get-MessageTraceV2 -SenderAddress "sender@company.com" -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | Where-Object { $_.RecipientAddress -like "*DistributionList*" } | Format-Table Received, SenderAddress, RecipientAddress, Subject, Status -AutoSize

# Store the target user's primary SMTP address in a reusable variable.
$User = "user@company.com"

# Retrieve current automated meeting processing, tentative booking, cleanup, and delegate forwarding settings.
Get-CalendarProcessing -Identity $User | Format-List AutomateProcessing, AddNewRequestsTentatively, RemoveOldMeetingMessages, ForwardRequestsToDelegates

# Configure the Calendar Attendant to automatically process meeting requests and cancellations onto the calendar.
Set-CalendarProcessing -Identity $User -AutomateProcessing AutoUpdate

# List all visible and hidden inbox rules along with their forwarding, routing, and deletion actions.
Get-InboxRule -Mailbox $User -IncludeHidden | Select-Object Name, Identity, Enabled, RedirectTo, ForwardTo, MoveToFolder, DeleteMessage, Description | Format-List
