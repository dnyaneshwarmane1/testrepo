################## Script By DNYANESHWAR #########################

Set-ExecutionPolicy unrestricted
#cls

##########################################################

do {

########################## Menu ##########################

$menu=@" 

1 Connect to O365
2 Show mailbox information including custom attributes
3 Show permisions on Mailbox
4 Show permisions on Calendar
5 Show Delegates
6 Show Mailbox/Resource specific user has access to
7 Add permissions on Mailbox
8 Add permisiions on Calendar
9 Remove permissions on Mailbox
10 Remove permissions on Calendar
11 Enable mailbox
12 Show mailbox size
13 Show permissions on Deleted Items
14 Add permissions on Deleted Items
15 Remove permissions on Deleted Items
16 Provide Delegate access on Mailbox
Q Quit

Select a task by number or Q to quit
"@ 

##########################################################

@"
"@
(Write-Host "Please select an option"-ForegroundColor green)

$r = Read-Host $menu

##########################################################

Switch ($r) {

######################### 1 ##############################

"1" {
Write-Host @"

Please enter your admin UPN and password
"@ -ForegroundColor Green
Import-Module Msonline
Connect-ExchangeOnline
#$Cred = Get-Credential
#Connect-MsolService -Credential $Cred
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
#Import-PSSession $Session -AllowClobber


Write-host "Welcome $Cred" -foregroundcolor green
}
######################## 2 ############################### 
"2" {
Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting $upn mailbox information including custom attributes
"@ -ForegroundColor cyan
Get-Mailbox -identity $upn -ResultSize Unlimited | Select-Object DisplayName,Name,PrimarySMTPAddress,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribute4,CustomAttribute5 | format-table

} 

######################## 3 ############################### 

"3" {Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting $upn mailbox access information
"@ -ForegroundColor cyan
Get-MailboxPermission -identity "$UPN" | format-table

}

######################## 4 ############################### 

"4" {Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting $upn calendar access information
"@ -ForegroundColor cyan
Get-MailboxFolderPermission -identity "$($UPN):\Calendar" | format-table

}

######################## 5 ############################### 

"5" {Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting $upn delegate access information
"@ -ForegroundColor cyan
Get-Mailbox -anr $upn | Get-CalendarProcessing | select ResourceDelegates

}

######################## 6 ############################### 

"6" {Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting all Mailbox/Resource $upn has access to
"@ -ForegroundColor cyan
Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -User $upn

}

######################## 7 ############################### 

"7" {Write-Host @"

Please enter UPN of mailbox
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you wish to add
"@ -ForegroundColor Yellow
$User = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter Access Level you wish to add

FullAccess
ExternalAccount
DeleteItem
ReadPermission
ChangePermission
ChangeOwner
"@ -ForegroundColor Yellow
$Access = Read-host
#--------------------------------------------------------#

Write-Host @"

Adding $access permisions for $user on $upn mailbox
"@ -ForegroundColor Cyan
Add-MailboxPermission -AccessRights $access -Identity $upn -User $user

Write-host "Sucsess!!!"-foregroundcolor green
}

######################## 8 ############################### 

"8" {Write-Host @"

Please enter UPN of calendar
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you wish to add
"@ -ForegroundColor Yellow
$User = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter Access Level you wish to add

Author
Contributor
Editor
None
NonEditingAuthor
Owner
PublishingEditor
PublishingAuthor
Reviewer
"@ -ForegroundColor Yellow
$Access = Read-host
#--------------------------------------------------------#

Write-Host @"

Adding $access permisions for $user on $upn calendar
"@ -ForegroundColor cyan
Add-MailboxFolderPermission -Identity "$($UPN):\Calendar" -User $User -AccessRights $Access

Write-host "Sucsess!!!"-foregroundcolor green
}

######################## 9 ############################### 

"9" {Write-Host @"

Please enter UPN of mailbox
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you wish to remove
"@ -ForegroundColor Yellow
$User = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter Access Level you wish to remove

FullAccess
SendAs
ExternalAccount
DeleteItem
ReadPermission
ChangePermission
ChangeOwner
"@ -ForegroundColor Yellow
$Access = Read-host
#--------------------------------------------------------#

Write-Host @"

Removing $access permisions for $user on $upn mailbox
"@ -ForegroundColor cyan
remove-MailboxPermission -AccessRights $access -Identity $upn -User $user

Write-host @"

Sucsess!!!
"@-foregroundcolor green
}

######################## 10 ############################### 

"10" {Write-Host @"

Please enter UPN of calendar
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you wish to remove
"@ -ForegroundColor Yellow
$User = Read-host
#--------------------------------------------------------#

Write-Host @"

Removing all permisions for $user on $upn calendar
"@ -ForegroundColor cyan
Remove-MailboxFolderPermission -Identity "$($UPN):\Calendar" -User $User

Write-host @"

Sucsess!!!
"@-foregroundcolor green
}

######################## 11 ############################### 

"11" {Write-Host @"

Please enter username of mailbox you wish to enable
"@ -ForegroundColor Yellow
$UPN = Read-host

Write-Host @"

Enabling mailbox for $upn calendar
"@ -ForegroundColor cyan
enable-remotemailbox -username $upn

Write-host @"

Sucsess!!!
"@-foregroundcolor green
}

######################## 12 ############################### 

"12" {Write-Host @"

Please enter UPN of mailbox
"@ -ForegroundColor Yellow
$UPN = Read-host

Write-Host @"

Retrieving mailbox size for $upn
"@ -ForegroundColor cyan
Get-Mailbox -identity $upn -ResultSize Unlimited | Get-MailboxStatistics | Select DisplayName,StorageLimitStatus,@{name=”TotalItemSize (MB)”;expression={[math]::Round(($_.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}},@{name=”TotalDeletedItemSize (MB)”;expression={[math]::Round(($_.TotalDeletedItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}},ItemCount,DeletedItemCount | Sort “TotalItemSize (MB)” -Descending | Format-List

Write-host @"
Sucsess!!!
"@-foregroundcolor green
}

######################## 13 ############################### 

"13" {
Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#
Write-Host @"

Getting $upn deleted items permissions
"@ -ForegroundColor cyan

Get-MailboxFolderPermission -Identity "$($upn):\deleted items"

} 

######################## 14 ###############################

"14" {
Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you would like to add
"@ -ForegroundColor Yellow
$user = Read-host

#--------------------------------------------------------#

Write-Host @"

Please enter Access Level you wish to add

Author
Contributor
Editor
None
NonEditingAuthor
Owner
PublishingEditor
PublishingAuthor
Reviewer
"@ -ForegroundColor Yellow
$Access = Read-host
#--------------------------------------------------------#
Write-Host @"

Adding $access permissions for $user on $upn deleted items

"@ -ForegroundColor cyan

Add-MailboxFolderPermission -Identity "$($upn):\deleted items" -user $user -accessrights $access

Write-host @"

Sucsess!!!
"@-foregroundcolor green
} 

######################## 15 ###############################

"15" {
Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you would like to remove
"@ -ForegroundColor Yellow
$user = Read-host

#--------------------------------------------------------#

Write-Host @"

Removing $access permissions for $user on $upn deleted items

"@ -ForegroundColor cyan

Remove-MailboxFolderPermission -Identity "$($upn):\deleted items" -user $user

Write-host @"
Sucsess!!!
"@-foregroundcolor green
}


######################## 16 ###############################

"16" {
write-warning "If you are assigning Mr.abc delegate access to Mr.xyz's calendar, make sure you remove abc's all existing access on xyz's calendar by choosing option 10 "
Write-Host @"

Please enter UPN
"@ -ForegroundColor Yellow
$UPN = Read-host
#--------------------------------------------------------#

Write-Host @"

Please enter UPN of user you would like to add Delegate access to 
"@ -ForegroundColor Yellow
$user = Read-host

#--------------------------------------------------------#

Write-Host @"

Providing Delegate access to $user on $upn Calendar. 

"@ -ForegroundColor cyan

#Remove-MailboxFolderPermission -Identity "$($upn):\deleted items" -user $user
Add-MailboxFolderPermission -Identity "$($upn):\Calendar" -User "$user" -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems


Write-host @"
Sucsess!!!
"@-foregroundcolor green
}



######################### Q ##############################



"Q" {
Write-Host @"

Quitting

...


...


...

To Confirm Quit
"@-ForegroundColor DarkYellow

Read-Host "Press ENTER" 


cls
Exit
}

############################################################

default {
Write-Host @"

Invalid Response
"@-ForegroundColor Yellow
}
}} while ($r -ne17)

###################### 
