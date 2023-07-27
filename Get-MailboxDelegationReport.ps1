#Connect-ExchangeOnline
#Connect-AzureAD

#gathers shared mailboxes and user mailboxes
$shared_Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object Identity, UserPrincipalName, User, Alias, AccessRights | Sort-Object Identity
$user_Mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Where-Object {($_.user -like '*@*')} | Select-Object Identity, UserPrincipalName, User, AccessRights | Sort-Object Identity 
 

# $report_type = Read-Host "Would you like Mailbox Delegation Report (1) or Shared Detailed Report (2)? Please type 1 or 2 as a response"
# Write-Host "--------------------------------------------------------------------------------------"
# Write-Host "Mailbox Delegation Report: Report includes only mailboxes that have delegated access"
# Write-Host "Mailbox Detailed Report: Report inlcudes detailed report including delegation"

#create Excel file and two worksheets in the workbook
$Mailbox_Access_Report = New-Object -ComObject excel.application
$Mailbox_Access_Report.visible = $true
$Mailbox_Access_Workbook = $Mailbox_Access_Report.workbooks.add()
$Mailbox_Access_Workbook.worksheets.add()
$Mailbox_Access_Workbook.worksheets.add()

#name the worksheets
$Sharedmailbox_Sheet = $Mailbox_Access_Workbook.worksheets.item(1)
$Usermailbox_Sheet = $Mailbox_Access_Workbook.worksheets.item(2)
$Sharedmailbox_Sheet.name = 'Shared Mailbox'
$Usermailbox_Sheet.name = 'User Mailbox'

#naming the headers of each column for shared mailbox report
$Sharedmailbox_Sheet.cells.item(1,1) = 'Shared Mailbox Access'
$Sharedmailbox_Sheet.cells.item(2,1) = 'Name'
$Sharedmailbox_Sheet.cells.item(2,2) = 'Alias'
$Sharedmailbox_Sheet.cells.item(2,3) = 'Account Enabled?'
$Sharedmailbox_Sheet.cells.item(2,4) = 'Forwarding?'
$Sharedmailbox_Sheet.cells.item(2,5) = 'Forwarded Users'
$Sharedmailbox_Sheet.cells.item(2,6) = 'Forwarded SMTP Users'
$Sharedmailbox_Sheet.cells.item(2,7) = 'Full Access Users'

#naming the headers of each column for user mailbox report
$Usermailbox_Sheet.cells.item(1,1) = 'User Mailbox Access'
$Usermailbox_Sheet.cells.item(2,1) = 'Name'
$Usermailbox_Sheet.cells.item(2,2) = 'Alias'
$Usermailbox_Sheet.cells.item(2,3) = 'Forwarding?'
$Usermailbox_Sheet.cells.item(2,4) = 'Forwarded Users'
$Usermailbox_Sheet.cells.item(2,5) = 'Forwarded SMTP Users'
$Usermailbox_Sheet.cells.item(2,6) = 'Full Access Users'


$count = 3;
#loop going through each shared mailbox
foreach($smuser in $shared_Mailboxes){
    
    #$smailboxuserarray = @()
    $teststring = ""
    $ducount = 0
    $azure_Info = Get-AzureADUser -ObjectId $smuser.UserPrincipalName
    $shared_Mailbox_Permission = $smuser | Get-MailboxPermission | Where-Object { $_.AccessRights -like "*FullAccess*" -and $_.IsInherited -eq $false }
    
    #
    foreach($u in $shared_Mailbox_Permission){
        if($u.User -ne "NT AUTHORITY\SELF"){
            #$smailboxuserarray += $u.User
            $teststring += $u.User + ", "
            $ducount += 1
        }
    }

    #shows information about the shared mailbox. Only shows the mailbox if there is more than 1 delegated user (aside from NT AUTHORITY\SELF)
    if($ducount -gt 0){
        $Sharedmailbox_Sheet.cells.item($count,1) = $smuser.UserPrincipalName
        $Sharedmailbox_Sheet.cells.item($count,2) = $smuser.Alias
        $Sharedmailbox_Sheet.cells.item($count,3) = $azure_Info.accountEnabled
        
        if($null -ne $smuser.DeliverToMailboxAndForward){
            $Sharedmailbox_Sheet.cells.item($count,4) = $smuser.DeliverToMailboxAndForward
            $Sharedmailbox_Sheet.cells.item($count,5) = $smuser.ForwardingAddress
            $Sharedmailbox_Sheet.cells.item($count,6) = $smuser.ForwardingSmtpAddress
        }
        else{
            $Sharedmailbox_Sheet.cells.item($count,4) = "None"
            $Sharedmailbox_Sheet.cells.item($count,5) = "None"
            $Sharedmailbox_Sheet.cells.item($count,6) = "None"
        }
            $Sharedmailbox_Sheet.cells.item($count,7) = $teststring
    }
    else{
        continue
    }
    
        # $Sharedmailbox_Sheet.cells.item($count,1) = $smuser.UserPrincipalName
        # $Sharedmailbox_Sheet.cells.item($count,2) = $smuser.Alias
        # $Sharedmailbox_Sheet.cells.item($count,3) = $azure_Info.accountEnabled
        
        # if($null -ne $smuser.DeliverToMailboxAndForward){
        #     $Sharedmailbox_Sheet.cells.item($count,4) = $smuser.DeliverToMailboxAndForward
        #     $Sharedmailbox_Sheet.cells.item($count,5) = $smuser.ForwardingAddress
        #     $Sharedmailbox_Sheet.cells.item($count,6) = $smuser.ForwardingSmtpAddress
        # }
        # else{
        #     $Sharedmailbox_Sheet.cells.item($count,4) = "None"
        #     $Sharedmailbox_Sheet.cells.item($count,5) = "None"
        #     $Sharedmailbox_Sheet.cells.item($count,6) = "None"
        # }

        #ignore this block
        # if($shared_Mailbox_Permission.User -ne "NT AUTHORITY\SELF"){
        #     $Sharedmailbox_Sheet.cells.item($count,7) = $shared_Mailbox_Permission.User
        # }
        # foreach($user in $smailboxuserarray){
        #     $Sharedmailbox_Sheet.cells.item($count,7) = $user
        # }

        #$Sharedmailbox_Sheet.cells.item($count,7) = $teststring
    $count += 1
}
<#
foreach($umuser in $user_Mailboxe){
    $Usermailbox_Sheet.cells.item($count,1) = $umuser.name
    $Usermailbox_Sheet.cells.item($count,2) = $umuser.alias
    $Usermailbox_Sheet.cells.item($count,3) = $umuser.DeliverToMailboxAndForward
    $Usermailbox_Sheet.cells.item($count,4) = $umuser.ForwardingAddress
    $Usermailbox_Sheet.cells.item($count,5) = $umuser.ForwardingSmtpAddress

    $Usermailbox_Sheet.cells.item($count,6) = $umuser.AccessRights
}

#>
#automatically sizes columns in each sheet
$Sharedmailbox_Sheet.columns.AutoFit()
$Usermailbox_Sheet.columns.AutoFit()

#file path specifying saving workbook in .xlsx format
$Report_FilePath = 'insert file path here'

#saves Workbook in the file path specified
$Mailbox_Access_Report.displayalerts = $false
$Mailbox_Access_Workbook.saveas($Report_FilePath)
$Mailbox_Access_Report.displayalerts = $true

