$DomainWithUsername = (Get-CimInstance -Class Win32_ComputerSystem).Username
$Username = $DomainWithUsername.Split('\')[1]
$File = "C:\Users\$($Username)\AppData\Roaming\Organizer\Profiles\$($Username).default\prefs.js"

$Filter = "(&(objectCategory=User)(samAccountName=$Username))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$ADDisplayName = $ADUser.DisplayName
$EmailAddress = $ADUser.EmailAddress

$MailServer = ""

Write-Output 'user_pref("mail.ab_remote_content.migrated", 1);' | Out-File $File -Encoding UTF8
Write-Output 'user_pref("mail.account.account1.identities", "id1");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.account.account1.server", "server1");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.account.account2.server", "server2");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.account.lastKey", 2);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.accountmanager.accounts", "account1,account2");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.accountmanager.defaultaccount", "account1");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.accountmanager.localfoldersserver", "server2");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.append_preconfig_smtpservers.version", 2);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.default_charsets.migrated", 1);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.folder.views.version", 1);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.font.windows.version", 2);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.archive_folder`", `"mailbox://$($Username)@$($MailServer)/Archives`");" | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.draft_folder`", `"mailbox://$($Username)@$($MailServer)/Drafts`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.drafts_folder_picker_mode", "0");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.fcc_folder`", `"mailbox://$($Username)@$($MailServer)/Sent`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.fcc_folder_picker_mode", "0");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.fullName`", `"$($ADDisplayName)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.reply_on_top", 1);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.smtpServer", "smtp1");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.stationery_folder`", `"mailbox://$($Username)@$($MailServer)/Templates`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.tmpl_folder_picker_mode", "0");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.identity.id1.useremail`", `"$($EmailAddress)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.identity.id1.valid", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.openMessageBehavior.version", 1);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.root.none`", `"C:\\Users\\$($Username)\\AppData\\Roaming\\Organizer\\Profiles\\$($Username).default\\Mail`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.root.none-rel", "[ProfD]Mail");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.root.pop3`", `"C:\\Users\\$($Username)\\AppData\\Roaming\\Organizer\\Profiles\\$($Username).default\\Mail`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.root.pop3-rel", "[ProfD]Mail");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.check_new_mail", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.delete_by_age_from_server", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.delete_mail_left_on_server", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.directory`", `"C:\\Users\\$($Username)\\AppData\\Roaming\\Organizer\\Profiles\\$($Username).default\\Mail\\$($MailServer)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.directory-rel`", `"[ProfD]Mail/$($MailServer)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.download_on_biff", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.hostname`", `"$($MailServer)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.lastFilterTime", 27858318);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.leave_on_server", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.login_at_startup", true);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.name`", `"$($EmailAddress)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.nextFilterTime", 27858328);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.port", 995);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.socketType", 3);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.spamActionTargetAccount`", `"mailbox://$($Username)@$($MailServer)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.storeContractID", "@mozilla.org/msgstore/berkeleystore;1");' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.server.server1.type", "pop3");' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.server.server1.userName`", `"$($Username)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.smtpserver.smtp1.authMethod", 3);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.smtpserver.smtp1.hostname`", `"$($MailServer)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.smtpserver.smtp1.port", 25);' | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.smtpserver.smtp1.try_ssl", 0);' | Out-File $File -Encoding UTF8 -Append
Write-Output "user_pref(`"mail.smtpserver.smtp1.username`", `"$($Username)`");" | Out-File $File -Encoding UTF8 -Append
Write-Output 'user_pref("mail.smtpservers", "smtp1");' | Out-File $File -Encoding UTF8 -Append