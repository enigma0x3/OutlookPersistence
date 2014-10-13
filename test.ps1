while($true){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$inbox.items | foreach { 
if($_.SenderEmailAddress -match "nelsomas@umail.iu.edu" -and $_.subject -match "Test")
{Invoke-Item C:\Windows\System32\calc.exe}}}
