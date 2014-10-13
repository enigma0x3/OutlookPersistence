OutlookPersistence
==================

This uses a specified email address and subject in Outlook to trigger a meterpreter shell. The macro creates a VBScript file in a hidden directory that silently download and executes a powershell script that parses the default inbox for the specified email address and subject. The script runs in the background all the time and is re-executed on startup.
