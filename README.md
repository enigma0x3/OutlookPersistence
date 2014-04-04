OutlookPersistence
==================

This is code for a "slow" persistence method. It includes an Office macro that drops a vbsript into a hidden folder and then schedules a task to execute the vbscript. Once executed, the vbscript calls powershell and executes a hosted script that invokes calc.exe.
