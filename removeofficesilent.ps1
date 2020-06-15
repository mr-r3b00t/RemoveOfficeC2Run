$url = "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls/OffScrubc2r.vbs"
$output = "C:\windows\Temp\OffScrubc2r.vbs"
$start_time = Get-Date

Import-Module BitsTransfer
Start-BitsTransfer -Source $url -Destination $output


Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

#remove the mark of the web
Unblock-File "c:\windows\temp\OffScrubc2r.vbs"

#run the vbs in quiet mode! (this will remove all office click to run components)

c:\windows\system32\cscript C:\windows\Temp\OffScrubc2r.vbs /Q
