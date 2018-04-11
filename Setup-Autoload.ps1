$loggedinuser = $env:USERPROFILE

###Fixed to Use USERPROFILE TO SUPPORT DOMAIN USERS
$Powershell =  "$loggedinuser\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"

$PowershellISE = "$loggedinuser\Documents\WindowsPowerShell\Microsoft.PowerShellISE_profile.ps1"

$AutoLoad = '
 $Scripts = "C:\Program Files (x86)\Vocal Recorders\WordWatch\WW5-Scripts" 
if (!(test-path $Scripts) ) {mkdir $Scripts }
$psdir= $Scripts 

#load all autoload scripts

Get-ChildItem "${psdir}\*.ps1" | %{.$_}

Write-Host "Wordwatch 5 PowerShell Functions  Loaded"

$dir = "C:\InstallFiles\Scripts"
 $Scripts = "C:\Program Files (x86)\Vocal Recorders\WordWatch\WW5-Scripts" 
if ((Test-Path -path $dir\* -Filter *WW5-INSTALL* ) -and (!(Test-Path "C:\Program Files (x86)\Vocal Recorders\WordWatch\WW5-Scripts\WW5-Install-*.ps1")))  { Copy-Item $dir\* -Filter *WW5-INSTALL*  -Destination $Scripts ; write-host "Script Found in Install Folder  and not in the Wordwatch Direcory , Copying " }
'


if (!(test-path $Powershell)) {New-Item -Path $Powershell -ItemType file -force -Value $AutoLoad ; New-Item -Path $PowershellISE -ItemType file -force -Value $AutoLoad }




