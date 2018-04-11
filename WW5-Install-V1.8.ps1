###WW5 INSTALL SCRIPT
###WW5 Function INSTALL SCRIPT
if ($ExecPolicy -eq "Restricted" ) { Write-Log -Path $Logsfile  -level Error -message "Please Set Execution Policy to Unrestricted" }
$global:InstallFolder = "C:\InstallFiles\"
$global:LogsFile = "C:\InstallFiles\WordwatchLogs"
if (Test-path $global:InstallFolder ) {gci $global:InstallFolder -Recurse  | Unblock-File}

function Test-IsAdmin {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator"
  )
}

function Write-Log  { 
<#
    .DESCRIPTION
        Writes a log message to file and to console to a file , the log file is defined within the $LogsFile , all of the WW5 powershell functions below utilise this.
    .EXAMPLE
        Write-Log -Path $Logsfile -level Error -message "Postgres 9.3 Failed to Install"
    #>

    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path="C:\InstallFiles\WordwatchLogs", 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info",

         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
} 

function Install-Postgres93 () {
    <#
    .DESCRIPTION
            
            Postgres Version 9.3 is no longer the current supported version of postgres for wordwatch , function has been kept for backward compatibility .
        
           
            Installs Postgres 9.3 to C drive , Postgres 9.3 is no longer used for WW5 and Postgres9.5 should be used.

            First the function checks if the postgres9.3 install file is in the  $global:InstallFolder driectory if not throws error "Postgres 9.3 install file not found in $global:InstallFolder directory"

            Second It checks if postgres9.3 is already installed by checking for a registry key REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3

            Third It installs Postgres using this commandline 'start-process -FilePath "$PostgresInstallFile" -ArgumentList '--unattendedmodeui none --mode unattended --enable_acledit 1 --prefix "c:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.3" --datadir "c:\ProgramData\Vocal Recorders\WordWatch\Postgres" --superpassword postgres --serverport 5432 --servicename Postgres9.3' -wait'
        
            Forth It then checks the registry key used in step 2 to see if it has now been installed and should write to console that the install either succeeded or failed.

    .EXAMPLE
            install-postgres93
    #>

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
    if (!(test-path -path $global:InstallFolder\* -Filter *postgres*)) {Write-Log -Path $Logsfile -level Error -message  "Postgres 9.3 install file not found in $global:InstallFolder directory"; }
$PostgresInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter *postgresql-9.3.5-1*).FullName
    If (test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3) {Write-Log -Path $Logsfile -level Warn -message "Postgres 9.3 Already Installed" }
        Else {$(Write-Log -Path $Logsfile -level Info -message "Postgres 9.3 is being installed this may take a minute or so") ; start-process -FilePath "$PostgresInstallFile" -ArgumentList '--unattendedmodeui none --mode unattended --enable_acledit 1 --prefix "c:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.3" --datadir "c:\ProgramData\Vocal Recorders\WordWatch\Postgres" --superpassword postgres --serverport 5432 --servicename Postgres9.3' -wait
    If (Get-itemproperty REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3)
       {Write-Log -Path $Logsfile -level Info -message "Installed Postgres 9.3"}}
    If (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3  )) { Write-Log -Path $Logsfile -level Error -message "Postgres 9.3 Failed to Install" } 
}

function Uninstall-postgres93  {
    <#
    .DESCRIPTION
        Uninstalls postgres 9.3 but leaves the data directory intact 

        Checks if Postgres9.3 is installed by querying if this reg key exisits REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3 if it does it runs start-process -FilePath "C:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.3\uninstall-postgresql" -ArgumentList '--mode unattended ' -wait

        Is hard coded for C drive install but could be tweaked 

        Checks if the uninstall has been successfull or not 

    .EXAMPLE
        Uninstall-postgres93


    #>

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
###9.3
If (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3)) {write-log -Path $Logsfile  -level Warn -message "Postgres 9.3 Not installed"}
Else {
start-process -FilePath "C:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.3\uninstall-postgresql" -ArgumentList '--mode unattended ' -wait
if (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.3  )) { Write-Log -Path $Logsfile -level Warn -message "Postgres 9.3 has been uninstalled , Data Directory Remains" } 
}
}

function Install-Postgres95 () {
    <#
    .DESCRIPTION

        Postgres Version 9.5 is used at the database for wordwatch.

        Installs Postgres 9.5 to C drive by default , binaries and data drive  , this can be changed by passing a parameter of -driverletter x , if you need the datadrive on another disk you can amend the install script below.

        First the function checks if the postgres9.5 install file is in the  $global:InstallFolder driectory if not throws error "Postgres 9.5 install file not found in $global:InstallFolder directory"

        Second It checks if postgres9.5 is already installed by checking for a registry key REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5

        Third It installs Postgres using this commandline start-process -FilePath "$PostgresInstallFile" -ArgumentList '
        $Arg1 = '--unattendedmodeui none --mode unattended --enable_acledit 1 --prefix'
        $Arg2 = "`"$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5`""
        $Arg3 = "--datadir `"$driveletter`:\ProgramData\Vocal Recorders\WordWatch\Postgres9.5`""
        $arg4 =  "--superpassword postgres --serverport 5432 --servicename Postgres9.5"
        
        Forth It then checks the registry key used in step 2 to see if it has now been installed and should write to console that the install either succeeded or failed.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to in stall to  , default is 'C'

    .EXAMPLE
        Install-Postgres95 -driveletter c


    #>
        [CmdletBinding()] param(
    [parameter(Mandatory=$false)]
    [string]$driveletter='c'

    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
    if (!(test-path -path $global:InstallFolder\* -Filter  *postgresql-9.5.1*)) {Write-Log -Path $Logsfile -level Error -message  "Postgres 9.5 install file not found in $global:InstallFolder directory"; }
$PostgresInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter *postgresql-9.5.1*).FullName
##MSiExec
$Arg1 = '--unattendedmodeui none --mode unattended --enable_acledit 1 --prefix'
$Arg2 = "`"$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5`""
$Arg3 = "--datadir `"$driveletter`:\ProgramData\Vocal Recorders\WordWatch\Postgres9.5`""
$arg4 =  "--superpassword postgres --serverport 5432 --servicename Postgres9.5"
$Argumentstring = "$arg1 $arg2 $arg3 $arg4"
    If (test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5) {Write-Log -Path $Logsfile -level Warn -message "Postgres 9.5 Already Installed"}
        Else {$(Write-Log -Path $Logsfile -level Info -message "Postgres 9.5 is being installed this may take a minute or so") ;start-process -FilePath "$PostgresInstallFile" -ArgumentList $Argumentstring  -wait
    If (Get-itemproperty REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5)
    {Write-Log -Path $Logsfile -level Info -message "Installed Postgres 9.5"}}
    if (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5  )) { Write-Log -Path $Logsfile -level Error -message "Postgres 9.5 Failed to Install" } 
}

function Uninstall-Postgres95 () {

    <#
    .DESCRIPTION
        Uninstalls postgres 9.5 but leaves the data directory intact  , can uninstall from other driver letters using the -driveletter parameter 

        Checks if Postgres9.5 is installed by querying if this reg key exisits REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5 if it finds it then it runs 'start-process -FilePath "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\uninstall-postgresql" -ArgumentList '--mode unattended ' -wait


        CHecks if the uninstall has been successfull or not 

     .PARAMETER driveletter
        -driveletter Allows you to pass a drive letter of your choice to uninstall from  , default is 'C'

    .EXAMPLE
        Uninstall-postgres95 -driveletter c


    #>
###9.5
        [CmdletBinding()] param(
    [parameter(Mandatory=$false)]
    [string]$driveletter='c'

    )
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
If (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5)) {write-log -Path $Logsfile  -level Warn -message "Postgres 9.5 Not installed"}
Else {if (!(Test-path "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\uninstall-postgresql.exe")) {Write-Log -Path $Logsfile -level Error -message "Uninstall not Found in path , have you slected the Correct Drive ?"}
            If (get-service Postgres9.5 | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue ) { Stop-service Postgres9.5 -ea SilentlyContinue -Force} 
                if (Test-path "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\uninstall-postgresql.exe") {start-process -FilePath "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\uninstall-postgresql" -ArgumentList '--mode unattended ' -wait}

if (!(test-path REGISTRY::HKLM\SOFTWARE\PostgreSQL\Installations\postgresql-x64-9.5  )) {Write-Log -Path $Logsfile -level Warn -message "Postgres 9.5 has been uninstalled , Data Directory Remains" } 
}
}

function Install-Erlang ( ) {
    <#
    .DESCRIPTION

        Erlang (6.1) is a language that Rabbitmq uses , therefore is a pre-req for rabbitmq

        Installs Erlang , a pre-req for Rabbitmq , first the function checks if the erlang install file is in the  $global:InstallFolder using this query 'test-path -path $global:InstallFolder\* -Filter  *otp_win64_17.1*' 
        if it is not then it thows this error "Erlang installer file not found in $global:InstallFolder directory"

        It then checks if Erlang is already installed by checking this path , if installed throws this message  "Erlang Already Installed" 

        Then runs the installer , first it sets some enviroment variables 
        [Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files (x86)\RabbitMQ Server\rabbitmq_server-3.3.1\sbin", [EnvironmentVariableTarget]::Machine)
        [Environment]::SetEnvironmentVariable("ERLANG_HOME","C:\Program Files\ERL6.1","MACHINE")
        [Environment]::SetEnvironmentVariable("RABBITMQ_BASE","$driveletter`:\RabbitMQ","MACHINE")
        [Environment]::SetEnvironmentVariable("RABBITMQ_CONFIG_FILE","$driveletter`:\RabbitMQ\rabbitmq.config.example" ,"MACHINE")
        set RABBITMQ_CONFIG_FILE ="C:\RabbitMQ\rabbitmq.config.example"
        set RABBITMQ_BASE=c:\RabbitMQ

        If then check if the rabbitmq base directory has been created and will display this message if successful "Enviroment Variables Installed"

        Then runs a few pre checks to see if rabbitmq or erlang have been installed before hand (Dirty install)

        Then runs installer start-process -FilePath "$global:InstallFolder\otp_win64_17.1.exe" -ArgumentList "/S /D=C:\Program Files\ERL6.1" -wait

        Checks if the Erlang exe is now in it's install location if so prints "Installed Erlang 6.1 " , or if failed "Erlang Failed to Install"

     .PARAMETER driveletter
        -driveletter Erlang can only be installed to C at this time.

    .EXAMPLE
        install-erlang


    #>


        [CmdletBinding()] param(
    [parameter(Mandatory=$false)]
    [string]$driveletter='c'
    )
$ErrorActionPreference = "SilentlyContinue"
$ErlangPath = "C:\Program Files\ERL6.1\bin\erl.exe"
###Erlang Test if already INstall if Not Install
### Enviroment Varibles
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
if (!(test-path -path $global:InstallFolder\* -Filter  *otp_win64_17.1*)) {Write-Log -Path $Logsfile -level Error -message  "Erlang installer file not found in $global:InstallFolder directory"}
if (Test-Path $ErlangPath  ) { write-log -Path $Logsfile  -level Warn -message "Erlang Already Installed" } 
Else {
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files (x86)\RabbitMQ Server\rabbitmq_server-3.3.1\sbin", [EnvironmentVariableTarget]::Machine)
[Environment]::SetEnvironmentVariable("ERLANG_HOME","C:\Program Files\ERL6.1","MACHINE")
[Environment]::SetEnvironmentVariable("RABBITMQ_BASE","$driveletter`:\RabbitMQ","MACHINE")
[Environment]::SetEnvironmentVariable("RABBITMQ_CONFIG_FILE","$driveletter`:\RabbitMQ\rabbitmq.config.example" ,"MACHINE")
set RABBITMQ_CONFIG_FILE ="C:\RabbitMQ\rabbitmq.config.example"
set RABBITMQ_BASE=c:\RabbitMQ
if (!(test-path C:\RabbitMQ)) {new-item -Path C:\RabbitMQ -ItemType Directory -force > $null } 
write-log -Path $Logsfile  -level Info -message "Enviroment Variables Installed"

#### Erlang Install

##Check If Erlang Process is Running from a Uninstall
if (!(Test-Path "C:\Program Files\ERL6.1\bin\erl.exe"  )) { write-log -Path $Logsfile  -level Warn -message "Enviroment Variables Installed" "Erlang Process Not Found Running Clean up" $(Get-Process -ProcessName *epmd* | Stop-Process -Force) $(start-sleep -second 3) $(Get-ChildItem -Path "C:\Program Files\ERL6.1\*" -Recurse | Remove-Item -Force -Recurse  )  } 
if (!(Test-Path "C:\Program Files\ERL6.1\bin\erl.exe"  )) { write-log -Path $Logsfile  -level Warn -message "Also Running RabbitMQ Clean Up" $(stop-service  RabbitMQ -Force) $(start-sleep -second 3) $(Get-Process -ProcessName *erlsrv* | Stop-Process -Force) $(Get-ChildItem -Path "C:\Program Files\ERL6.1\*" -Recurse | Remove-Item -Force -Recurse  )  } 
start-process -FilePath "$global:InstallFolder\otp_win64_17.1.exe" -ArgumentList "/S /D=C:\Program Files\ERL6.1" -wait
##write-log -Path $Logsfile -level Info -message "Install for Elang Pause"
##Start-Sleep -Seconds 30
If (test-path  -path $ErlangPath)
{write-log -Path $Logsfile  -level Info -message "Installed Erlang 6.1 "}
if (!(test-path $ErlangPath )) { write-log -Path $Logsfile  -level Error -message "Erlang Failed to Install" } 
}
}

function Uninstall-Erlang () {

    <#
    .DESCRIPTION
        First checks if erlang exe is in it's path 'test-path  -path $ErlangPath' if it is then it  continues , if not throws error  "Erlang Not Install"

        2nd step is to run the uninstaller, ' start-process -FilePath "C:\Program Files\ERL6.1\Uninstall.exe" -ArgumentList "/S / " -Wait'

        it then stops the rabbitmq service as it may be using the Erlang process , and epmd process and then checks the erlang path , if no longer there reports "Erlang Has been Uninstall"

     .PARAMETER driveletter
        -driveletter Erlang can only be installed to C at this time.

    .EXAMPLE
        uninstall-erlang


    #>
$ErlangPath = "C:\Program Files\ERL6.1\bin\erl.exe"
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
If (!(test-path  -path $ErlangPath)) { write-log -Path $Logsfile  -level Warn -message "Erlang Not Install" }
Else {
If (test-path  -path $ErlangPath) {start-process -FilePath "C:\Program Files\ERL6.1\Uninstall.exe" -ArgumentList "/S / " -Wait }
if (Test-Path "C:\Program Files\ERL6.1\bin\erl.exe"  ) {$(Get-Process -ProcessName *epmd* | Stop-Process -Force) ; (start-sleep -second 3); $(Get-ChildItem -Path "C:\Program Files\ERL6.1\*" -Recurse | Remove-Item -Force -Recurse  )  } 
if (Test-Path "C:\Program Files\ERL6.1\bin\erl.exe"  ) {$(stop-service  RabbitMQ -Force); $(start-sleep -second 3) ; (Get-Process -ProcessName *erlsrv* | Stop-Process -Force) ;$(Get-ChildItem -Path "C:\Program Files\ERL6.1\*" -Recurse | Remove-Item -Force -Recurse  )  } 
If (!(test-path  -path $ErlangPath)) {write-log -Path $Logsfile  -level Warn -message "Erlang Has been Uninstall"}
}
}

function Install-RabbitMQ () {

    <#
    .DESCRIPTION
        Rabbitmq is a messaging system , Wordwatch uses it as a message relay system (logs) and for heartbeats between services (status page)

        Rabbitmq can only be installed to C , the base directory can be placed on a another drive , contact dev support for details

        1st the function checks the rabbitmq install file is in the  $global:InstallFolder by testing this path test-path -path $global:InstallFolder\* -Filter  *rabbitmq-server-3.3.* if not thows  "Rabbitmq installer file not found in $global:InstallFolder directory" 

        It then checks if Erlang has already been installed as Rabbitmq requires erlang. If erlang is not installed thows this error "Erlang must be Installed First , Please Run Install-Erlang"

        Then refreshes some variables placed in the PATH , it then checks the user running the installer and checks if a cookie has been previously created if it finds one in the windows or userprofile thorws this message
         " Cookie in User Folder Found , Removing" or " Cookie in Windows Directory Found , Removing"  

         It then does some clean up's from potential bad installs by removing any temp files or base directory.

         Then runs the installer 'start-process -FilePath "$global:InstallFolder\rabbitmq-server-3.3.1.exe" -ArgumentList "/S"'

         Moves the config to the base directory  'Move-item C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ\rabbitmq.config.example  C:\RabbitMQ -force'

         It then checks if rabbitmq is installed by testing the path of the exe , if there are  reports "Installed RabbitMQ 3.3.1 or if failed "Rabbit MQ Failed Failed to Install"

         It then calls the rabbitmq-config , for more information see the rabbitmq-config function for help


     .PARAMETER driveletter
        -driveletter Rabbitmq can only be installed to C at this time.

    .EXAMPLE
        install-RabbitMQ


    #>

$RabbitMQPath = "C:\Program Files (x86)\RabbitMQ Server\rabbitmq_server-3.3.1"
$ErlangPath = "C:\Program Files\ERL6.1\bin\erl.exe"
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
set RABBITMQ_BASE="c:\RabbitMQ"
set RABBITMQ_CONFIG_FILE ="C:\RabbitMQ\rabbitmq.config.example"
if (!(test-path -path $global:InstallFolder\* -Filter  *rabbitmq-server-3.3.*)) {Write-Log -Path $Logsfile -level Error -message  "Rabbitmq installer file not found in $global:InstallFolder directory"}
if (Test-Path $RabbitMQPath  ) {  write-log -Path $Logsfile  -level Warn -message "RabbitMQ already Installed"} 
elseif (!(Test-Path $ErlangPath )) { write-log -Path $Logsfile  -level Warn -message "Erlang must be Installed First , Please Run Install-Erlang"} 
else {
foreach($level in "Machine","User") {
   [Environment]::GetEnvironmentVariables($level).GetEnumerator() | % {
      # For Path variables, append the new values, if they're not already in there
      if($_.Name -match 'Path$') { 
         $_.Value = ($((Get-Content "Env:$($_.Name)") + ";$($_.Value)") -split ';' | Select -unique) -join ';'
      }
      $_
   } | Set-Content -Path { "Env:$($_.Name)" }
}
####Rabbit MQ INstall
###Pre Rabbit Check
$loggedinuser = [Environment]::UserName
if (Test-Path c:\users\$loggedinuser\.erlang.cookie )  { write-log -Path $Logsfile  -level Warn -message " Cookie in User Folder Found , Removing" ;$(Remove-item c:\users\$loggedinuser\.erlang.cookie -force) }
if (Test-Path 'c:\windows\.erlang.cookie ') { write-log -Path $Logsfile  -level Warn -message " Cookie in Windows Directory Found , Removing"; $(Remove-item C:\windows\.erlang.cookie -force)  }
if (Test-Path C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ) { remove-item -path C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ  -Recurse -force }
if (Test-Path C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ) { remove-item -path C:\RabbitMQ  -Recurse -force }


start-process -FilePath "$global:InstallFolder\rabbitmq-server-3.3.1.exe" -ArgumentList "/S"
write-host -ForegroundColor White "Installing RabbitMQ "
Start-Sleep -Seconds 15
if (Test-Path C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ\rabbitmq.config.example ) { Move-item C:\Users\$loggedinuser\AppData\Roaming\RabbitMQ\rabbitmq.config.example  C:\RabbitMQ -force }
If (test-path  -path "$RabbitMQPath ") {write-log -Path $Logsfile  -level Info -message "Installed RabbitMQ 3.3.1 "}
if (!(test-path $RabbitMQPath)) { write-log -Path $Logsfile  -level Error -message "Rabbit MQ Failed Failed to Install"}

rabbitmq-config
}
}

function rabbitmq-config () {

    <#
    .DESCRIPTION

        Rabbitmq Config requires us to enable  the managment page and add a user used by the wordwatch service to send and recieve messages off the bus.


        1st checks that rabbitmq is installed if not throws message "RabbitMQ Must be Installed First"  

        Then refreshes the enviroment variables to memory , and enables the management page of rabbit 'rabbitmq-plugins enable rabbitmq_management'

        Restarts the rabbit service and moves the cookie file to the user profile

        It then calls the add-rabbituser function , for more details see the help for add-rabbituser


    .EXAMPLE
        rabbitmq-config


    #>


$RabbitMQPath = "C:\Program Files (x86)\RabbitMQ Server\rabbitmq_server-3.3.1"
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
set RABBITMQ_CONFIG_FILE ="C:\RabbitMQ\rabbitmq.config.example"
if (!(Test-Path $RabbitMQPath  )) {  write-log -Path $Logsfile  -level Warn -message "RabbitMQ Must be Installed First"} 
Else {
## Rabbit Config
$loggedinuser = [Environment]::UserName
##write-host -ForegroundColor White "Erlang and Rabbit Installed, Moving RabbitMQ Cookie"
#move-item -path C:\Windows\.erlang.cookie  -destination C:\users\administrator\ -force
### Refreshing Enviroment Varibles 
write-log -Path $Logsfile  -level Info -message "Session Enviroment updating"
foreach($level in "Machine","User") {
   [Environment]::GetEnvironmentVariables($level).GetEnumerator() | % {
      # For Path variables, append the new values, if they're not already in there
      if($_.Name -match 'Path$') { 
         $_.Value = ($((Get-Content "Env:$($_.Name)") + ";$($_.Value)") -split ';' | Select -unique) -join ';'
      }
      $_
   } | Set-Content -Path { "Env:$($_.Name)" }
}

write-log -Path $Logsfile  -level Info -message "Session Enviroment Updated , Running Rabbit Config"
#move-item -path C:\Windows\.erlang.cookie  -destination C:\users\administrator\ -force
rabbitmq-plugins enable rabbitmq_management
restart-service  rabbitmq 
move-item -path C:\Windows\.erlang.cookie  -destination C:\users\$loggedinuser\ -force
Add-rabbituser
}
}

Function Add-rabbituser () {
    <#
    .DESCRIPTION

        Rabbitmq Config requires us to enable  the managment page and add a user used by the wordwatch service to send and recieve messages off the bus


        Thows message  "Rabbit Setup Adding User accounts"
        Then uses rabbit commandline to add user Admin and set a password and set permission to all exchanges

        Then fires a HTTP get request to the rabbit managemnt page if successfull throws message  "RabbitMQ Responded to Get Request , Rabbit MQ Installed"

        This part can fail if IE has not been opened before hand.
        
          


    .EXAMPLE
        Add-rabbituser


    #>


start-sleep -Seconds 10
write-log -Path $Logsfile  -level Info -message "Rabbit Setup Adding User accounts"
$secpasswd = ConvertTo-SecureString 'guest' -AsPlainText -Force
$credGuest = New-Object System.Management.Automation.PSCredential ('guest', $secpasswd)
start-sleep -Seconds 5 ; rabbitmqctl add_user Admin Wordwatch1 ; rabbitmqctl set_permissions -p / Admin “.*” “.*” “.*"
write-log -Path $Logsfile  -level Info -message 'Using Rabbit Command line'
#Start-Process -FilePath 'C:\Program Files\Internet Explorer\iexplore.exe'

#new-item -ItemType file -Path c:\RabbitUser.bat
#Add-Content -Value "rabbitmqctl add_user admin Wordwatch1
#rabbitmqctl set_permissions -p / admin “.*” “.*” “.*"
#" -Path c:\RabbitUser.bat
'Retrieiving perms for new user to confirm...'
##Invoke-RestMethod 'http://localhost:15672/api/permissions/%2f/admin'  -Method get  -credential $credGuest
$request = Invoke-WebRequest -Uri http:\\localhost:15672
if ( $request.StatusCode -eq 200 ) {write-log -Path $Logsfile -message "RabbitMQ Responded to Get Request , Rabbit MQ Installed"}}

function Uninstall-RabbitMQ () {

    <#
    .DESCRIPTION
        Rabbitmq is a messaging system , Wordwatch uses it as a message relay system (logs) and for heartbeats between services (status page)

        Rabbitmq can only be installed to C , the base directory can be placed on a another drive , contact dev support for details

        1st checks if rabbitmq is installed by testing the rabbitmq path , if not there throws message "RabbitMQ Not Installed"

        2nd it stops the rabbitmq process and erlang process and runs 'start-process -FilePath "C:\Program Files (x86)\RabbitMQ Server\uninstall.exe" -ArgumentList "/S" -Wait'

        3rd it checks the path again and if its no longer there throws "RabbitMQ Has Been Uninstalled" or if still there "RabbitMQ Failed to Uninstall"
            

     .PARAMETER driveletter
        -driveletter Rabbitmq can only be installed to C at this time.

    .EXAMPLE
        uninstall-RabbitMQ


    #>


$RabbitMQPath = "C:\Program Files (x86)\RabbitMQ Server\rabbitmq_server-3.3.1"
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started $Functionname "
If (!(test-path  -path $RabbitMQPath)) {write-log -Path $Logsfile  -level warn -message "RabbitMQ Not Installed" }
Else {
If (get-service RabbitMQ | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue) {stop-service  RabbitMQ -Force; start-sleep -second 3 ; Get-Process -ProcessName *epmd* | Stop-Process -Force; Get-Process -ProcessName *erlsrv* | Stop-Process -Force; Stop-service RabbitMQ -ea SilentlyContinue  -Force} 
start-process -FilePath "C:\Program Files (x86)\RabbitMQ Server\uninstall.exe" -ArgumentList "/S" -Wait ; 
If (!(test-path $RabbitMQPath )) {write-log -Path $Logsfile  -level Warn -message "RabbitMQ Has Been Uninstalled"} 
else {Write-Log -Path $Logsfile  -level Error -message  "RabbitMQ Failed to Uninstall"}  }
 }

function Install-Server () {

    <#
    .DESCRIPTION

       Wordwatch.Server is the main process of wordwatch , acting as the Web server , API  and UI for Wordwatch. Wordwatch uses Nancyfx as it webserver but it is contain within the Wordwatch process.

       The server Install file contains Server , Compliance Hold , Compliance Manager , Management tool and retention

       1st checks if server is already installed by  using a wmi query 'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch”'

       2nd Checks if there are more than one file matching the wordwatch server file name , if more than one throws  "File Filter Detected more than one file match ,  run '$global:InstallFolder\* -Filter WordWatch.5.*' to see results "

       3rd it checks if the Wordwatch server install file is in the install directory  $global:InstallFolder\* -Filter WordWatch.5.*.msi  if not throws "Server install file not found in $global:InstallFolder directory"

       4th runs install  via msiexec with msi logs being dumped to "$global:InstallFolder`Server.log"
        
       $Arg1 = "/i"
       $Arg2 = $ServerInstallFile
       $Arg3 = "INSTALLFOLDER=`"$Driveletter"
       $arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
       $arg5 = '/qn /l*v'
       $arg6 = "$global:InstallFolder`Server.log"
       $Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
       $newArgumentstring = $Argumentstring

       5th Check if install was successful or not by using the same WMI query as above

       6th Also the prints the version installed to screen 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        Install-server -driveletter c


    #>



    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$SeverPath = "$Driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\Wordwatch.server.exe"
$ServerInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter WordWatch.5.*.msi).FullName
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
##MSiExec
$Arg1 = "/i"
$Arg2 = $ServerInstallFile
$Arg3 = "INSTALLFOLDER=`"$Driveletter"
$arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
$arg5 = '/qn /l*v'
$arg6 = "$global:InstallFolder`Server.log"
$Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
$newArgumentstring = $Argumentstring
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch”}) {write-log -Path $Logsfile  -level Warn -message "WordWatch Server Already Installed" ; } 
    Else{
    if ($ServerInstallFile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run '$global:InstallFolder\* -Filter WordWatch.5.*' to see results "}
        else{
        if (!(test-path -path $global:InstallFolder\* -Filter WordWatch.5.*)) {Write-Log -Path $Logsfile -level Error -message  "Server install file not found in $global:InstallFolder directory"}   
            Else {
            if (!(Get-WindowsFeature Desktop-Experience | where-object {$_.Installed -eq $true})) { write-log -Path $Logsfile  -level warn -message "Desktop-Experience Not Detected Installing ,  Reboot Required " ;  Add-WindowsFeature  Desktop-Experience } 
            Start-Process msiexec -ArgumentList $newArgumentstring   -wait
            $Wordwatchserver = (get-item $SeverPath).VersionInfo | select ProductVersion -ErrorAction SilentlyContinue
            If (test-path $SeverPath ) { write-log -Path $Logsfile  -level Info -message  "Installed Server Version $WordwatchServer"}}}

if (!(test-Path $SeverPath  )) { write-log -Path $Logsfile  -level Error -message "WordWatch Server Failed to install" } 

}
}

function Uninstall-Server () {

    <#
    .DESCRIPTION

       Wordwatch.Server is the main process of wordwatch , acting as the Web server and UI for wordwatch. Wordwatch uses nancyfx as it webserver but is contain within the wordwatch process.

       The server Install file contains Server , Compliance Hold , Compliance Manager , Management tool and retention so these will also be uninstalled

       1st if attempts to backup the server config file and renames it to the \Wordwatch.server.exe.config.Dateofbackup$Date+Serverversion$Serverversion"

       2nd Checks if server is installed or not by using wmi query  'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch”})'  if not throws "Server Not Installed"

       3rd Stops services which will also be uninstalled , the services are (server , Retention , compiance manager and  complicance hold) then uninstalls via wmi

       4th then checks the server exe path to see if uninstall was successful or not.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        unInstall-server -driveletter c


    #>

    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$SeverPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\Wordwatch.server.exe"
$SeverConfigPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\Wordwatch.server.exe.config"
if (test-path $SeverConfigPath) {$Serverversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$SeverPath").FileVersion}
$Date = get-date -format ("yyyMMddHHmm") 
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
if (!(test-path $SeverConfigPath )) {Write-Log -Path $Logsfile  -level Warn -message  "Server Config Not found"}
    Else {Copy-Item -path $SeverConfigPath   -Destination "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\Wordwatch.server.exe.config.Dateofbackup$Date+Serverversion$Serverversion"}
If (!( Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch”})) {write-log -Path $Logsfile  -level Warn -message "Server Not Installed" }
If (test-path $SeverPath) {
If (get-service wordwatch.server | Where-Object {$_.Status -eq "Running"}) { Stop-service wordwatch.server -ea SilentlyContinue  -Force} 
If (get-service WordWatch.Compliance.Manager | Where-Object {$_.Status -eq "Running"}) { Stop-service WordWatch.Compliance.Manager -ea SilentlyContinue  -Force} 
If (get-service WordWatch.ComplianceHold.Manager | Where-Object {$_.Status -eq "Running"}) { Stop-service WordWatch.ComplianceHold.Manager -ea SilentlyContinue -Force} 
If (get-service WordWatch.Ingester -ErrorAction SilentlyContinue | Where-Object {$_.Status -eq "Running"}) { Stop-service Wordwatch.Ingester -ea SilentlyContinue -Force} 
If (get-service WordWatch.Retention.Manager | Where-Object {$_.Status -eq "Running"}) { Stop-service WordWatch.Retention.Manager -ea SilentlyContinue -Force}
if ((get-process WordWatch.ComplianceHold.Manager -ea SilentlyContinue  ) -or (get-process WordWatch.Compliance.Manager -ea SilentlyContinue  )) {Stop-Process -ProcessName WordWatch.Compliance.Manager -force -ErrorAction SilentlyContinue ; stop-process -ProcessName WordWatch.ComplianceHold.Manager -ErrorAction SilentlyContinue  -Force}


$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -eq "Wordwatch" }
    $($app.Uninstall()) >$null
        if (-not(test-path $SeverPath)) {Write-Log -Path $Logsfile  -level info -message  "Server Has Been Uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "Server Failed to Uninstall"}
}
}

function Install-Ingester () {


    <#
    .DESCRIPTION

       Wordwatch.Ingester processes audio files and metadata and insert the metadata via the Wordwatch.server API and then move the audio to a share for server to retrieve and process.

       1st checks if ingester is already installed by  using a wmi query 'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Ingester”}'

       2nd Checks if there are more than one file matching the wordwatch ingester  file name , if more than one throws "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder\* -Filter *WordWatch.Ingester.*' to see results "

       3rd it checks if the Wordwatch ingester install file is in the install directory "Ingester install file not found in $global:InstallFolder directory"

       4th runs install  via msiexec with msi logs being dumped to "$global:InstallFolder`Ingester.log"
        
       ##MSiExec
       $Arg1 = "/i"
       $Arg2 = $IngesterInstallFile
       $Arg3 = "INSTALLFOLDER=`"$Driveletter"
       $arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
       $arg5 = '/qn /l*v'
       $arg6 = "$global:InstallFolder`Ingester.log"
       $Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"

       5th Checks if install was successful or not by using the same WMI query as above

       6th Also the prints the version installed to screen 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        Install-ingester -driveletter c


    #>


    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )

$IngesterPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe"
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
$IngesterInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter *WordWatch.Ingester.*).FullName
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
##MSiExec
$Arg1 = "/i"
$Arg2 = $IngesterInstallFile
$Arg3 = "INSTALLFOLDER=`"$Driveletter"
$arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
$arg5 = '/qn /l*v'
$arg6 = "$global:InstallFolder`Ingester.log"
$Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Ingester”}) {write-log -Path $Logsfile  -level Warn -message "WordWatch Ingester Already Installed" }
    Else{
    if ($IngesterInstallFile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder\* -Filter *WordWatch.Ingester.*' to see results "}
        else{   
        if (!(test-path -path $global:InstallFolder\* -Filter *WordWatch.Ingester.*)) {Write-Log -Path $Logsfile -level Error -message  "Ingester install file not found in $global:InstallFolder directory"}
            Else {
            if (!(Get-WindowsFeature Desktop-Experience | where-object {$_.Installed -eq $true})) { write-log -Path $Logsfile  -level warn -message "Desktop-Experience Not Detected Installing ,  Reboot Required " ;  Add-WindowsFeature  Desktop-Experience }
            Start-Process msiexec -ArgumentList $Argumentstring   -wait
            $WordwatchIngester = (get-item $IngesterPath).VersionInfo | select ProductVersion 
            If (test-path  -path $IngesterPath)  { Write-Log -Path $Logsfile  -level Info -message "Installed Ingester Version $WordwatchIngester"}}}
if (!(test-path $IngesterPath )) { Write-Log -Path $Logsfile  -level Error -message  "Ingester Failed to Install" } 
}
}

function Uninstall-Ingester () {

    <#
    .DESCRIPTION

       Wordwatch.Ingester processes audio files and metadata and insert the metadata via the Wordwatch.server API and then move the audio to a share for server to retrieve and process.

       1st if attempts to backup the server config file and renames it to the '\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config.Dateofbackup$Date+Ingesterversion$Ingesterversion"'

       2nd Checks if ingester is installed or not by using wmi query  Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Ingester”})  if not throws "Ingester Not Installed"

       3rd then checks the server exe path to see if uninstall was successful or not.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to uninstall from  , default is 'C'

    .EXAMPLE
        unInstall-ingester -driveletter c


    #>

    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$IngesterPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe"
$IngesterConfigPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config"
$Date = get-date -format ("yyyMMddHHmm") 
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
if (test-path $IngesterConfigPath) {$Ingesterversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$IngesterPath ").FileVersion}
if (!(test-path $IngesterConfigPath )) {Write-Log -Path $Logsfile  -level Warn -message  "Ingester Config Not found"}
    Else {Copy-Item -path "$IngesterConfigPath"  -Destination "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config.Dateofbackup$Date+Ingesterversion$Ingesterversion"}
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
If (!( Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Ingester”})) {write-log -Path $Logsfile  -level Warn -message "Ingester Not Installed" } 
if (test-path $IngesterPath) {
    If (get-service  WordWatch.Ingester | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue ) { Stop-service WordWatch.Ingester -ea SilentlyContinue  -Force} 
        $app = Get-WmiObject -Class Win32_Product | Where-Object { 
            $_.Name -eq "Wordwatch Ingester" }
            $($app.Uninstall()) >$null 
                if (-not(test-path $IngesterPath)) {Write-Log -Path $Logsfile  -level info -message  "Ingester Has Been Uninstalled" }
                    else {Write-Log -Path $Logsfile  -level Error -message  "Ingester Failed to Uninstall"}

}
}

function Install-Grazer () {
    <#
    .DESCRIPTION

       Wordwatch.Grazer downloads audio from a Red box recorder via the API

       1st checks if Grazer is already installed by  using a wmi query 'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Grazer”} '

       2nd Checks if there are more than one file matching the wordwatch   file name , if more than one throws "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder\* -Filter *WordWatch.Ingester.*' to see results "

       3rd it checks if the Wordwatch Grazer install file is in the install directory "Grazer install file not found in $global:InstallFolder directory"

       4th runs install  via msiexec with msi logs being dumped to "$global:InstallFolder`Grazer.log"
        
        ##MSiExec
        $Arg1 = "/i"
        $Arg2 = $GrazerInstallFile
        $Arg3 = "INSTALLFOLDER=`"$Driveletter"
        $arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
        $arg5 = '/qn /l*v'
        $arg6 = "$global:InstallFolder`Grazer.log"
        $Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"

       5th Check if install was successful or not by using the same WMI query as above

       6th Also the prints the version installed to screen 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        Install-grazer -driveletter c


    #>


    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$GrazerPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe"
$GrazerInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter *WordWatch.RbrGrazer.*).FullName
##MSiExec
$Arg1 = "/i"
$Arg2 = $GrazerInstallFile
$Arg3 = "INSTALLFOLDER=`"$Driveletter"
$arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
$arg5 = '/qn /l*v'
$arg6 = "$global:InstallFolder`Grazer.log"
$Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Grazer”}  ) {write-log -Path $Logsfile  -level Warn -message "Grazer Already Installed" } 
Else { 
    if ($GrazerInstallFile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter WordWatch.Monitor.Installer.*' to see results "}
    Else {
    if (!(test-path -path $global:InstallFolder\* -Filter *WordWatch.RbrGrazer.*)) {Write-Log -Path $Logsfile -level Error -message  "Grazer install file not found in $global:InstallFolder directory"}
    Else {Start-Process msiexec -ArgumentList $Argumentstring  -wait
    $WordwatchGrazer = (get-item $GrazerPath).VersionInfo | select ProductVersion 
    If (test-path $GrazerPath) { write-log -Path $Logsfile  -level Info -message  "Installed Grazer Version $WordwatchGrazer"}}}
if (!(test-path $GrazerPath )) {  write-log -Path $Logsfile  -level Warn -message "Grazer Failed to Install" }
}
}

function Uninstall-Grazer () {


    <#
    .DESCRIPTION

       Wordwatch.Grazer downloads audio from a Red box recorder via the API

       1st it attempts to backup the grazer config file and renames it to the '\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config.Dateofbackup$Date+Grazerversion$Grazerversion'

       2nd Checks if Grazer is installed or not by using wmi query  (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Grazer”})  if not throws "Grazer Not Installed"

       3rd then check the grazer exe path to see if uninstall was successful or not.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to uninstall from  , default is 'C'

    .EXAMPLE
        Uninstall-Grazer -driveletter c


    #>


    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Uninstaller Started for $Functionname "
$GrazerPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe"
$GrazerConfigPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config"
$Date = get-date -format ("yyyMMddHHmm")
$GrazerInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter *WordWatch.RbrGrazer.*).FullName
if (test-path $GrazerConfigPath) {$Grazerversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe").FileVersion}
if (!(test-path $GrazerConfigPath )) {Write-Log -Path $Logsfile  -level Warn -message  "Grazer Config Not found"}
    Else {Copy-Item -path "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config"  -Destination "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config.Dateofbackup$Date+Grazerversion$Grazerversion"}
If (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Grazer”})) {write-log -Path $Logsfile  -level Warn -message "Grazer Not Installed"}
If (test-path $GrazerPath) {
If (get-service  WordWatch.Grazer | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue ) { Stop-service Wordwatch.grazer -ea SilentlyContinue -Force}
$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -eq "Wordwatch Grazer" }
    $($app.Uninstall()) >$null 
        if (-not(test-path $GrazerPath)) {Write-Log -Path $Logsfile  -level info -message  "Grazer Has Been Uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "Grazer Failed to Uninstall"}
}
}

function Install-Monitor () {

    <#
    .DESCRIPTION

       Wordwatch.Monitor is used to relay messages via rabbitmq to Monior.Alarm for alerts and feedback information into the Wordwatch.Server status page

       1st checks if monitor is already installed by  using a wmi query 'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor” '

       2nd Checks if there are more than one file matching the wordwatch.monitor   file name , if more than one throws "File Filter Detected more than one file match ,  run 'File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter WordWatch.Monitor.Installer.*' to see results "

       3rd it checks if the Wordwatch monitor install file is in the install directory "Monitor install file not found in $global:InstallFolder directory"

       4th runs install  via msiexec with msi logs being dumped to "$global:InstallFolder`Monitor.log"
        
        ##MSiExec
        $Arg1 = "/i"
        $Arg2 = $MonitorInstallFile
        $Arg3 = "INSTALLFOLDER=`"$Driveletter"
        $arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
        $arg5 = '/qn /l*v'
        $arg6 = "$global:InstallFolder`Monitor.log"
        $Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"

       5th Check if install was successful or not by using the same WMI query as above

       6th Also the prints the version installed to screen 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        Install-monitor -driveletter c


    #>

    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$MonitorPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe"
$MonitorInstallFile  = (Get-ItemProperty -Path $global:InstallFolder\* -Filter WordWatch.Monitor.Installer*).FullName 
##MSiExec
$Arg1 = "/i"
$Arg2 = $MonitorInstallFile
$Arg3 = "INSTALLFOLDER=`"$Driveletter"
$arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
$arg5 = '/qn /l*v'
$arg6 = "$global:InstallFolder`Monitor.log"
$Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor”}  ) {  write-log -Path $Logsfile -level warn -message "Monitor Already Installed" } 
else{
if ($MonitorInstallFile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter WordWatch.Monitor.Installer.*' to see results "}

else {
if (!(test-path -path $global:InstallFolder\* -Filter WordWatch.Monitor.Installer*)) {write-log -Path $Logsfile -level Error -message  "Monitor install file not found in $global:InstallFolder directory"}
       Else {start-process msiexec  -ArgumentList $Argumentstring  -wait
            $WordwatchMonitor = (get-item $MonitorPath).VersionInfo | select ProductVersion 
                If (test-path $MonitorPath) {write-log -Path $Logsfile  -level Info -message "Installed Monitor Version $WordwatchMonitor"}}}
if (!(test-path $MonitorPath)) {  write-log -Path $Logsfile  -level Error -message  "Monitor Failed to Install" } 
}
}

function Uninstall-Monitor () {


    <#
    .DESCRIPTION

       Wordwatch.Monitor is used to relay messages via rabbitmq to Monior.Alarm for alerts and feedback information into the Wordwatch.Server status page

       1st if attempts to backup the monitor config file and renames it to the \Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe.config.Dateofbackup$Date+Monitorversion$Monitorversion"

       2nd Checks if monitor is installed or not by using wmi query  Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor”})  if not throws "Monitor Not Installed"

       3rd then checks the monitor exe path to see if uninstall was successful or not.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to uninstall from  , default is 'C'

    .EXAMPLE
        Uninstall-Monitor -driveletter c


    #>
    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "$Functionname Started"
$MonitorPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe"
$MonitorConfigPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe"
$Date = get-date -format ("yyyMMddHHmm")
if (test-path $MonitorConfigPath) {$Monitorversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe").FileVersion}
if (!(test-path $MonitorConfigPath )) {Write-Log -Path $Logsfile  -level Warn -message  "Monitor Config Not found"}
    Else {Copy-Item -path "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe.config"  -Destination "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe.config.Dateofbackup$Date+Monitorversion$Monitorversion"}
If (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor”})) {write-log -Path $Logsfile  -level Warn -message "Monitor Not Installed"}
If (test-path $MonitorPath) {
If (get-service "wordwatch monitor" | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue) { Stop-service  "wordwatch monitor" -ea SilentlyContinue -Force}
$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -eq "Wordwatch Monitor" }
    $($app.Uninstall()) >$null 
        if (-not(test-path $MonitorPath)) {Write-Log -Path $Logsfile  -level info -message  "Monitor Has Been Uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "Monitor Failed to Uninstall"}
}
}

function Install-Alarm () {


    <#
    .DESCRIPTION

       Wordwatch.Monitor.Alarm connects to rabbitmq and is used to sent out alerts via SMTP or syslog

       1st checks if Alarm is already installed by  using a wmi query Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor Alarm”}  )'

       2nd Checks if there are more than one file matching the wordwatch.monitor.alarm   file name , if more than one throws ""File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter WordWatch.Monitor.Alarm*' to see results " "

       3rd it checks if the Wordwatch monitor.alarm install file is in the install directory "Alarm install file not found in $global:InstallFolder directory"

       4th runs install via msiexec with msi logs being dumped to ""$global:InstallFolder`Alarm.log""
        
        ##MSiExec
        $Arg1 = "/i"
        $Arg2 = $AlarmInstallFile
        $Arg3 = "INSTALLFOLDER=`"$Driveletter"
        $arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
        $arg5 = '/qn /l*v'
        $arg6 = "$global:InstallFolder`Alarm.log"
        $Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"

       5th Checks if install was successful or not by using the same WMI query as above

       6th Also print the version installed to screen 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to install to  , default is 'C'

    .EXAMPLE
        Install-Alarm -driveletter c


    #>

    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$AlarmPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe"
$AlarmInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter WordWatch.Monitor.Alarm*).FullName
##MSiExec
$Arg1 = "/i"
$Arg2 = $AlarmInstallFile
$Arg3 = "INSTALLFOLDER=`"$Driveletter"
$arg4 = ':\Program Files (x86)\Vocal Recorders\WordWatch"'
$arg5 = '/qn /l*v'
$arg6 = "$global:InstallFolder`Alarm.log"
$Argumentstring = "$arg1 `"$arg2`" $arg3$arg4 $arg5 $arg6"
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor Alarm”}  ) { write-log -Path $Logsfile -level warn -message "Alarm Already Installed" } 
    else{
    if ($AlarmInstallFile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter WordWatch.Monitor.Alarm*' to see results "}  
        Else {
        if (!(test-path -path $global:InstallFolder\* -Filter WordWatch.Monitor.Alarm.*)) {write-log -Path $Logsfile -level Error -message "Alarm install file not found in $global:InstallFolder directory"}
             else {Start-Process msiexec -ArgumentList $Argumentstring -wait
                $WordwatchAlarm = (get-item $AlarmPath).VersionInfo | select ProductVersion 
                    If (test-path  -path $AlarmPath) {write-log -Path $Logsfile  -level Info -message "Installed Alarm Version $WordwatchAlarm"}}}
if (!(test-path $AlarmPath)) { write-log -Path $Logsfile  -level Error -message "Alarm Failed to Install" } 
}
}

function Uninstall-Alarm () {

    <#
    .DESCRIPTION

       Wordwatch.Monitor.Alarm connects to rabbitmq and is used to sent out alerts via SMTP or syslog

       1st it attempts to backup the alarm config file and renames it to the \Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\Wordwatch.Monitor.Alarm.exe.config.Dateofbackup$Date+Monitorversion$Monitorversion"

       2nd Checks if monitor is installed or not by using wmi query  'Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor Alarm”}  )'  if not throws "Alarm Not Installed"

       3rd then checks the alarm exe path to see if uninstall was successful or not.

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to uninstall from  , default is 'C'

    .EXAMPLE
        Uninstall-alarm -driveletter c


    #>
    Param(
    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 
    )
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "$Functionname Started"
$AlarmPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe"
$AlarmConfigPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config"
$Date = get-date -format ("yyyMMddHHmm")
if (test-path $AlarmConfigPath) {$AlarmVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\Wordwatch.Monitor.Alarm.exe").FileVersion}
if (!(test-path $AlarmConfigPath )) {Write-Log -Path $Logsfile  -level Warn -message  "Alarm Config Not found"}
    Else {Copy-Item -path "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config"  -Destination "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\Wordwatch.Monitor.Alarm.exe.config.Dateofbackup$Date+Monitorversion$Monitorversion"} 
If (-not(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Wordwatch Monitor Alarm”}  )) {write-log -Path $Logsfile  -level Warn -message "Alarm Not Installed"}
If (test-path $AlarmPath) { 
$app = Get-WmiObject -Class Win32_Product | Where-Object {
If (get-service "wordwatch alarm" | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue) { Stop-service  "wordwatch alarm" -ea SilentlyContinue -Force} 
    $_.Name -eq "Wordwatch Monitor Alarm" }
    $($app.Uninstall()) >$null 
        if (-not(test-path $AlarmPath)) {Write-Log -Path $Logsfile  -level info -message  "Alarm Has Been Uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "Alarm Failed to Uninstall"}
}
}

function postgres-backup () {

    <#
    .DESCRIPTION

       Executes a Postgres Backup of the wordwatch Database utilising pg_dump

       1st checks the logged in user and does a test-path against the pgpass file , if not created sets up using default creds

       2nd Executes backup using 'C:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\bin\pg_dump.exe'; can be amended if on another disk

       3rd Reports backup and how long backup took to complete  to windows events using source pg_dump

       Can be setup as schedule task example  'SchTasks /Create /SC DAILY /TN “Postgres Backup and House Keeping” /TR 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command ''postgres-backup; postgres-housekeeping'' ' /ST 09:00'


    .EXAMPLE
        postgres-backup


    #>


#############################################  
## PostgreSQL PGDump 
## Author: Toby Hersey  
## Date : 10 May 2015      
#############################################
$loggedinuser = [Environment]::UserName


####Postgre Pre-Check - OLD method
#IF (-not(test-path C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf))  {
#New-Item -Path C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf -Force -ItemType file 
#Add-Content -Path C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf -Value "localhost:5432:*:postgres:postgres"
#}

####Postgres Pgpass content check
$file = Get-Content "C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf"
$SEL = Select-String -Path "C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf" -Pattern 'localhost:5432:*:postgres:postgres' -SimpleMatch

if ($SEL -ne $null)
{
   write-log -Path $Logsfile -level warn -message "Pgpass file already exists"
}
 else {
  New-Item -Path C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf -Force -ItemType file
Add-Content -Path C:\Users\$loggedinuser\AppData\Roaming\postgresql\pgpass.conf -Value "localhost:5432:*:postgres:postgres"
write-log -Path $Logsfile -level warn -message "Postgres pgpass file was did not exist, was blank or incorrect text inside file"
}
 
# Define Paths
$BackupRoot = 'C:\Database\Backup\';
$BackupLabel = (Get-Date -Format 'yyyy-MM-dd_HHmmss');
if (!(test-path $BackupRoot)) {mkdir $BackupRoot} 
# Postgres Bin Location
$PgBackupExe = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\PostreSQL\9.5\bin\pg_dump.exe';
$PgUser = 'postgres';
 

# log settings
$EventSource = 'pg_Dump';

# log erros to Windows Application Event Log
function Log([string] $message, [System.Diagnostics.EventLogEntryType] $type){
    # create EventLog source
    if (![System.Diagnostics.EventLog]::SourceExists($EventSource)){
        New-Eventlog -LogName 'Application' -Source $EventSource;
    }

    # write to EventLog
    Write-EventLog -LogName 'Application'`
        -Source $EventSource -EventId 1 -EntryType $type -Message $message;

}
 
$BackupDir = Join-Path $BackupRoot $BackupLabel;
$PgBackupErrorLog = Join-Path $BackupRoot ($BackupLabel + '-tmp.log');
 
# execution time
$StartTS = (Get-Date);

# start pg_basebackup
try
{
    Start-Process $PgBackupExe -ArgumentList "--host", "localhost", "--port 5432", "--username", "postgres", "--role", "postgres", "--no-password", "--format custom", "--verbose", "--file $BackupRoot$BackupLabel", "wordwatch" -Wait -NoNewWindow -RedirectStandardError $PgBackupErrorLog;
}
catch
{
    Write-Error $_.Exception.Message;
    Log $_.Exception.Message Error;
    ;
}

# check pg_basebackup output
If (Test-Path $PgBackupErrorLog){
 
    # read errors
    $errors = Get-Content $PgBackupErrorLog;
 
    If($errors -eq $null){
        # backup successful, purge old backups
        Purge $BackupRoot;
    }
    #delete tmp error log
    Remove-Item $PgBackupErrorLog -Force;
}

# Log backup duration
$ElapsedTime = $(get-date) - $StartTS;
Log "Successful PG Backup done in $($ElapsedTime.TotalMinutes) minutes" Information;






}

function Postgres-Housekeeping () {

    <#
    .DESCRIPTION

       Used to provide postgres backups with a retention on X amount of days , Defined within $limit = (Get-Date).AddDays(-30)

       Default path for files to check is held within this variable $BackupRoot = "C:\Database\Backup\";

       Deletes files older than $limit in the $backuproot directory and puts the message in windows events as source postgres-housekeeping

       Can be setup as schedule task example  'SchTasks /Create /SC DAILY /TN “Postgres Backup and House Keeping” /TR 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command ''postgres-backup; postgres-housekeeping'' ' /ST 09:00'


    .EXAMPLE
        postgres-housekeeping


    #>

$BackupRoot = "C:\Database\Backup\";

$limit = (Get-Date).AddDays(-30)
$path = "$BackupRoot"
$message = "Files older the $limit have been deleted from $path"
# Delete files older than the $limit.
Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force

# Delete any empty directories left behind after deleting the old files.
Get-ChildItem -Path $path -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse

    if (![System.Diagnostics.EventLog]::SourceExists("Postgres-Housekeeping")){
        New-Eventlog -LogName 'Application' -Source "Postgres-Housekeeping";
    }

 Write-EventLog -LogName 'Application'`
        -Source Postgres-Housekeeping -EventId 1 -EntryType Information -Message $message;

}

function Add-User () {
    <#
    .DESCRIPTION
        Imports a CSV File with the following Headers RoleID,FullName,Email,Username and Password. RoleID and Token MUST be created on the server before hand
        Creates the users by making an API call to the WW5 Server.
    .PARAMETER File
        The -File Parameter should be the full path of your CV file.
    .PARAMETER Server
        The -Server Parameter is the address of your WW5 , IP address is preferred, but DNS should work if resolvable.
    
    .PARAMETER token
        The -token Parameter is how Request to the server are authenicated.
    .EXAMPLE
        Add-User -file "C:\users\devadmin\Desktop\starwars.csv" -server 10.10.10.109 -token "eyJ0eXAiOiJKV1QiLCJhbGcaW9wb3J0YWwiLCJF58rJIJmGwbd_Yw"
    #>



    
    [CmdletBinding()] param(
    [parameter(Mandatory=$true)]
    [string]$file ,
    [string]$server,
    [string]$token
    )
#$token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJJc3N1ZXIiOiJodHRwOi8vd3d3LmJ1c2luZXNzc3lzdGVtc3VrLmNvbS93b3Jkd2F0Y2giLCJBdWRpZW5jZSI6Imh0dHA6Ly93d3cuYnVzaW5lc3NzeXN0ZW1zdWsuY29tL3dvcmR3YXRjaC9wb3J0YWwiLCJFeHBpcnkiOiJcL0RhdGUoMjUzNDAyMzAwNzk5OTk5KVwvIiwiTmFtZSI6IkltcG9ydC1Vc2VycyIsIlBlcm1pc3Npb25zIjpbIlVzZXJBZG1pbiJdfQ.notCai7lsaQA2QUspp2Z6-mZZp50KjM7s8jeitHdctw"
#$server = "10.10.10.125"
#$file =  "C:\Intel\Starwars.csv" 

#add-user -server $server -file $file -token $token


$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Function  $Functionname Started "
#Start-Transcript -literalpath c:\failedusers.txt -force -append
$DebugPreference = "Continue"
##Import the CSV file
if (!(Test-path $file )) {write-log -Path $Logsfile  -level Error -message "CSV File Not Found"}
 Else {$import = Import-Csv $file 
    ##Make objects out of the Data
    ForEach ($item in $import)
{ 
  $roleID = $($item.RoleID)
  $Fullname = $($item.Fullname)
  $email = $($item.Email)
  $Username = $($item.Username)
  $Password = $($item.Password)
  $Active = $($item.Active)

##Create JSON form
$newuser = @"
{
"roleID":"$roleID",
"fullName":"$fullname",
"email":"$email",
"userName":"$username",
"password":"$password",
"active":"$active"
}
"@



try {
      invoke-webrequest -Method POST -uri "http://$server/api/users" -Body $newuser -Headers @{Authorization = $token; Accept = 'application/json'} -ContentType application/json       
       }
catch {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $responseBody = $reader.ReadToEnd();
        If ({$result.StatusCode -eq 400 -or $result.status -eq 500})
        { write-log -Path $Logsfile  -level Warn -message "Failed to add User + $newuser + $responseBody" }
        Else { write-log -Path $Logsfile  -level Info -message "Added User $newuser + $responseBody " }      
}
}
}
}

function Get-Version () {
    <#
    .DESCRIPTION

       Get-version will print the version of the Installed Wordwatch modules to screen.
       Example
       Server Process not Found displayed if process not found
       Your Version of Grazer is 5.5.6.0 displayed if process found

 

     .PARAMETER driveletter
            The -driveletter Allows you to pass a drive letter of your choice to check for version information  , default is 'C'

    .EXAMPLE
        Get-Version -driveletter c


    #>


[CmdletBinding()] param(
    [parameter(Mandatory=$false)]
    [string]$Server ,
    [string]$Ingester,
    [string]$Grazer,
    [string]$Monitor,
    [string]$Alarm,
    [Parameter(Mandatory=$true)] 
    [string]$driveletter 

    ) 
$ErrorActionPreference = "SilentlyContinue"


$ServerPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\Wordwatch.server.exe"
$Serverversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ServerPath).FileVersion

$IngesterPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe"
$Ingesterversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($IngesterPath).FileVersion


$GrazerPath="$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe"
$Grazerversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($GrazerPath).FileVersion

$MonitorPath="$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\Wordwatch.Monitor.exe"
$Monitorversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($MonitorPath).FileVersion

$AlarmPath="$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\Wordwatch.Monitor.Alarm.exe"
$AlarmVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($AlarmPath).FileVersion




if (test-path $Serverpath) { Write-Host "Your Version of Server is $ServerVersion"}
Else {Write-host "Server Process not Found"  }

if (test-path $IngesterPath) { Write-Host "Your Version of Ingester is $IngesterVersion"}
Else {Write-host "Ingester Process not Found"  }

if (test-path $GrazerPath) {Write-Host "Your Version of Grazer is $GrazerVersion" }
Else {Write-host "Grazer Process not Found"  }

if (test-path $MonitorPath) { Write-Host "Your Version of Monitor is $Monitorversion"}
Else {Write-host "Monitor Process not Found"  }

if (test-path $AlarmPath) { Write-Host "Your Version of Alarm is $AlarmVersion"}
Else {Write-host "Alarm Process not Found"  }
   
}

function Zip-Directory {
    Param(
      [Parameter(Mandatory=$True)][string]$DestinationFileName,
      [Parameter(Mandatory=$True)][string]$SourceDirectory,
      [Parameter(Mandatory=$False)][string]$CompressionLevel = "Optimal",
      [Parameter(Mandatory=$False)][switch]$IncludeParentDir
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $CompressionLevel    = [System.IO.Compression.CompressionLevel]::$CompressionLevel  
    [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectory, $DestinationFileName, $CompressionLevel, $IncludeParentDir)
}

function get-logs ($server , $ingester ,$grazer , $compliancemanager , $compliancehold , $retention , [int]$limit ) {



 
$ErrorActionPreference = "SilentlyContinue"
param ($server , $ingester ,$grazer ,[int]$limit  )
$date =(get-date -Format 'yyyy-MM-dd_HHmmss')
$limitforlogs = (Get-Date).AddHours($limit)
$Serverlogs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Server\logs'
#$Serverlogname = $storage + "\Serverlogs" + $date + ".zip"
$ServerLogCollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\ServerLogCollection'
$Ingesterlogs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\logs'
#$Ingesterlogname = $storage + "\Ingesterlogs" + $date + ".zip"
$IngesterLogCollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\IngesterLogCollection'
$Grazerlogs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\logs'
#$Grazerlogname = $storage + "\Grazerlogs" + $date + ".zip"
$GrazerCurrent = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\logs'
$GrazerLogCollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\GrazerLogCollection'
$Compliancemanagerlogs = "C:\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\logs"
$CompliancemanagerLogCollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\Compliancemanager'
$ComplianceHoldlogs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\logs'
$ComplianceHoldLogCollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\ComplianceHold'
$Retentionlogs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\logs'
$RetentionLogcollection = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection\Retention'


$Logs = 'C:\Program Files (x86)\Vocal Recorders\WordWatch\LogCollection'

if (!(Test-path $ServerLogCollection)) {remove-item -path $ServerLogCollection -Filter *.* ;  mkdir $ServerLogCollection}
if (!(Test-path $ingesterLogCollection )) {remove-item -path  $ServerLogCollection -Filter *.* ;    mkdir $ingesterLogCollection}
if (!(Test-path $grazerLogCollection )) {remove-item -path $grazerLogCollection  -Filter *.* ;   mkdir $grazerLogCollection}
if (!(Test-path $CompliancemanagerLogCollection )) {remove-item -path $CompliancemanagerLogCollection   -Filter *.* ;   mkdir $CompliancemanagerLogCollection }
if (!(Test-path $ComplianceHoldLogCollection )) {remove-item -path $ComplianceHoldLogCollection  -Filter *.* ;   mkdir $ComplianceHoldLogCollection }
if (!(Test-path $RetentionLogcollection )) {remove-item -path $RetentionLogcollection -Filter *.* ;   mkdir $RetentionLogcollection }


if ($server -eq "Y") {
    if (!(test-path $serverlogs)) { Write-Host "Server Log directory not found" }
        else { Get-childitem  $Serverlogs -Recurse | where-object { $_.LastWriteTime -gt $limitforlogs} | copy-item -Destination $ServerLogCollection  }}
if ($ingester -eq "Y") {
    if (!(test-path $ingesterlogs)) { Write-Host "Ingester Log directory not found" }
        else {Get-childitem  $Ingesterlogs -Recurse | where-object { $_.LastWriteTime -gt $limitforlogs} | copy-item -Destination  $IngesterLogCollection  }}
if ($grazer -eq "Y") {
    if (!(test-path $grazerlogs)) { Write-Host "Grazer Log directory not found" }
        else { Get-childitem  $Grazerlogs -Recurse | where-object {$_.LastWriteTime -gt $limitforlogs} | copy-item -Destination $GrazerLogCollection  }}
if ($compliancemanager -eq "Y") {
    if (!(test-path $Compliancemanagerlogs)) { Write-Host "Compliance Log directory not found" }
        else { Get-childitem  $Compliancemanagerlogs -Recurse | where-object {$_.LastWriteTime -gt $limitforlogs} | copy-item -Destination $CompliancemanagerLogCollection  }}
if ($compliancehold -eq "Y") {
    if (!(test-path $ComplianceHoldlogs)) { Write-Host "Compliance Hold Log directory not found" }
        else { Get-childitem  $ComplianceHoldlogs -Recurse | where-object {$_.LastWriteTime -gt $limitforlogs} | copy-item -Destination $ComplianceHoldLogCollection  }}
if ($retention -eq "Y") {
    if (!(test-path $Retentionlogs)) { Write-Host "Retention Log directory not found" }
        else { Get-childitem  $Retentionlogs -Recurse | where-object {$_.LastWriteTime -gt $limitforlogs} | copy-item -Destination $RetentionLogcollection  }}


Zip-Directory -SourceDirectory $logs -DestinationFileName "C:\Program Files (x86)\Vocal Recorders\WordWatch\WordwatchLogs$date.zip"

}

Function Configure-Configs ( ) {
    <#
    .DESCRIPTION
     Configures default Wordwatch config files with default Rabbitmq connection string and machine hostname. Note must of used the same U/P as defined in the add-rabbituser otherwise will fail.

     Function Expects the connection string to be default so the find and replace command can work.

     Can also be used to configure postgres config  with some default  settings

    .PARAMETER RabbitServer
        The -RabbitServer Parameter is used when amending the wordwatch connection string to rabbit. It defaults to rabbitmq being installed on the localhost but can be overridden by passing a hostname or IP address of rabbitmq on another host **Note Don't forget to open port :15672
    .PARAMETER Server
        The -Server should be set to 'y' if you want to configure the server config file
    
    .PARAMETER Ingester
        The -Ingester should be set to 'y' if you want to configure the Ingester config file

    .PARAMETER Grazer
        The -Grazer should be set to 'y' if you want to configure the Grazer config file
    
    .PARAMETER compliancemanager
        The -compliancemanager should be set to 'y' if you want to configure the compliancemanager config file

    .PARAMETER compliancehold
        The -compliancehold should be set to 'y' if you want to configure the compliancehold config file
    
    .PARAMETER retention
        The -retention should be set to 'y' if you want to configure the retention config file 

    .PARAMETER monitor
        The -monitor should be set to 'y' if you want to configure the monitor config file

    .PARAMETER alarm
        The -alarm should be set to 'y' if you want to configure the alarm config file

    .PARAMETER postgres
        The -postgres should be set to 'y' if you want to configure the postgres config file 

    .PARAMETER driveletter 
        The -driveletter should be set to the drive letter your config exisit on 
         

    .EXAMPLE
        Configure-Configs -server y -ingester y -grazer y -ComplianceHold y -ComplianceManager y -retention y -Monitor y -Alarm y -driveletter c
    #>


    Param 
    ( 

    [Parameter(Mandatory=$false)] 
        [string]$RabbitServer = "localhost",

        [Parameter(Mandatory=$false)] 
        [string]$server,

        [Parameter(Mandatory=$false)] 
        [string]$ingester,

        [Parameter(Mandatory=$false)] 
        [string]$grazer,

        [Parameter(Mandatory=$false)] 
        [string]$retention,

        [Parameter(Mandatory=$false)] 
        [string]$ComplianceHold ,

        [Parameter(Mandatory=$false)] 
        [string]$ComplianceManager ,

        [Parameter(Mandatory=$false)] 
        [string]$Monitor ,

        [Parameter(Mandatory=$false)] 
        [string]$Alarm ,

        [Parameter(Mandatory=$false)] 
        [string]$postgres ,

        [Parameter(Mandatory=$true)] 
        [string]$driveletter 



    ) 

    
$Hostname = [System.Net.Dns]::GetHostByName((hostname)).HostName
if (!($server -eq "y")) {}
Elseif ($server -eq "y")  { 
    if (Test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Server\WordWatch.Server.exe.config" ) {
##Configure Server Hostname
$DefaultServervalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.Server`" />"
$NewServerValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.Server`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Server\WordWatch.Server.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultServervalue , $NewServerValue  }) | Set-Content $_ }


##Configure Rabbit MQ String
$DefaultServerRabbitmqvalue =  '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0;persistentMessages=false" />'
$NewServerRabbitmqValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Server\WordWatch.Server.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultServerRabbitmqvalue , $NewServerRabbitmqValue  }) | Set-Content $_ }


##Configure Server Nlog
$Defaultserverrabbitmq_nlog = '<target name="rabbitmq_nlog" xsi:type="RabbitMQ" username="guest" password="guest" hostname="localhost" exchange="WordWatch.Messages.Monitor.Logging" port="5672" topic="Logging" vhost="/" useJSON="true" maxBuffer="10240" heartBeatSeconds="3">'
$Newserverrabbitmq_nlog = "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Server\WordWatch.Server.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $Defaultserverrabbitmq_nlog , $Newserverrabbitmq_nlog  }) | Set-Content $_ }

Write-Host "Server Configs Have Been Configured"
}
        Else {Write-host "Server Config Not found "}}



if (!( $ingester -eq "y")) {}
ElseIF ($Ingester -eq "y")  { 
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config"  ) {
##Configure Server Hostname
$Defaultingestervalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.Ingester`" />"
$NewingesterValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.Ingester`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $Defaultingestervalue , $NewingesterValue  }) | Set-Content $_ }


##Configure Rabbit MQ String
$DefaultingesterRabbitmqvalue =  '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=guest;virtualHost=/;port=5672;timeout=0;persistentMessages=false" />'
$NewingesterRabbitmqValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=Wordwatch1;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultingesterRabbitmqvalue, $NewingesterRabbitmqValue }) | Set-Content $_ }


##Configure Server Nlog
$Defaultingesterrabbitmq_nlog = '<target name="rabbitmq_nlog" xsi:type="RabbitMQ" username="guest" password="guest" hostname="localhost" exchange="WordWatch.Messages.Monitor.Logging" port="5672" topic="Logging" vhost="/" useJSON="true" maxBuffer="10240" heartBeatSeconds="3">'
$Newingesterrabbitmq_nlog = "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"Wordwatch1`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\Wordwatch.Ingester.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $Defaultingesterrabbitmq_nlog, $Newingesterrabbitmq_nlog  }) | Set-Content $_ }
Write-Host "Ingester Configs Have Been Configured"}
        Else {Write-host "Ingester Config Not found "}}


if (!( $grazer -eq "y")) {}
Elseif ($grazer -eq "y") { 
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" ) {
$GrazerPath = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe"
$WordwatchGrazer = (get-item $GrazerPath).VersionInfo | select ProductVersion 

##Configure Grazer Hostname
$DefaultGrazervalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.RbrGrazer`" />"
$NewGrazerValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.RbrGrazer`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultGrazervalue , $NewGrazerValue   }) | Set-Content $_ }


##Configure Grazer Rabbit MQ String
$DefaultGrazerRabbitmqvalue =  '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=guest;virtualHost=/;port=5672;timeout=0;persistentMessages=false" />'
$NewGrazerRabbitmqValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=Wordwatch1;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultGrazerRabbitmqvalue , $NewGrazerRabbitmqValue  }) | Set-Content $_ }


##Configure Grazer  Nlog
$DefaultGrazerrabbitmq_nlog = '<target name="rabbitmq_nlog" xsi:type="RabbitMQ" username="guest" password="guest" hostname="localhost" exchange="WordWatch.Messages.Monitor.Logging" port="5672" topic="Logging" vhost="/" useJSON="true" maxBuffer="10240" heartBeatSeconds="3">'
$NewGrazerrabbitmq_nlog = "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"Wordwatch1`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultGrazerrabbitmq_nlog, $NewGrazerrabbitmq_nlog  }) | Set-Content $_ }

  Write-host "Grazer Grazer Configs Configured" } 
        Else {Write-host "Grazer Config Not found "}}



if (!( $Retention -eq "y")) {}
Elseif ( $Retention -eq "y") { 
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\WordWatch.Retention.Manager.exe.config" ) {
##Configure retention Hostname

$DefaultRetensionvalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.Retention.Manager`" />"
$NewRetentionValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.Retention.Manager`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\WordWatch.Retention.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultRetensionvalue , $NewRetentionValue    }) | Set-Content $_ }

 
##Configure retention Rabbit MQ String

$DefaultRetentionRabbitmqvalue =  '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0;persistentMessages=false" />'
$NewretentionRabbitmqValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\WordWatch.Retention.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultRetentionRabbitmqvalue  , $NewretentionRabbitmqValue  }) | Set-Content $_ }


##Configure retention  Nlog
$Defaultretentionrabbitmq_nlog = '<target name="rabbitmq_nlog" xsi:type="RabbitMQ" username="guest" password="guest" hostname="localhost" exchange="WordWatch.Messages.Monitor.Logging" port="5672" topic="Logging" vhost="/" useJSON="true" maxBuffer="10240" heartBeatSeconds="3">'
$Newretentionrabbitmq_nlog  = "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\WordWatch.Retention.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $Defaultretentionrabbitmq_nlog , $Newretentionrabbitmq_nlog  }) | Set-Content $_ }
  Write-Host "Retention Configs Configured"    }
        Else {Write-host "Retention Config Not found "}}


if (!( $ComplianceHold -eq "y")) {}
Elseif ( $ComplianceHold -eq "y") { 
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\WordWatch.ComplianceHold.Manager.exe.config" ) {
##Configure ComplianceHold Hostname
$DefaultComplianceHoldvalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.ComplianceHold.Manager`" />"
$NewComplianceHoldValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.ComplianceHold.Manager`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\WordWatch.ComplianceHold.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultComplianceHoldvalue  , $NewComplianceHoldValue    }) | Set-Content $_ }

##Configure ComplianceHold Rabbit MQ String
$DefaultComplianceHoldvalue = '<add name="compliance_hold" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0" />'
$NewComplianceHoldValue = "<add name=`"compliance_hold`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\WordWatch.ComplianceHold.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultComplianceHoldvalue , $NewComplianceHoldValue   }) | Set-Content $_ }


##Configure ComplianceHold  Nlog       
$DefaultComplianceHoldRabbitmqvalue =  "target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"guest`" password=`"guest`" hostname=`"localhost`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
$NewComplianceHoldRabbitmqValue =     "target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\WordWatch.ComplianceHold.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultComplianceHoldRabbitmqvalue , $NewComplianceHoldRabbitmqValue }) | Set-Content $_ }
Write-host "Complicance Hold Configs Configured" 
}
        Else {"Compliance Hold Config Not found "}}


if (!( $ComplianceManager -eq "y")){}
Elseif(( $ComplianceManager -eq "y")) {    
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\WordWatch.Compliance.Manager.exe.config" ) {
 ##Configure ComplianceHold Hostname
$DefaultCompliancevalue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"Hostname.WordWatch.Compliance.Manager`" />"
$NewComplianceValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.Compliance.Manager`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\WordWatch.Compliance.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultCompliancevalue  , $NewComplianceValue    }) | Set-Content $_ }

##Configure ComplianceHold Rabbit MQ String
$DefaultCompliancevalue = "<add name=`"compliance_export`" connectionString=`"host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0`" />"
$NewComplianceValue = "<add name=`"compliance_export`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\WordWatch.Compliance.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultCompliancevalue , $NewComplianceValue   }) | Set-Content $_ }


##Configure ComplianceHold  Nlog       
$DefaultComplianceRabbitmqvalue =  "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"guest`" password=`"guest`" hostname=`"localhost`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
$NewComplianceRabbitmqValue =     "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\WordWatch.Compliance.Manager.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultComplianceRabbitmqvalue , $NewComplianceRabbitmqValue }) | Set-Content $_ }
  Write-host "Complicance Manager Configs Configured" 
  
  }
        Else{Write-host "Compliance Manager Config Not found "}}


if (!( $Monitor -eq "y")){}
Elseif(( $Monitor -eq "y")) {    
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\WordWatch.Monitor.exe.config" ) {
 ##Configure Montitor Hostname
$DefaultMonitorvalue = '<add key="WordWatch.Monitor.SubscriptionId" value="Hostname.WordWatch.Monitor" />'
$NewMonitorValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.Monitor`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\WordWatch.Monitor.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultMonitorvalue  , $NewMonitorValue    }) | Set-Content $_ }

##Configure Monitor Rabbit MQ String
$DefaultMonitorvalue = '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0" />'
$NewMonitorValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\WordWatch.Monitor.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultMonitorvalue , $NewMonitorValue   }) | Set-Content $_ }


##Configure Monitor  Nlog       
$DefaultMonitorRabbitmqvalue =  "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"guest`" password=`"guest`" hostname=`"localhost`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
$NewMonitorRabbitmqValue =     "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor\WordWatch.Monitor.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultMonitorRabbitmqvalue , $NewMonitorRabbitmqValue }) | Set-Content $_ }
        Write-Host "Montior Configs have been Configured"}
        Else{Write-host "Monitor Config Not found "}}
        

if (!( $Alarm -eq "y")){}
Elseif(( $Alarm -eq "y")) {    
    if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config" ) {
 ##Configure Alarm Hostname
$DefaultAlarmvalue = '<add key="WordWatch.Monitor.SubscriptionId" value="Hostname.WordWatch.Alarm" />'
$NewAlarmValue = "<add key=`"WordWatch.Monitor.SubscriptionId`" value=`"$Hostname.WordWatch.WordWatch.Alarm`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultAlarmvalue   , $NewAlarmValue    }) | Set-Content $_ }

##Configure Alarm Rabbit MQ String
$DefaultAlarmvalue = '<add name="rabbitmq_heartbeat" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0;persistentMessages=false" />'
$NewAlarmValue = "<add name=`"rabbitmq_heartbeat`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0;persistentMessages=false`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultAlarmvalue , $NewAlarmValue   }) | Set-Content $_ }

###Alarm Connection string
                      ###<add name="rabbitmq_alarm" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0" />
$DefaultalarmString = '<add name="rabbitmq_alarm" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0" />'
$NewAlarmString = "<add name=`"rabbitmq_alarm`" connectionString=`"host=$RabbitServer;username=Admin;password=39769F419A1B285DF075ED5F7DE2FE33;virtualHost=/;port=5672;timeout=0`" />"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultalarmString , $NewAlarmString  }) | Set-Content $_ }

###<add name="rabbitmq_alarm" connectionString="host=localhost;username=guest;password=A7DBCF870700FE61F5465BA3DAC2B0B9;virtualHost=/;port=5672;timeout=0" />
##Configure Alarm  Nlog       
$DefaultAlarmRabbitmqvalue =  "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"guest`" password=`"guest`" hostname=`"localhost`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
$NewAlarmRabbitmqValue =     "<target name=`"rabbitmq_nlog`" xsi:type=`"RabbitMQ`" username=`"Admin`" password=`"39769F419A1B285DF075ED5F7DE2FE33`" hostname=`"$RabbitServer`" exchange=`"WordWatch.Messages.Monitor.Logging`" port=`"5672`" topic=`"Logging`" vhost=`"/`" useJSON=`"true`" maxBuffer=`"10240`" heartBeatSeconds=`"3`">"
gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Monitor Alarm\WordWatch.Monitor.Alarm.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultAlarmRabbitmqvalue  , $NewAlarmRabbitmqValue  }) | Set-Content $_ }
        Write-Host "Alarm Configs have been configured"}
        Else{Write-host "Alarm Config Not Found "}}

if (!($Postgres -eq "y" )) {}
ElseIF ($Postgres -eq "y" ) {
    if(test-path $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" ) {
###Effective Postgres Value
$DefaultEffectivePostgresValue = "#effective_cache_size = 2600MB"
$NewEffectivePostgresValue = "effective_cache_size = 4GB" 
gci $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultEffectivePostgresValue , $NewEffectivePostgresValue    }) | Set-Content $_ }

###Shared Buffers Postgres Value
$DefaultEffectivePostgresValue = "shared_buffers = 128MB"
$NewEffectivePostgresValue = "shared_buffers = 1000MB" 
gci $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultEffectivePostgresValue , $NewEffectivePostgresValue   }) | Set-Content $_ }

###Work Mem Postgres Value
$DefaultEffectivePostgresValue = "#work_mem = 4MB"
$NewEffectivePostgresValue = "work_mem = 128MB" 
gci $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultEffectivePostgresValue , $NewEffectivePostgresValue   }) | Set-Content $_ }


###Maintenance Postgres Value
$DefaultEffectivePostgresValue = "#maintenance_work_mem = 64MB"
$NewEffectivePostgresValue = "maintenance_work_mem = 1000MB" 
gci $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultEffectivePostgresValue , $NewEffectivePostgresValue   }) | Set-Content $_ }


###Default Stat Postgres Value
$DefaultEffectivePostgresValue = "#default_statistics_target = 100"
$NewEffectivePostgresValue = "default_statistics_target = 1000" 
gci $driveletter":\ProgramData\Vocal Recorders\WordWatch\Postgres9.5\postgresql.conf" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultEffectivePostgresValue , $NewEffectivePostgresValue   }) | Set-Content $_ }
  Write-Host "Postgres Config Configured"
            $Ask = Read-Host "Postgres Configured with New values Would you like to restart Enter Y / N "
            if ($ask -eq "Y" ) { stop-service Wordwatch.Ingester -ea SilentlyContinue ; stop-service Wordwatch.Server -ea SilentlyContinue  ; stop-service Postgres9.5 -ea SilentlyContinue   ; Start-service Postgres9.5 -ea SilentlyContinue   ; Start-service wordwatch.server -ea SilentlyContinue  ; Start-Service wordwatch.ingester -ea SilentlyContinue  ; start-service Wordwatch.ingester -ea SilentlyContinue  }

}
Else {Write-Host Postgres Config Not Found!}
}


}

function WW5-Ports () {

    <#
    .DESCRIPTION
       Opens the required firewall ports on the domain profile for Wordwatch Server to function this includes Postgres , Powershell , Browser 80 & 443 and rabbitmq.

       Will only work for OS 2012R2 and above. Will throw error message "This Function only supports OS 2012R2 and Higher" if below OS.build 9000


    .EXAMPLE
        ww5-ports
    #>



    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 

$OS = [environment]::OSVersion.Version

if ($os.build -lt "9000") { Write-Log -Path $Logsfile -level Error -message  "This Function only supports OS 2012R2 and Higher"  } 
Else {


###Turn On Windows Powershell Domain Profile
Set-NetFirewallProfile -Profile Domain -Enabled True


###Web browser Access Inbound
if (Get-NetFirewallRule -Name "WW5 Browser Access"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "WW5 Browser Access Rule Already Exisits"}
Else {New-NetFirewallRule -name "WW5 Browser Access"-DisplayName "WW5 Browser Access" -Action Allow -Description "Browser Access" -Direction Inbound -Enabled True -Group Domain -LocalPort 80 ,443 -Protocol TCP > $NULL 
if (Get-NetFirewallRule -Name "WW5 Browser Access"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message  "WW5 Browser Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "WW5 Browser Access Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }


###Rabbit Mq Tool And Que
if (Get-NetFirewallRule -Name "Rabbit MQ Tool and Que"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Rabbit MQ Tool and Que Rule Already Exisits"}
Else {New-NetFirewallRule -Name "Rabbit Mq Tool and Que" -DisplayName "Rabbit Mq Tool and Que" -Action Allow -Description "Rabbit Mq Tool and Que" -Direction Inbound -Enabled True -Group Domain -LocalPort 5672, 15672 -Protocol TCP > $NULL 
if (Get-NetFirewallRule -Name "Rabbit MQ Tool and Que"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message  "Rabbit MQ Tool and Que Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "WW5 Browser Access Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }

##New-NetFirewallRule -Name "Rabbit Mq Tool and Que" -DisplayName "Rabbit Mq Tool and Que" -Action Allow -Description "Rabbit Mq Tool and Que" -Direction Inbound -Enabled True -Group Domain -LocalPort 5672, 15672 -Protocol TCP

###Sys Log Alarm / MONITOR ONLY
if (Get-NetFirewallRule -Name "Sys Log"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Sys Log  Rule Already Exisits"}
Else {New-NetFirewallRule -name "Sys Log" -DisplayName "Sys Log" -Action Allow -Description "Sys Log" -Direction Inbound -Enabled True -Group Domain -LocalPort 514 -Protocol TCP > $NULL
if (Get-NetFirewallRule -Name "Sys Log"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "Sys Log Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "Sys Log Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }

##New-NetFirewallRule -name "Sys Log" -DisplayName "Sys Log" -Action Allow -Description "Sys Log" -Direction Inbound -Enabled True -Group Domain -LocalPort 514 -Protocol TCP

###Archiving SMB and EMC
if (Get-NetFirewallRule -Name "Archiving SMB and EMC"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Archiving SMB and EMC  Rule Already Exisits"}
Else {New-NetFirewallRule -name "Archiving SMB and EMC"  -DisplayName "Archiving SMB and EMC" -Action Allow -Description "Archiving SMB and EMC" -Direction Inbound -Enabled True -Group Domain -LocalPort 445, 3682, 3218 -Protocol TCP > $NULL
if (Get-NetFirewallRule -Name "Archiving SMB and EMC"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "Archiving SMB and EMC Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "Archiving SMB and EMC Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }

##New-NetFirewallRule -name "Archiving SMB and EMC"  -DisplayName "Archiving SMB and EMC" -Action Allow -Description "Archiving SMB and EMC" -Direction Inbound -Enabled True -Group Domain -LocalPort 445, 3682, 3218 -Protocol TCP

if (Get-NetFirewallRule -Name "Postgres"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Postgres  Rule Already Exisits"}
Else {New-NetFirewallRule -name "Postgres" -DisplayName "Postgres" -Action Allow -Description "Postgres" -Direction Inbound -Enabled True -Group Domain -LocalPort 5432 -Protocol TCP > $NULL
if (Get-NetFirewallRule -Name "Postgres"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "Postgres Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "Postgres Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }

####New-NetFirewallRule -name "Postgres" -DisplayName "Postgres" -Action Allow -Description Postgres -Direction Inbound -Enabled True -Group Domain -LocalPort 5432 -Protocol TCP

###Alarm SMTP
if (Get-NetFirewallRule -Name "Alarm SMTP"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Alarm SMTP  Rule Already Exisits"}
Else {New-NetFirewallRule -name "Alarm SMTP" -DisplayName "Alarm SMTP" -Action Allow -Description "Alarm SMTP" -Direction Inbound -Enabled True -Group Domain -LocalPort 25 -Protocol TCP  > $NULL
if (Get-NetFirewallRule -Name "Alarm SMTP"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "Alarm SMTP Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "Alarm SMTP Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }
###New-NetFirewallRule "Alarm SMTP" -DisplayName "Alarm SMTP" -Action Allow -Description "Alarm SMTP" -Direction Inbound -Enabled True -Group Domain -LocalPort 25 -Protocol TCP

###Optional Management PSremoting and RDP
if (Get-NetFirewallRule -Name "RDP"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Postgres  Rule Already Exisits"}
Else {New-NetFirewallRule -name "RDP" -DisplayName "RDP" -Action Allow -Description Management -Direction Inbound -Enabled True -Group Domain -LocalPort 3389 -Protocol TCP  > $NULL
if (Get-NetFirewallRule -Name "RDP"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "RDP Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "RDP Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }



###Optional Management PSremoting and RDP
if (Get-NetFirewallRule -Name "RDP/UDP"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "RDP/UDP Rule Already Exisits"}
Else {New-NetFirewallRule -name "RDP/UDP" -DisplayName "RDP/UDP" -Action Allow -Description Management -Direction Inbound -Enabled True -Group Domain -LocalPort 3389 -Protocol UDP  > $NULL
if (Get-NetFirewallRule -Name "RDP/UDP"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "RDP/UDP Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "RDP/UDP Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }

 
 if (Get-NetFirewallRule -Name "Powershell Remoting"  -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level Info -message  "Powershell Remoting Rule Already Exisits"}
Else {New-NetFirewallRule -Name "Powershell Remoting" -DisplayName "PsRemoting" -Action Allow -Description Psremoting -Direction Inbound -Enabled True -Group Domain -LocalPort 5985, 5986 -Protocol TCP  > $NULL
if (Get-NetFirewallRule -Name "Powershell Remoting"   -ErrorAction SilentlyContinue ) {Write-Log -Path $Logsfile -level iNFO -message "Powershell Remoting Access Rule Created"}
Else {Write-Log -Path $Logsfile -level Error -message  "Powershell Remoting Rule Not Found after attempting to create , remove >`$null redirection to see more details"} }



}

}

function Install-Apache24-Reverse-Proxy-Server () {

    <#
    .DESCRIPTION
       Sets up a reverse proxy with Apache for panintellgince 

       Checks if Apache already installed via 'Get-WmiObject -Class Win32_Service | select Name  | where { $_.Name -eq  “Apache2.4”}'

       Checks there is only one apache24 folder in $global:InstallFolder "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter Apache24' to see results "

       Checks that the apache folder exisits or throws "Apache24 install folder not found in $global:InstallFolder directory"

       Copies files and config to "$driveletter`:\apps\Apache24"

       Installs Apache as a service "Apache2.4"

       Checks apache is running after install

      



     .PARAMETER driveletter
            The -driveletter install apache to the drive letter , default is 'C'


     .PARAMETER InsightHostname 
            The DNS address of the server running Panintelligence software


    .EXAMPLE
        Install-Apache24-Reverse-Proxy-Server -$InsightHostname dnsaddressofserver
    #>

    param(

    [Parameter(Mandatory=$true)]
    [string]$InsightHostname,

     [Parameter(Mandatory=$false)]
    [string]$driveletter = "C"

    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$Apache24Path = "$driveletter`:\apps\Apache24"
$Apache24InstallFolder = (Get-ItemProperty -Path $global:InstallFolder\* -Filter Apache24).FullName
$VCRedistInstallFile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter vcredist_x64*).FullName

if (Get-WmiObject -Class Win32_Service | select Name  | where { $_.Name -eq  “Apache2.4”} ) { write-log -Path $Logsfile -level warn -message "Apache2.4 Already Installed" }
    else {
    if ($Apache24InstallFolder.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter Apache24' to see results "}
        Else {
        if (!(test-path -path $global:InstallFolder\* -Filter Apache24)) {write-log -Path $Logsfile -level Error -message "Apache24 install folder not found in $global:InstallFolder directory"}
        else{  if (!(test-path -path $global:InstallFolder\* -Filter vcredist_x64*)) {write-log -Path $Logsfile -level Error -message "VSRedist install file not found in $global:InstallFolder directory"}
             else {
                    if (!(Test-Path -path $Apache24Path)) {

                        New-Item $Apache24Path -type directory > $null

                        }
                    Copy-Item $Apache24InstallFolder "$driveletter`:\apps\" -force -recurse  

                    & $VCRedistInstallFile /q /norestart   *>> $Logsfile
                    write-log -Path $Logsfile -level warn -message "Waiting for Microsoft Visual C++ 2012 Redistributable Package (x64) to Install."
                    Start-Sleep -second 5

                    $Apache24exe = $Apache24Path + "\bin\httpd.exe"
                    & $Apache24exe -k install  *>> $Logsfile

                    $ApacheConf = "$driveletter`:\apps\Apache24\conf\extra\"
                    $Hostname = [System.Net.Dns]::GetHostByName((hostname)).HostName
                    
                    if (Test-path $ApacheConf ) {
                        ##Configure Server Hostname
                        $DefaultServervalue = "Define SRVNAME HOSTNAME"
                        $NewServerValue = "Define SRVNAME $Hostname"

                        gci "$ApacheConf`httpd-vhosts.conf" -recurse | ForEach {
                          (Get-Content $_ | ForEach {$_ -replace $DefaultServervalue , $NewServerValue  }) | Set-Content $_ }
                        gci "$ApacheConf`httpd-ssl.conf" -recurse | ForEach {
                          (Get-Content $_ | ForEach {$_ -replace $DefaultServervalue , $NewServerValue  }) | Set-Content $_ }

                        ##Configure Insight Server Hostname
                        $DefaultInsightServervalue = "Define INSIGHTSRVNAME INSIGHTHOSTNAME"
                        $NewInsightServerValue = "Define INSIGHTSRVNAME $InsightHostname"

                        gci "$ApacheConf`httpd-vhosts.conf" -recurse | ForEach {
                          (Get-Content $_ | ForEach {$_ -replace $DefaultInsightServervalue , $NewInsightServerValue  }) | Set-Content $_ }
                        gci "$ApacheConf`httpd-ssl.conf" -recurse | ForEach {
                          (Get-Content $_ | ForEach {$_ -replace $DefaultInsightServervalue , $NewInsightServerValue  }) | Set-Content $_ }

                         Start-service  "Apache2.4" -ea SilentlyContinue;
                         Start-sleep -second 3 
                        }}}}
If (get-service "Apache2.4" | Where-Object {$_.status -eq "Running"}){write-log -Path $Logsfile  -level info -message  "Apache2.4 Has Been Installed and Running" }
ElseIf (get-service "Apache2.4" | Where-Object {$_.status -eq "Stopped"}){write-log -Path $Logsfile  -level info -message  "Apache2.4 Has Been Installed but Not Running" }
Else { write-log -Path $Logsfile  -level Error -message "Apache2.4 Failed to Install" }
}
}

function Uninstall-Apache24-Reverse-Proxy-Server () {

   <#
    .DESCRIPTION
       Sets up a reverse proxy with Apache for panintellgince 

       Checks if Apache already installed via Get-WmiObject -Class Win32_Service | select Name  | where { $_.Name -eq  “Apache2.4”}

       Checks the apache binary folder , if cannot find throws "Apache2.4 cannot be Uninstalled as binary folder not found at $Apache24Path."

       Stops apache service and runs uninstaller $Apache24exe -k uninstall

       Reports if uninstall was successful by testing apache path

     .PARAMETER driveletter
            The -driveletter install apache to the drive letter , default is 'C'



    .EXAMPLE
        Uninstall-Apache24-Reverse-Proxy-Server -driveletter c
    #>

    [Parameter(Mandatory=$false)] 
    [string]$Driveletter = "C" 



$Apache24Path = "$driveletter`:\apps\Apache24"
If (-not(Get-WmiObject -Class Win32_Service | select Name  | where { $_.Name -eq  “Apache2.4”}  )) {write-log -Path $Logsfile  -level Warn -message "Apache2.4 Not Installed"}
ElseIf (!(test-path $Apache24Path)) {write-log -Path $Logsfile  -level Warn -message "Apache2.4 cannot be Uninstalled as binary folder not found at $Apache24Path."}
Else {
    $Apache24exe = $Apache24Path + "\bin\httpd.exe"
    If (get-service "Apache2.4" | Where-Object {$_.Status -eq "Running"} -ea SilentlyContinue) 
        { 
        Stop-service  "Apache2.4" -ea SilentlyContinue -Force;
        start-sleep -second 3 ;
        & $Apache24exe -k uninstall  *>> $Logsfile;
        Remove-Item $Apache24Path -recurse;
        if (!(test-path $Apache24Path -ErrorAction SilentlyContinue)) {write-log -Path $Logsfile  -level info -message  "Apache2.4 Has Been Uninstalled" }
        else {write-log -Path $Logsfile  -level Error -message  "Apache2.4 Failed to Uninstall"}
        }}
}

function read-logs () {

  Param 
    ( 

        [Parameter(Mandatory=$true)] 
        [string]$driveletter = "c",

        [Parameter(Mandatory=$false)] 
        [string]$numberoflines = "10",

        [Parameter(Mandatory=$false)] 
        [string]$server = "",

        [Parameter(Mandatory=$false)] 
        [string]$ingester= "",

        [Parameter(Mandatory=$false)] 
        [string]$grazer = "",

        [Parameter(Mandatory=$false)] 
        [string]$compliancemanager = "",

        [Parameter(Mandatory=$false)] 
        [string]$compliancehold = "",

        [Parameter(Mandatory=$false)] 
        [string]$retention= "",


        [Parameter(Mandatory=$false)] 
        [string]$tail= ""




    ) 

$ww5serverlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Server\logs\log.current.log"
$ww5ingesterlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Ingester\logs\log.current.log"
$ww5grazerlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\logs\log.current.log"
$ww5complancemanagerlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Manager\logs\log.current.log"
$ww5complancemanagerholdlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Compliance Hold Manager\logs\log.current.log"
$ww5retentionlog = "$driveletter`:\Program Files (x86)\Vocal Recorders\WordWatch\Retention Manager\logs\log.current.log"


 
if (!( $server -eq "y" )) {}
    ElseIF (($server = "Y") -and (!($tail -eq "Y")))  { 
        if (test-path $ww5serverlog ) { get-content $ww5serverlog  | select -last $numberoflines} 
        Else {Write-host "Server Logs Not found "}}
    If (($server = "Y") -and ($tail -eq "Y")) {
        if (test-path $ww5serverlog) { get-content $ww5serverlog -Wait  -Tail 5 } 
        Else {Write-host "Server Logs Not found "}}
  
 
 
 
 
 
if (!( $ingester -eq "y" )) {}
    ElseIF (($ingester = "Y") -and (!($tail -eq "y")))  { 
        if (test-path $ww5ingesterlog ) { get-content $ww5ingesterlog  | select -last $numberoflines} 
        Else {Write-host "Ingester Logs Not found "}}
    If (($ingester = "Y") -and ($tail -eq "y")) {
        if (test-path $ww5ingesterlog) { get-content $ww5ingesterlog -Wait  -Tail 5 } 
        Else {Write-host "Ingester Logs Not found "}}


 
if (!( $grazer -eq "y" )) {}
    ElseIF (($grazer = "Y") -and (!($tail -eq "y")))  { 
        if (test-path $ww5grazerlog ) { get-content $ww5grazerlog  | select -last $numberoflines} 
        Else {Write-host "Grazer Logs Not found "}}
    If (($grazer = "y") -and ($tail -eq "y")) {
        if (test-path $ww5grazerlog) { get-content $ww5grazerlog -Wait  -Tail 5 } 
        Else {Write-host "Grazer Logs Not found "}}
 

  
 
if (!( $compliancemanager  -eq "y" )) {}
    ElseIF (($grazer = "Y") -and (!($tail -eq "y")))  { 
        if (test-path $ww5complancemanagerlog  ) { get-content $ww5complancemanagerlog   | select -last $numberoflines} 
        Else {Write-host "Compliance Manager Logs Not found "}}
    If (($compliancemanager  = "Y") -and ($tail -eq "y")) {
        if (test-path $ww5grazerlog) { get-content $ww5complancemanagerlog  -Wait  -Tail 5 } 
        Else {Write-host "Compliance Manager Logs Not found "}}
  
    
    
    if (!( $compliancehold  -eq "y" )) {}
    ElseIF (($compliancehold  = "Y") -and (!($tail -eq "y")))  { 
        if (test-path $ww5complancemanagerholdlog ) { get-content $ww5complancemanagerholdlog  | select -last $numberoflines} }
        Else {Write-host "Compliance hold logs Not found "}
    If (($compliancehold  = "Y") -and ($tail -eq "y")) {
        if (test-path $ww5complancemanagerholdlog) { get-content $ww5complancemanagerholdlog -Wait  -Tail 5 } 
        Else {Write-host "Compliance hold  Logs Not found "}}


    if (!( $retention -eq "y" )) {}
    ElseIF (($ww5retentionlog = "Y") -and (!($tail -eq "y")))  { 
        if (test-path $ww5retentionlog ) { get-content $ww5retentionlog | select -last $numberoflines} }
        Else {Write-host "Retention Logs Not found "}
    If (($retention = "Y") -and ($tail -eq "y")) {
        if (test-path $ww5retentionlog) { get-content $ww5retentionlog -Wait  -Tail 5 } 
        Else {Write-host "Retention  Logs Not found "}}  
 
          
 
}

function install-dotnet452 () {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$getnetversion = Get-ChildItem "hklm:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" | Get-ItemProperty  -Name Release | Select-Object -expandProperty Release
$dotnet452file= 'C:\Installfiles\NDP452-KB2901907-x86-x64-AllOS-ENU.exe'
$dotnetinstallfile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter NDP452-KB2901907-x86-x64-AllOS-ENU.exe*).FullName



if ($getnetversion -eq "379893" ) {write-log -Path $Logsfile -level Warn -message "Dontnet4.5.2 Already Installed" }
else { 
    write-log -Path $Logsfile -level Info -message "Dotnet4.5.2 not installed , Installing"
        if (test-path $dotnetinstallfile) {
        $arg1 = " /qb /passive /norestart"
        start-process  $dotnet452file    -ArgumentList "$arg1 " -wait
        if ($getnetversion -eq "379893" ) {write-log -Path $Logsfile -level warn -message "Dontnet4.5.2 Has been installed successfully , you will need to reboot!"}
    }
    Else {write-log -Path $Logsfile -level Error -message "Dotnet Install File not found in $global:Installfolder "}
    }


}

function uninstall-dotnet452 () {
write-log -Level Warn "Uninstalling .net 4.5.2 cannot be done from the script as powershell replies on it , please run the .net setup file manually to remove"
} 

function install-vcc () {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$vcc= 'C:\Installfiles\vcredist_x64.exe'
$vccfile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter vcredist_x64.exe* ).FullName
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq "Microsoft Visual C++ 2012 x64 Additional Runtime - 11.0.61030”}  ) { write-log -Path $Logsfile -level warn -message "VC Redist 2012 already Installed" } 
    else{
    if ($vccfile.count -gt 1 ) {write-log -path $Logsfile  -Level error -Message "File Filter Detected more than one file match ,  run 'gci $global:InstallFolder -Filter vcredist_x64*' to see results "}  
        Else {
        if (!(test-path -path $global:InstallFolder\* -Filter vcredist_x64.exe*)) {write-log -Path $Logsfile -level Error -message "VCCREDIST install file not found in $global:InstallFolder directory"}
             else {start-process $vcc -ArgumentList "/q" -wait
                    if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq "Microsoft Visual C++ 2012 x64 Additional Runtime - 11.0.61030”}  ) { write-log -Path $Logsfile -level info -message "VC Redist 2012 has been  Installed" } 
                    else {write-log -Path $Logsfile -level error -message "VC Redist 2012 has failed to Install"}
}
}
}
}

function uninstall-vcc () {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "$Functionname Started"
$vcc= 'C:\Installfiles\vcredist_x64.exe'
$vccfile = (Get-ItemProperty -Path $global:InstallFolder\* -Filter vcredist_x64.exe* ).FullName
$Date = get-date -format ("yyyMMddHHmm")
If (-not(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Microsoft Visual C++ 2012 x64 Additional Runtime - 11.0.61030”}  )) {write-log -Path $Logsfile  -level Warn -message "VCREDIST 2012 11.0.61030  Not Installed"}
else {
        if (test-path 'C:\Installfiles\vcredist_x64.exe') {start-process 'c:\Installfiles\vcredist_x64.exe' -ArgumentList "/q /uninstall /norestart" -wait}
        else {Write-Log -Path $Logsfile  -level Error -message  "$vcc not found , unable to uninstall via command line without :( "}
        If (-not(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “Microsoft Visual C++ 2012 x64 Additional Runtime - 11.0.61030”})) {Write-Log -Path $Logsfile  -level info -message  "VCREDIST 2012 11.0.61030 Has Been Uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  " VCREDIST 2012 11.0.61030 Failed to Uninstall"}
}
}

function install-prereqs () {

    <#
    .DESCRIPTION
    Installs Pre-req software for wordwatch , VCC and Dotnet4.5.2


    .EXAMPLE
        install-prereqs
    #>


install-vcc 
install-dotnet452
}

function uninstall-prereqs () {
uninstall-vcc
uninstall-dotnet452
} 

function get-ww5help () {

Write-Host "Welcome to WW5 powershell help"

Write-Host "Here is a list of Powershellfunctions "

Write-host "Write-Log"

Write-Host "Install-Postgres95"

Write-Host "UnInstall-Postgres95"

Write-Host "Install-erlang"

Write-Host "UnInstall-erlang"

Write-Host "Install-rabbitmq"

Write-Host "UnInstall-rabbitmq"

Write-Host "Install-Server"

Write-Host "UnInstall-Server"

Write-Host "Install-Ingester"

Write-Host "UnInstall-Ingester"

Write-Host "Install-Grazer"

Write-Host "UnInstall-Grazer"

Write-Host "Install-Monitor"

Write-Host "UnInstall-Monitor"

Write-Host "Install-Alarm"

Write-Host "UnInstall-Alarm"

Write-Host "Install-prereqs"

Write-Host "Install-Apache24-Reverse-Proxy-Server"

Write-Host "UnInstall-Apache24-Reverse-Proxy-Server"

Write-Host "Configure-Configs"

Write-Host "Get-version"

Write-Host "WW5-Ports"

Write-Host "postgres-backup"

Write-Host "Postgres-Housekeeping"

Write-Host "add-user"

Write-host "Install-postgres93"

Write-host "UnInstall-postgres93"






}





function Install-MSSQL2017() {

  Param 
    ( 

    [Parameter(Mandatory=$true)] 
        [string]$binarydriveletter = "C",

        [Parameter(Mandatory=$true)] 
        [string]$Logsdriveletter= "C",

        [Parameter(Mandatory=$true)] 
        [string]$Datadriveletter= "C"

        #Fails If changed 
        #[Parameter(Mandatory=$false)] 
        #[string]$Instancename="MSSQLSERVER"  
        ) 





#InstallConfigFile forSql
$arg1 = "/CONFIGURATIONFILE=`"C:\InstallFiles\MSSQL2017\ConfigurationFile.ini`"" 

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
#SQL Server 2017 Database Engine Services



#Binary
#Instance Dir
#INSTANCEDIR="C:\Program Files\Microsoft SQL Server"

$OriginalINSTANCEDIR="INSTANCEDIR=`"C:\Program Files\Microsoft SQL Server" 
$NewINSTANCEDIR="INSTANCEDIR=`"$binarydriveletter`:\Program Files\Microsoft SQL Server"

#Install Share Dir
#INSTALLSHAREDDIR=INSTALLSHAREDDIR="C:\Program Files\Microsoft SQL Server"

$OriginalINSTALLSHAREDDIR="INSTALLSHAREDDIR=`"C:\Program Files\Microsoft SQL Server"
$NewINSTALLSHAREDDIR="INSTALLSHAREDDIR=`"$binarydriveletter`:\Program Files\Microsoft SQL Server"


#Install Share WOW Dir
#INSTALLSHAREDWOWDIR="C:\Program Files (x86)\Microsoft SQL Server"
$OriginalINSTALLSHAREDWOWDIR="INSTALLSHAREDWOWDIR=`"C:\Program Files (x86)\Microsoft SQL Server"
$NewINSTALLSHAREDWOWDIR="INSTALLSHAREDWOWDIR=`"$binarydriveletter`:\Program Files (x86)\Microsoft SQL Server"



#Data
#SQLBACKUPDIR
#SQLBACKUPDIR="D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Backup"
$OriginalSQLBACKUPDIR="SQLBACKUPDIR=`"D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Backup"
$nEWSQLBACKUPDIR="SQLBACKUPDIR=`"$Datadriveletter`:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Backup"

#SQLTEMPDBDIR
#SQLTEMPDBDIR="D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$ORIGINALSQLTEMPDBDIR="SQLTEMPDBDIR=`"D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$NewSQLTEMPDBDIR="SQLTEMPDBDIR=`"$Datadriveletter`:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"


#sqluserdir
#SQLUSERDBDIR="D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$originalsqluserdbdir="SQLUSERDBDIR=`"D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$newsqluserdbdir="SQLUSERDBDIR=`"$Datadriveletter`:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"



#Logs


#SQLUSERDBLOGDIR="D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$OriginalSQLUSERDBLOGDIR="SQLUSERDBLOGDIR=`"D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$NewSQLUSERDBLOGDIR="SQLUSERDBLOGDIR=`"$Logsdriveletter`:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"


#SQLTEMPDBLOGDIR="D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$originalSQLTEMPDBLOGDIR="SQLTEMPDBLOGDIR=`"D:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"
$newSQLTEMPDBLOGDIR="SQLTEMPDBLOGDIR=`"$Logsdriveletter`:\Program Files\Microsoft SQL Server\MSSQL14.MSSQLSERVER\MSSQL\Data"


#InstanceName
#INSTANCENAME="MSSQLSERVER"
$originalINSTANCENAME="INSTANCENAME=`"MSSQLSERVER"
$NewINSTANCENAME="INSTANCENAME=`"$Instancename"

$InstanceID=$Instancename

#InstanceID
#INSTANCEID="MSSQLSERVER"
$OriginalINSTANCEID="INSTANCEID=`"MSSQLSERVER"
$NewINSTANCEID="INSTANCEID=`"$Instancename"


$original_file = "C:\InstallFiles\MSSQL2017\OriginalConfigurationFile.ini"
$destination_file =  "C:\InstallFiles\MSSQL2017\ConfigurationFile.ini"
$path=":\"
If (!(test-path $original_file)) {Write-Log -Path $Logsfile -level Error -message  "$original_file Not found" ; break }
If (!(test-path $binarydriveletter$path )) {Write-Log -Path $Logsfile  -level Error -message  "Drive Letter $binarydriveletter Not found for Binary Drive" ; break }
If (!(test-path $Logsdriveletter$path )) {Write-Log -Path  $Logsfile  -level Error -message  "Drive Letter $Logsdriveletter  Not found for Logs Drive" ;break }
If (!(test-path $Datadriveletter$path )) {Write-Log -Path $Logsfile  -level Error -message  "Drive Letter $Datadriveletter Not found for DataDrive" ; break }

(Get-Content $original_file) | ForEach-Object {
    $_.replace($OriginalINSTANCEDIR, $NewINSTANCEDIR).replace($OriginalINSTALLSHAREDDIR, $NewINSTALLSHAREDDIR).replace($OriginalINSTALLSHAREDWOWDIR, $NewINSTALLSHAREDWOWDIR).replace($OriginalSQLBACKUPDIR, $nEWSQLBACKUPDIR).replace($ORIGINALSQLTEMPDBDIR, $NewSQLTEMPDBDIR).replace($originalsqluserdbdir, $newsqluserdbdir).replace($OriginalSQLUSERDBLOGDIR, $NewSQLUSERDBLOGDIR).replace($originalSQLTEMPDBLOGDIR, $newSQLTEMPDBLOGDIR).replace($originalINSTANCENAME, $NewINSTANCENAME).replace($OriginalINSTANCEID, $NewINSTANCEID)
    if (test-path $destination_file) {Write-Log -Path $Logsfile -level Info -message  "SQL Configuration file $Destination has been written " }
    else {Write-Log -Path $Logsfile -level Error -message  "No SQL Configuration file $Destination has been written!! " }
 } | Set-Content $destination_file





if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server 2017 Database Engine Services”}) {write-log -Path $Logsfile  -level Warn -message "SQL Server 2017 Database Engine Services already Installed" ; } 
    else { 
        if (!(test-path -path $global:InstallFolder\MSSQL2017\* -Filter *setup.exe*)) {Write-Log -Path $Logsfile -level Error -message  "MSSQL2017 Setup.exe not found in  $global:InstallFolder directory"; }
            $MSSQL2017InstallFile = (Get-ItemProperty -Path $global:InstallFolder\MSSQL2017\* -Filter *setup.exe).FullName
            $process = Start-Process -Filepath $MSSQL2017InstallFile -ArgumentList  "$arg1" -wait
            If($process.Exitcode -ne 0)
            {
            write-log -Path $Logsfile  -level Error -message "Error encountered , Your Error Code is $($process.ExitCode) , Try a reboot or check error code here 'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376931(v=vs.85).aspx'"
            #Throw "Errorlevel $($process.ExitCode)"
            }
            if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server 2017 Database Engine Services”})) {write-log -Path $Logsfile  -level Error -message "Failed to install SQL Server 2017 Database Engine Services , see C:\Program Files\Microsoft SQL Server\140\Setup Bootstrap\Log\Summary.txt for further details " ; }
            if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server 2017 Database Engine Services”}) {write-log -Path $Logsfile  -level INFO -message "SQL Server 2017 Database Engine Services , Instaled Successfully " ; }
 
            }





}


function UnInstall-MSSQL2017() {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
#Uninstall Reg Keys For SQL
$key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft SQL Server SQL2017"
$value = "UninstallString"
#SQL Server 2017 Database Engine Services
if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server 2017 Database Engine Services”})) {write-log -Path $Logsfile  -level Error -message "SQL Server 2017 Database Engine Services is not Installed" ; } 
else{
        if (!(test-path $key  )) {Write-Log -Path $Logsfile  -level Error -message  "SQL Server 2017 Uninstall Key not found at $key" ; break}
        $SQLUninstallFile = (Get-ItemProperty -Path $key -Name $value).$value
        #Don't Even Ask ....
        $SQLUninstallFile1 = $SQLUninstallFile -replace '"', ""
        if (!(test-path $SQLUninstallFile1 )) {Write-Log -Path $Logsfile  -level Error -message  "Uninstall Key from registry gave us this address $SQLUninstallFile , but file was not found " ; break}
        
        $process = Start-process -FilePath $SQLUninstallFile1  -Wait 
        #Call Install String and throw Error code if encountered.
            If($process.Exitcode -ne 0)
            {
            write-log -Path $Logsfile  -level Error -message "Error encountered , Your Error Code is $($process.ExitCode) , Try a reboot or check error code here 'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376931(v=vs.85).aspx'"
            #Throw "Errorlevel $($process.ExitCode)"
            }
        if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server 2017 Database Engine Services”})) {Write-Log -Path $Logsfile  -level info -message  "SQL Server 2017 Database Engine Service has been uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "SQL Server 2017 Database Engine Service Failed to Uninstall"}

}

}

function Install-SMSS17.5() {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
$ARG1 = '/install /quiet /passive /norestart'
if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  'SQL Server Management Studio'}) {write-log -Path $Logsfile  -level Error -message "SQL Server Management Studio Already Installed" ; } 
    else { 
        if (!(test-path -path $global:InstallFolder\MSSQL2017\* -Filter *SSMS-Setup-ENU.exe*)) {Write-Log -Path $Logsfile -level Error -message  "SSMS-Setup-ENU.exe not found in  $global:InstallFolder\\MSSQL2017\ directory"; }
            $SMSS17InstallFile = (Get-ItemProperty -Path $global:InstallFolder\MSSQL2017\* -Filter *SSMS-Setup-ENU.exe* ).FullName
            #Install String
            $process = Start-Process -FilePath "$SMSS17InstallFile" -ArgumentList "$ARG1" -PassThru -Wait  -NoNewWindow
            #Call Install String and throw Error code if encountered.
            If($process.Exitcode -ne 0)
            {
            write-log -Path $Logsfile  -level Error -message "Error encountered , Your Error Code is $($process.ExitCode) , Try a reboot or check error code here 'https://msdn.microsoft.com/en-us/library/windows/desktop/aa376931(v=vs.85).aspx'"
            #Throw "Errorlevel $($process.ExitCode)"
            }

            if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server Management Studio”})) {write-log -Path $Logsfile  -level Error -message "Failed to install SQL Server Management Studio " ; }
            if (Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  “SQL Server Management Studio”}) {write-log -Path $Logsfile  -level INFO -message "SQL Server Management Studio , Instaled Successfully " ; }
 
            }

}


function UnInstall-SMSS17.5() {
$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Installer Started for $Functionname "
#SQL Server Management Studio 
if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  'SQL Server Management Studio'})) {write-log -Path $Logsfile  -level Error -message "SQL Server Management Studio not found" ; } 
else{
$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -eq "SQL Server Management Studio" }
    $($app.Uninstall()) >$null
        if (!(Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -eq  'SQL Server Management Studio'})) {Write-Log -Path $Logsfile  -level info -message  "SQL Server 2017 Database Engine Service has been uninstalled" }
        else {Write-Log -Path $Logsfile  -level Error -message  "SQL Server Management Studio  Failed to Uninstall"}

}

}




