"Test Script Download and Executed from 55"  | out-file  c:\test.txt  -Append
$global:LogsFile = "C:\InstallFiles\WordwatchLogs"
$global:GA52PATH = "\\10.10.10.55\share\Wordwatch\V5\Releases\5.2\installfiles\*"
$global:Installfiles = "C:\installfiles\"


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
        [string]$Path="C:\InstallFiles\WordwatchLogs.txt", 
         
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

function Install-Wordwatch ()  {
    <#
    .DESCRIPTION

     
     .PARAMETER driveletter


    .EXAMPLE
        Install-server -driveletter c


    #>



    Param(
    )


$psdir = "$global:Installfiles\Scripts"
##Copy GA folder down
if (!(test-path $global:Installfiles\scripts)) {Write-log -Path $logsfile -Level info -Message "Installfiles folder not found , copying items" ; copy-item $global:GA52PATH -Recurse $global:Installfiles -verbose    }
Else {Write-log -Path $logsfile -Level info -Message "Installfiles folder already exisits"}


#Dot Source Script folder
Get-ChildItem "${psdir}\*.ps1" | %{.$_} 

#Install sOFTWARE AND pREREQS
#get-version -driveletter c
Install-postgres95 -Driveletter D
Install-erlang 
Install-rabbitmq
install-prereqs
Install-server -Driveletter c
Install-ingester -Driveletter c
Install-grazer -Driveletter c
Install-monitor -Driveletter c
Install-alarm -Driveletter c

#Create Ingester Share
$share = "c:\share"
if (!(test-path $share ) ) { Write-log -Path $logsfile -Level info -Message "Share folder not found creating..."
New-Item “C:\Share" –type directory
net share "Share=c:\share" "/GRANT:Everyone,Full" }
else {Write-log -Path $logsfile -Level info -Message "Share folder already created" }

#ConfigureServices
#Configure-configs -postgres y -driveletter d
Configure-Configs -server y -ingester y -grazer y -retention y -ComplianceHold y -ComplianceManager y -Monitor y -Alarm y -driveletter c
Send-smtp
Grazer-config -driveletter C
Restart-Computer -Force 

}

function SEND-SMTP() {



$ip=(Get-NetIPAddress | ?{ $_.AddressFamily -eq “IPv4” -and !($_.IPAddress -match “169”) -and !($_.IPaddress -match “127”) }).IPAddress
$hostname = $env:computername
$versions = Get-Version-output -driveletter c

Send-MailMessage -SmtpServer 10.10.10.99 -To thersey@businesssystemsuk.com -From devsupport@businesssystemsuk.com -subject "New Wordwatch Host" -Body "

New Wordwatch Host : Details
Ip address         : $IP
Hostnanme          : $Hostname
WordWatch Installed Version 
$versions
 "

}

function Get-Version-output () {
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




if (test-path $Serverpath) { Write-output "Your Version of Server is $ServerVersion"}
Else {Write-output "Server Process not Found"  }

if (test-path $IngesterPath) { Write-output "Your Version of Ingester is $IngesterVersion"}
Else {Write-output "Ingester Process not Found"  }

if (test-path $GrazerPath) {Write-output"Your Version of Grazer is $GrazerVersion" }
Else {Write-output "Grazer Process not Found"  }

if (test-path $MonitorPath) { Write-output "Your Version of Monitor is $Monitorversion"}
Else {Write-output "Monitor Process not Found"  }

if (test-path $AlarmPath) { Write-output "Your Version of Alarm is $AlarmVersion"}
Else {Write-output "Alarm Process not Found"  }
   
}


function Grazer-Config () {


    Param(
     [string]$driveletter='c'
    )

$Functionname = ' {0}.' -f $MyInvocation.MyCommand
write-log -Path $Logsfile  -level Info -message "Fucntion Started for $Functionname "

if (test-path $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" ) {
Write-log -Path $logsfile -Level info -Message "Grazer Config Found Configuring for DEV Red Box"

#Configure Grazer IP
$DefaultGrazerIP = "<add key=`"IP`" value=`"IPADDRESS`" />"
$NewGrazerIP = "<add key=`"IP`" value=`"10.10.10.10`" />"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultGrazerIP , $NewGrazerIP  }) | Set-Content $_ }

#DefaultGrazerAdminCreds
$DefaultGrazerAdminCreds = "01D12C947916DEB0AC01A1A9EC962029"
$NewGrazerAdminCreds = "741F8CE82B5C507AEC06A060FF9E4EC6"

gci $driveletter":\Program Files (x86)\Vocal Recorders\WordWatch\Grazer\RbrGrazer.exe.config" -recurse | ForEach {
  (Get-Content $_ | ForEach {$_ -replace $DefaultGrazerAdminCreds , $NewGrazerAdminCreds  }) | Set-Content $_ }
  }
Else {Write-log -Path $logsfile -Level Warn -Message "Grazer Config Not FounD"}
}

Install-Wordwatch







