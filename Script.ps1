#DB global variable declaration
$global:dbserver = "EC2AMAZ-2RP601Q"
$global:dbinstance ="EC2AMAZ-2RP601Q\SQLEXPRESS"
$global:dbservice = 'MSSQL$SQLEXPRESS'
$global:databasename = "OfflineEvents"
$global:tablename = 'OfflineEventViewer'
# Set the archive location
$global:Sharedpath = "\\EC2AMAZ-2RP601Q.test.local\SharedPath"
$global:Retentiondays = 0

function Register-BackupScheduler
{
    #check for existing task name
    $checktaskexists = Get-ScheduledTaskInfo -TaskName "Backup_EventLog_Scheduler" -ErrorAction SilentlyContinue
    if($checktaskexists)
    {
        Write-Log -Message "Scheduler with Name Backup_EventLog_Scheduler exists already "
        Exit
    }

    $BasicParameters = @{
        TaskName = "Backup_EventLog_Scheduler"
        }

    $Cred = Get-Credential
    $CredentialParameters = @{
        User     = $cred.UserName
        Password = $cred.GetNetworkCredential().Password  # Uncomment this out if using LogOnType Password or InteractiveOrPassword
    }
    # For Scheduled Tasks running as User
    $PrincipalParameters = @{
        UserID    = $cred.UserName # Domain\User
        RunLevel  = "Highest"      # Highest | Limited
        LogOnType = "Password"     # None | Password | S4U | Interactive | Group | ServiceAccount (Localservice, Networkservice, or System) | InteractiveOrPassword
    }
    # Define 1 trigger
    $TriggerParameters = @{
        At   = "06:00pm"    		    # Set a specific start time
        Daily = $true                    # Set schedule to run only once
    }
    $PowershellCommand = "& {AutoLogSyncScheduler}"
    $ActionParameters = @{
        Argument         =  "-NoLogo -NoProfile -ExecutionPolicy Bypass -command $PowershellCommand"
        Execute          = "PowerShell.exe"
        #WorkingDirectory = "C:\SomeFolder"
    }

    # Define each part of a scheduled task object
    $Principal      = New-ScheduledTaskPrincipal @PrincipalParameters
    $Actions        = New-ScheduledTaskAction @ActionParameters
    $Triggers       = New-ScheduledTaskTrigger @TriggerParameters
    # Create a new scheduled task object
    $UserTask = New-ScheduledTask `
        -Principal $Principal `
        -Action $Actions `
        -Trigger $Triggers 
    Register-ScheduledTask @BasicParameters @CredentialParameters -InputObject $UserTask 
    Write-Log -Message "Schedule task created to run @ 6pm daily "
}#schedulerloop

function AutoLogSyncScheduler
{
#checking Database Connectivity
Test-DatabaseConnection
#Getting Server List from Database
$queryserver = Invoke-Sqlcmd2 -ServerInstance $global:dbinstance -Database $global:databasename -Query " SELECT DISTINCT machinename FROM serverdb ORDER BY Machinename "
        if(!($queryserver)) { 
            Write-Log -Message "No Nodes are registered under Backup System.So Log Syncing not applicable." -Level Warn
            Exit;
        }#if
foreach ($Comp in $($queryserver.machinename)){
    try {
        If (Test-Connection -ComputerName $Comp -Count 2 -Quiet -BufferSize 16){
            If(!(Test-WSMan -ComputerName $Comp -ErrorAction SilentlyContinue )){
                Write-Log -Message "WinRM service is not listening properly on $comp. Please check"  -Level Error
            }
            else{
                Sync-NodeSecurityData($Comp)
            }#ifWMI 
        }#ifping
        else{
            Write-Log -Message "$Server is either offline or does not exisit...aborting Eventlog Sync Progress and same backed up in next Schedule for $comp" -Level Error
        }
    }#try
    catch {
        Write-Log -Message "Encountered Error while processing Request $_.Exception.Message " -Level Error
    }#catch
}#foreach
}

function Sync-NodeSecurityData([string]$Servername)
{       
    try {
        $oldestlog = Get-EventLog -LogName Security -ComputerName $Servername | Sort-Object -Property Timegenerated | Select-Object -First 1
        $loghistory = (Get-Date) - ($oldestlog.Timegenerated)
        # Check the security log
        $Log = Get-WmiObject Win32_NTEventLogFile -Filter "logfilename = 'Security'" -ComputerName $Servername -ErrorAction SilentlyContinue
        Write-Log -Message "Query Security Log from $Servername"
        if($loghistory.days -ge $Retentiondays){
            Write-Log -Message "The security event has $($loghistory.days) days old data.The maximum retention period is $Retentiondays days.Hence exceeded the threshold." 
            $localTemp = "\\$Servername\c$\temp"
            if(!(Test-Path $localTemp)){
                Write-Log -Message "$ArchiveFolder not exists in $Servername.So creating.." 
                New-Item -Path "C:\Temp" -ItemType Directory -Force
            }
        $ArchiveFile = "Security-" + (Get-Date -Format "yyyy-MM-dd@HHmm") + ".evtx"
        $Archivepath = $localTemp+"\"+$ArchiveFile
        $Results = ($Log.BackupEventlog($Archivepath)).ReturnValue
            If($Results -eq 0) {
                # Successful backup of security event log
                $logclear = ($Log.ClearEventlog()).ReturnValue
                    If($logclear-eq 0){
                    Write-Log -Message "The security event log was successfully archived to $Archivepath and cleared."
                    $sourcecopy = "\\"+$Servername+"\c$\temp\"+$ArchiveFile
                    Copy-Item -Path $sourcecopy -Destination $Sharedpath -Force
                    $sharedpathfile = $Sharedpath+"\"+$ArchiveFile
                    $logquery = Get-WinEvent -Path $sharedpathfile -ErrorAction SilentlyContinue| select RecordId,TimeCreated,KeywordsDisplayNames,ProviderName,Id,Message,Machinename
                    $dt = $logquery | Out-DataTable
                    Write-DataTable -ServerInstance $dbinstance -Database $databasename -TableName $tablename -Data $dt
                    Write-Log -Message "Offline Backup of Security Log completed for $Servername " -Level Info
                    $latestlogoffline = Invoke-Sqlcmd2 -ServerInstance $dbinstance -Database $databasename -Query " SELECT  top 1 TimeCreated,machinename FROM OfflineEventViewer where machinename LIKE '%$Servername%' ORDER BY TimeCreated DESC"
                    $serverdbtbcheck = Invoke-Sqlcmd2 -ServerInstance $dbinstance -Database $databasename -Query " SELECT  top 1 machinename FROM serverdb where machinename LIKE '%$Servername%'"
                    $convertsqltime = "{0:yyyy-MM-dd HH:mm:ss:fff}" -f ($latestlogoffline.TimeCreated)
                    $servername_db = $latestlogoffline.Machinename
                    if(!($serverdbtbcheck)){
                        Write-Log -Message "Server $Servername registerted in node info Table"
                        $updatequery =  @"
                        INSERT INTO [dbo].[serverdb](Machinename,Lastupdated,RecentFilename)VALUES ('$servername_db','$convertsqltime','$ArchiveFile')
"@ 
                    }
                    else{
                    Write-Log -Message "Last backup details refreshed in Node info Table"
                    $updatequery =  @"
                    update  [dbo].[serverdb] set LastUpdated='$convertsqltime',RecentFilename='$ArchiveFile' where machinename LIKE '%$Servername%'
"@ 
                    }
                    Invoke-Sqlcmd2 -ServerInstance $dbinstance -Database $databasename -Query $updatequery
                    }
                    else{
                        Write-Log -Message "The security event log was successfully archived to $ArchiveFile and not cleared." }
                    }#elselogclear
            else{
                Write-Log -Message "The security event log could not be archived to $ArchiveFile and was not cleared.  Review and resolve security event log issues on $Servername ASAP!"
            }#elselogback
        }#iflogcheck
        else{
            Write-Log -Message "The security event log has $($loghistory.days) days old Data.The maximum retention period is $Retentiondays days. Hence not exceeded the threshold.So no action was taken."
        }
        }#try
        catch {
                Write-Log -Message "$_.Exception.Message" -Level Error           
        }#catch
}

function Test-DatabaseConnection
{
    Write-Log -Message "DBServer Name:$dbserver || DBService Name: $dbservice ||Database Name: $databasename || Table Name: $tablename"
    $tempfilter = "name= '$dbservice' and state = 'Running'"
    try{
        If((Test-Connection $dbserver -Quiet) -and (Get-WmiObject Win32_service -ComputerName $dbserver -Filter $tempfilter)){ 
            Write-Log -Message "Database Connectivity Successful...!" -Level Info
        }
        else
        { 
            Write-Log -Message "Database Connectivity Unsuccessful...Please verify." -Level Error
            Exit                
        }
    }#try
    catch {
        Write-Log -Message "Database Connectivity Unsuccessful...Please verify $($_.Exception.Message)" -Level Error           
        Exit
    }#catch

}

Function Add-SecurityBackupNode{
    [cmdletbinding()]
      PARAM
	(
        [Parameter(Mandatory=$false,Position=0,HelpMessage='Enter ComputerName')]
        [string[]]$Computername = $env:COMPUTERNAME
	)
#calling database check function
Test-DatabaseConnection
#Convert localhost to HOSTNAME
$server = @()
        $Computername | Foreach { if($_ -eq "localhost") {
            Write-Log -Message "Converting localhost Entry to $env:COMPUTERNAME"
                $server += $env:COMPUTERNAME  
                } 
                else { $server += $_ }
            }
foreach ($Servername in $server)
{
    try{
    If(Test-Connection -ComputerName $Servername -Count 2 -Quiet -BufferSize 16 -ErrorAction Stop) 
    {
        Write-Log -Message "Connection to Node $Servername Succeeded" -Level Info
        Write-Log -Message "Checking WinRM service Status on $Servername" -Level Info
        If(!(Test-WSMan -ComputerName $Servername -ErrorAction SilentlyContinue )) 
		{
            Write-Log -Message "WinRM service not configured in $Servername. If service is running, Use command 'Enable-PSRemoting' on $Servername " -Level Error
		}
		else
		{
            $serverexists = Invoke-Sqlcmd2 -ServerInstance $dbinstance -Database $databasename -Query " Select count(1) from $tablename where Machinename LIKE '%$Servername%' "
            if($serverexists.column1 -ne 0 ) { 
                Write-Log -Message "BackupNode $Servername exists already in database" -Level Warn
                break
            }
            else      
            {
                Sync-NodeSecurityData($Servername)
            }#elselogbackup
        }#elseWMI
    }#ifping
    else{
        Write-Log -Message "Node: $Servername is either offline or unreachable...Aborting Eventlog backup " -Level Info
        }
    }#try
    catch {
        Write-Log -Message "$_.Exception.Message" -Level Error           
    }
} 
    Write-Log -Message "================================================================="
}

function Write-Log 
{ 
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
        [string]$Path="$env:USERPROFILE\ADLogArchivingManagerLog$(Get-Date -Format "yyyy-MM-dd").log", 
         
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

####################### 
function Get-Type 
{ 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
} #Get-Type 
 
function Out-DataTable 
{ 
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = new-object Data.datatable   
        $First = $true  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
} #Out-DataTable

function Write-DataTable 
{ 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$true)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$true)] [string]$TableName, 
    [Parameter(Position=3, Mandatory=$true)] $Data, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=5, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$BatchSize=50000, 
    [Parameter(Position=7, Mandatory=$false)] [Int32]$QueryTimeout=0, 
    [Parameter(Position=8, Mandatory=$false)] [Int32]$ConnectionTimeout=15 
    ) 
     
    $conn=new-object System.Data.SqlClient.SQLConnection 
 
    if ($Username) 
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout } 
    else 
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout } 
 
    $conn.ConnectionString=$ConnectionString 
 
    try 
    { 
        $conn.Open() 
        $bulkCopy = new-object ("Data.SqlClient.SqlBulkCopy") $connectionString 
        $bulkCopy.DestinationTableName = $tableName 
        $bulkCopy.BatchSize = $BatchSize 
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut 
        $bulkCopy.WriteToServer($Data) 
        $conn.Close() 
    } 
    catch 
    { 
        $ex = $_.Exception 
        Write-Error "$ex.Message" 
        continue 
    } 
 
} #Write-DataTable

function Invoke-Sqlcmd2 
{ 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$false)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$false)] [string]$Query, 
    [Parameter(Position=3, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=5, Mandatory=$false)] [Int32]$QueryTimeout=600, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$ConnectionTimeout=15, 
    [Parameter(Position=7, Mandatory=$false)] [ValidateScript({test-path $_})] [string]$InputFile, 
    [Parameter(Position=8, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As="DataRow" 
    ) 
 
    if ($InputFile) 
    { 
        $filePath = $(resolve-path $InputFile).path 
        $Query =  [System.IO.File]::ReadAllText("$filePath") 
    } 
 
    $conn=new-object System.Data.SqlClient.SQLConnection 
      
    if ($Username) 
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout } 
    else 
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout } 
 
    $conn.ConnectionString=$ConnectionString 
     
    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller 
    if ($PSBoundParameters.Verbose) 
    { 
        $conn.FireInfoMessageEventOnUserErrors=$true 
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {Write-Verbose "$($_)"} 
        $conn.add_InfoMessage($handler) 
    } 
     
    $conn.Open() 
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn) 
    $cmd.CommandTimeout=$QueryTimeout 
    $ds=New-Object system.Data.DataSet 
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 
    [void]$da.fill($ds) 
    $conn.Close() 
    switch ($As) 
    { 
        'DataSet'   { Write-Output ($ds) } 
        'DataTable' { Write-Output ($ds.Tables) } 
        'DataRow'   { Write-Output ($ds.Tables[0]) } 
    } 
 
} #Invoke-Sqlcmd2


Export-modulemember -function Add-SecurityBackupNode
Export-modulemember -function AutoLogSyncScheduler
Export-modulemember -function Register-BackupScheduler
