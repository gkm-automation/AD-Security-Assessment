
# First Clear any variables
Remove-Variable * -ErrorAction SilentlyContinue


# Start script execution time calculation
$ScrptStartTime = (Get-Date).ToString('dd-MM-yyyy hh:mm:ss')
$sw = [Diagnostics.Stopwatch]::StartNew()

# Get Script Directory
$Scriptpath = $($MyInvocation.MyCommand.Path)
$Dir = $(Split-Path $Scriptpath);

# Report
$runntime= (get-date -format dd_MM_yyyy-HH_mm_ss)-as [string]
$HealthReport = "$dir\Reports" + "$runntime" + ".htm"

# Logfile 
$Logfile = "$dir\Log" + "$runntime" + ".log"

#---------------------------------------------------------------------------------------------------------------------------------------------
# Functions Section
#---------------------------------------------------------------------------------------------------------------------------------------------
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Warning','Error')]
        [string]$Severity = 'Information'
    )
    $LogContent = (Get-Date -f g)+" " + $Severity +"  "+$Message
    Add-Content -Path $logFile -Value $LogContent -PassThru | Write-Host

 }

 function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        �^\[(.+)\]� # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        �^(;.*)$� # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = �Comment� + $CommentCount
            $ini[$section][$name] = $value
        }
        �(.+?)\s*=(.*)� # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

#import AD Module

try {
 Import-Module ActiveDirectory   
}
catch [System.Management.Automation.ParameterBindingException] {
     Write-Log -Message "Failed Importing Active Directory Module..!" -Severity Error
      Break;
}

#Import Configuration Params

$params = Get-IniContent -filePath "$dir\Config.ini"

# E-mail report details
$SendEmail     = $params.SMTPSettings.SendEmail.Trim()
$emailFrom     = $params.SMTPSettings.EmailFrom.Trim()
$emailTo       = $params.SMTPSettings.EmailTo.Trim()
$smtpServer    = $params.SMTPSettings.SmtpServer.Trim()
$emailSubject  = $params.SMTPSettings.EmailSubject.Trim()


$DCtoConnect = $params.Config.ConnectorDC.Trim()
[string]$date = Get-Date

$DCList = @()


#---------------------------------------------------------------------------------------------------------------------------
# Setting the header for the Report
#---------------------------------------------------------------------------------------------------------------------------

[DateTime]$DisplayDate = ((get-date).ToUniversalTime())

$header = Get-Content -Raw -Path "$dir\header.html"
Add-Content $HealthReport $header



#---------------------------------------------------------------------------------------------------------------------------
# Domain INfo
#---------------------------------------------------------------------------------------------------------------------------

try 
{ 
      $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()  
} 
catch 
{ 
      Write-Log "Cannot connect to current Domain."
      Break;
} 
 
$Domain.DomainControllers | ForEach-Object {
$DCList += $_.Name
}

if(!$DCList)
{
   Write-Log "No Domain Controller found. Run this solution on AD server. Please try again."
   Break;
}

                                   
Write-Log "List of Domain Controllers Discovered"

# List out all machines discovered in Log File and Console
foreach ($D in $DCList) 
{
Write-Log "$D"
}

Add-Content $HealthReport $dataRow

# Check if any domain controllers left
if($DCList.Count -eq 0) {
    Write-Log -Message "As no machines left script won't continue further" -Severity Error
    Break
}

# Start Container Div and Sub container div
$dataRow = "<div id=container><div id=portsubcontainer>"
$dataRow += "<table border=1px>
<caption><h2><a name='Domain Info'>Domain Info</h2></caption>"

$forestinfo = Get-ADForest -Server $DCtoConnect
$domaininfo = Get-ADDomain -Server $DCtoConnect

$dataRow += "<tr>
<td class=bold_class>ForestName</td>
<td >$($($forestinfo.Name).ToUpper())</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>DomainName</td>
<td >$($($domaininfo.Name).ToUpper())</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>ForestMode(FFL)</td>
<td >$($forestinfo.ForestMode)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>DomainMode(DFL)</td>
<td >$($domaininfo.DomainMode)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>SchemaMaster</td>
<td >$($forestinfo.SchemaMaster)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>DomainNamingMaster</td>
<td >$($forestinfo.DomainNamingMaster)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>PDCEmulator</td>
<td >$($domaininfo.PDCEmulator)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>RIDMaster</td>
<td >$($domaininfo.DomainMode)</td>
</tr>"
$dataRow += "<tr>
<td class=bold_class>InfrastructureMaster</td>
<td >$($domaininfo.InfrastructureMaster)</td>
</tr>"

Add-Content $HealthReport $dataRow
Add-Content $HealthReport "</table></div>" # End Sub Container Div

#---------------------------------------------------------------------------
# Domain Users Validation
#---------------------------------------------------------------------------
Write-Log -Message "Performing Domain Users Validation..............."
# Start Sub Container
$DomainUsers= "<Div id=DomainUserssubcontainer><table border=1px>
   <caption><h2><a name='DomainUsers'>Domain Users</h2></caption>
        "
Add-Content $HealthReport $DomainUsers

## Get Domain User Information
$LastLoggedOnDate = $(Get-Date) - $(New-TimeSpan -days $params.Config.UserLogonAge)  
$PasswordStaleDate = $(Get-Date) - $(New-TimeSpan -days $params.Config.UserPasswordAge)
$ADLimitedProperties = @("Name","Enabled","SAMAccountname","DisplayName","Enabled","LastLogonDate","PasswordLastSet","PasswordNeverExpires","PasswordNotRequired","PasswordExpired","SmartcardLogonRequired","AccountExpirationDate","AdminCount","Created","Modified","LastBadPasswordAttempt","badpwdcount","mail","CanonicalName","DistinguishedName","ServicePrincipalName","SIDHistory","PrimaryGroupID","UserAccountControl")

[array]$DomainUsers = Get-ADUser -Filter * -Property $ADLimitedProperties -Server $DCtoConnect 
[array]$DomainEnabledUsers = $DomainUsers | Where {$_.Enabled -eq $True }
[array]$DomainDisabledUsers = $DomainUsers | Where {$_.Enabled -eq $false }
[array]$DomainEnabledInactiveUsers = $DomainEnabledUsers | Where { ($_.LastLogonDate -le $LastLoggedOnDate) -AND ($_.PasswordLastSet -le $PasswordStaleDate) }

[array]$DomainUsersWithReversibleEncryptionPasswordArray = $DomainUsers | Where { $_.UserAccountControl -band 0x0080 } 
[array]$DomainUserPasswordNotRequiredArray = $DomainUsers | Where {$_.PasswordNotRequired -eq $True}
[array]$DomainUserPasswordNeverExpiresArray = $DomainUsers | Where {$_.PasswordNeverExpires -eq $True}
[array]$DomainKerberosDESUsersArray = $DomainUsers | Where { $_.UserAccountControl -band 0x200000 }
[array]$DomainUserDoesNotRequirePreAuthArray = $DomainUsers | Where {$_.DoesNotRequirePreAuth -eq $True}
[array]$DomainUsersWithSIDHistoryArray = $DomainUsers | Where {$_.SIDHistory -like "*"}

$domainusersrow = "<thead><tbody><tr>
<td class=bold_class>Total Users</td>
<td width='40%' style= 'text-align: center'>$($DomainUsers.Count)</td>
</tr>"
$domainusersrow += "<tr>
<td class=bold_class>Enabled Users</td>
<td width='40%' style= 'text-align: center'>$($DomainEnabledUsers.Count)</td>
</tr>"
$domainusersrow += "<tr>
<td class=bold_class>Disabled Users</td>
<td width='40%' style= 'text-align: center'>$($DomainDisabledUsers.Count)</td>
</tr>"
$domainusersrow += "<tr>
<td class=bold_class>Inactive Users</td>
<td width='40%' style= 'text-align: center'>$($DomainEnabledInactiveUsers.Count)</td>
</tr>"
$domainusersrow += "<tr>
<td class=bold_class>Users With Password Never Expires</td>
<td width='40%' style= 'text-align: center'>$($DomainUserPasswordNeverExpiresArray.Count)</td>
</tr>"

$domainusersrow += "<tr>
<td class=bold_class>Users With SID History</td>
<td width='40%' style= 'text-align: center'>$($DomainUsersWithSIDHistoryArray.Count)</td>
</tr>"
If($($DomainUsersWithReversibleEncryptionPasswordArray.Count) -gt 0){
    $temp = @()
    $DomainUsersWithReversibleEncryptionPasswordArray | ForEach-Object { $temp = $temp + $_.SamAccountName + "<br>" }
    $domainusersrow += "<tr>
    <td class=bold_class>Users With ReversibleEncryptionPasswordArray</td>
    <td class=failed width='40%' style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $temp +"')"">$($DomainUsersWithReversibleEncryptionPasswordArray.Count)</a></td>
    </tr>" 
}
else{
    $domainusersrow += "<tr>
    <td class=bold_class>Users With ReversibleEncryptionPasswordArray</td>
    <td class=passed width='40%' style= 'text-align: center'>$($DomainUsersWithReversibleEncryptionPasswordArray.Count)</td>
    </tr>" 
}
If($($DomainUserPasswordNotRequiredArray.Count) -gt 0){
    $temp = @()
    $DomainUserPasswordNotRequiredArray | ForEach-Object { $temp = $temp + $_.SamAccountName + "<br>" }
    $domainusersrow += "<tr>
    <td class=bold_class>Users With Password Not Required</td>
    <td class=failed width='40%' style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $temp +"')"">$($DomainUserPasswordNotRequiredArray.Count)</a></td>
    </tr>" 
}
else{
    $domainusersrow += "<tr>
    <td class=bold_class>Users With Password Not Required</td>
    <td class=passed width='40%' style= 'text-align: center'>$($DomainUserPasswordNotRequiredArray.Count)</td>
    </tr>" 
}

If($($DomainKerberosDESUsersArray.Count) -gt 0){
    $temp = @()
    $DomainKerberosDESUsersArray | ForEach-Object { $temp = $temp + $_.SamAccountName + "<br>" }
    $domainusersrow += "<tr>
    <td class=bold_class>Users With Kerberos DES</td>
    <td class=failed width='40%' style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $temp +"')"">$($DomainKerberosDESUsersArray.Count)</a></td>
    </tr>" 
}
else{
    $domainusersrow += "<tr>
    <td class=bold_class>Users With Kerberos DES</td>
    <td class=passed width='40%' style= 'text-align: center'>$($DomainKerberosDESUsersArray.Count)</td>
    </tr>" 
}

If($($DomainUserDoesNotRequirePreAuthArray.Count) -gt 0){
    $temp = @()
    $DomainUserDoesNotRequirePreAuthArray | ForEach-Object { $temp = $temp + $_.SamAccountName + "<br>" }
    $domainusersrow += "<tr>
    <td class=bold_class>Users That Do Not Require Kerberos Pre-Authentication</td>
    <td class=failed width='40%' style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $temp +"')"">$($DomainUserDoesNotRequirePreAuthArray.Count)</a></td>
    </tr>" 
}
else{
    $domainusersrow += "<tr>
    <td class=bold_class>Users That Do Not Require Kerberos Pre-Authentication</td>
    <td class=passed width='40%' style= 'text-align: center'>$($DomainUserDoesNotRequirePreAuthArray.Count)</td>
    </tr>" 
}

Add-Content $HealthReport $domainusersrow

Add-Content $HealthReport "</tbody></table></div></div>" # End Sub Container Div and Container Div
#-----------------------
# Domain Password Policy 
#-----------------------
Write-Log -Message "Determining Domain Password Policy........... "
#Start Container and Sub Container Div
$Pwdpoly = "<div id=container><div id=pwdplysubcontainer><table border=1px>
            <caption><h2><a name='Pwd Policy'>Domain Password Policy</h2></caption>
            "
Add-Content $HealthReport $Pwdpoly

[array]$DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy -Server $DCtoConnect
$props = @("ComplexityEnabled","DistinguishedName","LockoutDuration","LockoutObservationWindow","LockoutThreshold","MaxPasswordAge","MinPasswordAge","MinPasswordLength","PasswordHistoryCount","ReversibleEncryptionEnabled")
foreach($item in $props){
    $flag= 'passed'
    If(($item -eq 'ComplexityEnabled') -and ($DomainPasswordPolicy.ComplexityEnabled -ne 'True')) { $flag = "failed" }
    If(($item -eq 'LockoutDuration') -and $DomainPasswordPolicy.LockoutDuration -lt 15) { $flag = "failed" }
    If(($item -eq 'MaxPasswordAge') -and $DomainPasswordPolicy.MaxPasswordAge -gt 60) { $flag = "failed" }
    If(($item -eq 'MinPasswordAge') -and $DomainPasswordPolicy.MinPasswordAge -lt 1) { $flag = "failed" }
    If(($item -eq 'PasswordHistoryCount') -and $DomainPasswordPolicy.PasswordHistoryCount -le '24') { $flag = "failed" }
    If(($item -eq 'ReversibleEncryptionEnabled') -and $DomainPasswordPolicy.ReversibleEncryptionEnabled -eq 'True') { $flag = "failed" }
    If(($item -eq 'MinPasswordLength') -and $DomainPasswordPolicy.MinPasswordLength -le 14) { $flag = "failed" }
    If(($item -eq 'LockoutDuration') -and $DomainPasswordPolicy.LockoutDuration -le 15) { $flag = "failed" }
    If(($item -eq 'LockoutThreshold') -and ($DomainPasswordPolicy.LockoutThreshold -gt 10 -or $DomainPasswordPolicy.LockoutThreshold -eq 0)) { $flag = "failed" }
    If(($item -eq 'LockoutObservationWindow') -and $DomainPasswordPolicy.LockoutObservationWindow -le 15) { $flag = "failed" }

    $Pwdpolyrow += "<tr>
    <td class=bold_class>$item</td>
    <td class=$flag width='40%' style= 'text-align: center'>$($DomainPasswordPolicy.$item)</td>
    </tr>"

}
Add-Content $HealthReport $Pwdpolyrow

Add-Content $HealthReport "</table></Div>" #End Sub Container

#---------------------------------------------------------------------------------------------------------------------------------------------
# Tombstone and Backup Information
#---------------------------------------------------------------------------------------------------------------------------------------------
Write-Log -Message "Checking Tombstone and Backup Information........"
# Start Sub Container
$tsbkp = "<Div id=TLBkbsubcontainer><table border=1px>
   <caption><h2><a name='tsbkp'>Tombstone & Partitions Backup</h2></caption>
         "
Add-Content $HealthReport $tsbkp

$ADRootDSE = get-adrootdse  -Server $DCtoConnect
$ADConfigurationNamingContext = $ADRootDSE.configurationNamingContext
$TombstoneObjectInfo = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$ADConfigurationNamingContext" `
-Partition "$ADConfigurationNamingContext" -Properties * 
[int]$TombstoneLifetime = $TombstoneObjectInfo.tombstoneLifetime
IF ($TombstoneLifetime -eq 0) { $TombstoneLifetime = 60 }

$tsbkprow += "<tr>
<td class=bold_class>TombstoneLifetime</td>
<td width='30%' style= 'text-align: center'>$TombstoneLifetime</td>
</tr>"

[string[]]$Partitions = (Get-ADRootDSE -Server $DCtoConnect).namingContexts
$contextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain
$context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext($contextType,$($domaininfo.DNSRoot))
$domainController = [System.DirectoryServices.ActiveDirectory.DomainController]::findOne($context)
ForEach($partition in $partitions)
{
   $domainControllerMetadata = $domainController.GetReplicationMetadata($partition)
   $dsaSignature = $domainControllerMetadata.Item(�dsaSignature�)
   Write-Log �$partition was backed up $($dsaSignature.LastOriginatingChangeTime.DateTime)"
    $tsbkprow += "<tr>
    <td class=bold_class>Last backup of '$partition'</td>
    <td width='30%' style= 'text-align: center'>$($dsaSignature.LastOriginatingChangeTime.ToShortDateString())</td>
    </tr>"
}

Add-Content $HealthReport $tsbkprow

Add-Content $HealthReport "</table></Div></Div>" # End Sub Container and Container Div
#---------------------------------------------------------------------------------------------------------------------------------------------
# Kerberos delegation Info
#---------------------------------------------------------------------------------------------------------------------------------------------
Write-Log -Message "Checking Kerberos delegation Info........"
# Start Sub Container
$krbtgtdel = "<div id=container><Div id=delegationsubcontainer><table border=1px>
   <caption><h2><a name='krbtgtdel'>Kerberos Delegation (Unconstrained)</h2></caption>
         <thead>
		<th>ObjectClass</th>
		<th>Count</th>
        </thead>
	       "
Add-Content $HealthReport $krbtgtdel

## Identify Accounts with Kerberos Delegation
$KerberosDelegationArray = @()
[array]$KerberosDelegationObjects =  Get-ADObject -filter { (UserAccountControl -BAND 0x0080000) -AND (PrimaryGroupID -ne '516') -AND (PrimaryGroupID -ne '521') } -Server $DCtoConnect -prop Name,ObjectClass,PrimaryGroupID,UserAccountControl,ServicePrincipalName

ForEach ($KerberosDelegationObjectItem in $KerberosDelegationObjects)
 {
    IF ($KerberosDelegationObjectItem.UserAccountControl -BAND 0x0080000)
     { $KerberosDelegationServices = 'All Services' ; $KerberosType = 'Unconstrained' }
    ELSE 
     { $KerberosDelegationServices = 'Specific Services' ; $KerberosType = 'Constrained' } 
     $KerberosDelegationObjectItem | Add-Member -MemberType NoteProperty -Name KerberosDelegationServices -Value $KerberosDelegationServices -Force
     [array]$KerberosDelegationArray += $KerberosDelegationObjectItem
 }

$Requiredpros = $KerberosDelegationArray | Select Name,ObjectClass
$Groupedresult = $Requiredpros |  Group ObjectClass -AsHashTable

$Groupedresult.Keys | ForEach-Object {
    $objs = ""
    $($Groupedresult.$PSItem.Name) | foreach { $objs = $objs + $_ + "<br>" }
    $krbtgtdelrow += "<tr>
    <td class=bold_class>$($PSItem)</td>
    <td class=failed style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $objs +"')"">$($Groupedresult.$PSItem.Name.count)</a></td>
    </tr>"
}

Add-Content $HealthReport $krbtgtdelrow

Add-Content $HealthReport "</table></Div>" # End Sub Container 

#---------------------------------------------------------------------------------------------------------------------------------------------
# Scan SYSVOL for Group Policy Preference Passwords
#---------------------------------------------------------------------------------------------------------------------------------------------
Write-Log -Message "Scan SYSVOL for Group Policy Preference Passwords......."
# Start Sub Container
$gpppwd = "<Div id=gpppwdsubcontainer><table border=1px>
          <caption><h2><a name='krbtgtdel'>Scan SYSVOL for Group Policy Preference Passwords</h2></caption>
	       "
Add-Content $HealthReport $gpppwd

$domainname = ($domaininfo.DistinguishedName.Replace("DC=","")).replace(",",".")
$DomainSYSVOLShareScan = "\\$domainname\SYSVOL\$domainname\Policies\"
[int]$Count = 0
$Passfoundfiles = ""
$flag = "passed"
Get-ChildItem $DomainSYSVOLShareScan -Filter *.xml -Recurse |  % {
If(Select-String -Path $_.FullName -Pattern "Cpassword"){ $Passfoundfiles += $_.FullName + "</br>" ; $Count += 1; $flag= "failed" }
}
    $gpppwdrow += "<tr>
    <td class=bold_class>Items Found</td>
    <td class=$flag style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $Passfoundfiles +"')"">$Count</a></td>
    </tr>"

Add-Content $HealthReport $gpppwdrow

Add-Content $HealthReport "</table></Div></Div>" # End Sub Container and Container Div
#-------------------------------------
# KRBTGT account info
#-------------------------------------
Write-Log -Message "Checking KRBTGT account info........"
$krbtgt = "<div id=container><div id=krbtgtcontainer><table border=1px>
   <caption><h2><a name='krbtgt'>KRBTGT Account Info</h2></caption>
         <thead>
		<th>DistinguishedName</th>
		<th>Enabled</th>
        <th>msds-keyversionnumber</th>
        <th>PasswordLastSet</th>
        <th>Created</th>
        </thead>
	        <tr>"

Add-Content $HealthReport $krbtgt

$DomainKRBTGTAccount = Get-ADUser 'krbtgt' -Server $DCtoConnect -Properties 'msds-keyversionnumber',Created,PasswordLastSet

If($(New-TimeSpan -Start ($DomainKRBTGTAccount.PasswordLastSet) -End $(Get-Date)).Days -gt 180) { $flag = "failed"
}
else { $flag = "passed" }

$SelectedPros = @("DistinguishedName","Enabled","msds-keyversionnumber","PasswordLastSet","Created")

$SelectedPros | % {

$krbtgtrow += "
    <td class=$flag style= 'text-align: center'>$($DomainKRBTGTAccount.$PSItem)</td>" 
 }

 Add-Content $HealthReport $krbtgtrow
Add-Content $HealthReport "</tr></table></div></div>"
#-----------------------
# Privileged AD Group Report
#-----------------------
Write-Log -Message "Performing Privileged AD Group Report......."
#Start Container and Sub Container Div
$group = "<div id=container><div id=groupsubcontainer><table border=1px>
            <caption><h2>Privileged AD Group Info</h2></caption>
            <thead>
		<th>Privileged Group Name</th>
		<th>Members Count</th>
        </thead>
            "
Add-Content $HealthReport $group
$ADPrivGroupArray = @(
 'Administrators',
 'Domain Admins',
 'Enterprise Admins',
 'Schema Admins',
 'Account Operators',
 'Server Operators',
 'Group Policy Creator Owners',
 'DNSAdmins',
 'Enterprise Key Admins',
 'Exchange Domain Servers',
 'Exchange Enterprise Servers',
 'Exchange Admins',
 'Organization Management',
 'Exchange Windows Permissions'
)
foreach($group in $ADPrivGroupArray){
    try
    {
    $GrpProps = Get-ADGroupMember -Identity $group -Recursive -Server $DCtoConnect -ErrorAction SilentlyContinue | select SamAccountName,distinguishedName
    $tempobj = ""
        $GrpProps | % {
            $tempobj = $tempobj + $_.SamAccountName +"(" + $_.distinguishedName + ")" + "</br>"
        }
        $grouprow += "<tr>
        <td class=bold_class>$group</td>   
        <td style= 'text-align: center'><a href='javascript:void(0)' onclick=""Powershellparamater('"+ $tempobj +"')"">$($GrpProps.SamAccountName.count)</a></td>
        </tr>"
    }
    catch{
        $grouprow += "<tr>
        <td class=bold_class>$group</td>   
        <td style= 'text-align: center'>NA</td>
        </tr>"
    }
}

Add-Content $HealthReport $grouprow

Add-Content $HealthReport "</table></Div></div>" #End Container

#---------------------------------------------------------------------------------------------------------------------------------------------
# Script Execution Time
#---------------------------------------------------------------------------------------------------------------------------------------------
$myhost = $env:COMPUTERNAME

$ScriptExecutionRow = "<div id=scriptexecutioncontainer><table>
   <caption><h2><a name='Script Execution Time'>Execution Details</h2></caption>
      <th>Start Time</th>
      <th>Stop Time</th>
		<th>Days</th>
      <th>Hours</th>
      <th>Minutes</th>
      <th>Seconds</th>
      <th>Milliseconds</th>
      <th>Script Executed on</th>
	</th>"

# Stop script execution time calculation
$sw.Stop()
$Days = $sw.Elapsed.Days
$Hours = $sw.Elapsed.Hours
$Minutes = $sw.Elapsed.Minutes
$Seconds = $sw.Elapsed.Seconds
$Milliseconds = $sw.Elapsed.Milliseconds
$ScriptStopTime = (Get-Date).ToString('dd-MM-yyyy hh:mm:ss')
$Elapsed = "<tr>
               <td>$ScrptStartTime</td>
               <td>$ScriptStopTime</td>
               <td>$Days</td>
               <td>$Hours</td>
               <td>$Minutes</td>
               <td>$Seconds</td>
               <td>$Milliseconds</td>
               <td>$myhost</td>
               
            </tr>
         "
$ScriptExecutionRow += $Elapsed
Add-Content $HealthReport $ScriptExecutionRow
Add-Content $HealthReport "</table></div>"


#---------------------------------------------------------------------------------------------------------------------------------------------
# Sending Mail
#---------------------------------------------------------------------------------------------------------------------------------------------

if($SendEmail -eq 'Yes' ) {

    # Send ADHealthCheck Report
    if(Test-Path $HealthReport) 
    {
        try {
            $body = "Please find AD Health Check report attached."
            #$port = "25"
           Send-MailMessage -Priority High -Attachments $HealthReport -To $emailTo -From $emailFrom -SmtpServer $smtpServer -Body $Body -Subject $emailSubject -Credential $Credentials -UseSsl -Port 587 -ErrorAction Stop
        } catch {       
            Write-Log 'Error in sending AD Health Check Report'
        }
    }

    
    #Send an ERROR mail if Report is not found 
    if(!(Test-Path $HealthReport)) 
    {

        try {
            $body = "ERROR: NO AD Health Check report"
            $port = "25"
            Send-MailMessage -Priority High -To $emailTo -From $emailFrom -SmtpServer $smtpServer -Body $Body -Subject $emailSubject -Port $port -ErrorAction Stop
        } catch {
            Write-Log 'Unable to send Error mail.'
        }
    }

}
else
{
    Write-Log "As Send Email is NO so report through mail is not being sent. Please find the report in Script directory."
}

