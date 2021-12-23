<#
    .SYNOPSIS
    Creates an HTML report describing the On-Premises Exchange environment.
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 2.4 December 2021

    Based on the original 1.6.2 version by Steve Goodman

    Ideas, comments and suggestions to support@granikos.eu
	
    .DESCRIPTION
	
    This script creates an HTML report showing the following information about an Exchange 
    2019, 2016, 2013, 2010, and, to a lesser extent, 2007 and 2003 environment. 
    
    The reports shows the following:
	
    * Report Generation Time
    * Total Servers per Exchange Version (2003 > 2010 or 2007 > 2019)
    * Total Mailboxes per Exchange Version, Office 365, and Organisation
    * Total Roles in the environment
		
    Then, per site:
    * Total Mailboxes per site
    * Internal, External and CAS Array Hostnames
    * Exchange Servers with:
      o Exchange Server Version
      o Service Pack
      o Number of preferred and maximum active databases
      o Update Rollup and rollup version
      o Roles installed on server and mailbox counts
      o OS Version and Service Pack
      
    Then, per Database availability group (Exchange 2010/2013/2016/2019):
    * Total members per DAG
    * Member list
    * Databases, detailing:
      o Mailbox Count and Average Size
      o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
      o Database Size and whitespace
      o Database and log disk free
      o Last Full Backup (Only shown if one or more DAG database has been backed up)
      o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
      o Mailbox server hosting active copy
      o List of mailbox servers hosting copies and number of copies
		
    Finally, per Database (Non DAG DBs/Exchange 2007/Exchange 2003)
    * Databases, detailing:
      o Storage Group (if applicable) and DB name
      o Server hosting database
      o Mailbox Count and Average Size
      o Archive Mailbox Count and Average Size (Only shown if DAG includes Archive Mailboxes)
      o Database Size and whitespace
      o Database and log disk free
      o Last Full Backup (Only shown if one or more DAG database has been backed up)
      o Circular Logging Enabled (Only shown if one or more DAG database has Circular Logging enabled)
		
    This does not detail public folder infrastructure, or examine Exchange 2007/2003 CCR/SCC clusters
    (although it attempts to detect Clustered Exchange 2007/2003 servers, signified by ClusMBX).
	
    IMPORTANT NOTE: The script requires WMI and Remote Registry access to Exchange servers from the server 
    it is run from to determine OS version, Update Rollup, Exchange 2007/2003 cluster and DB size information.
  
    .LINK  
    http://scripts.granikos.eu

    .PARAMETER HTMLReport
    Filename to write HTML Report to
	
    .PARAMETER SendMail
    Send Mail after completion. Set to $True to enable. If enabled, -MailFrom, -MailTo, -MailServer are mandatory
	
    .PARAMETER MailFrom
    Email address to send from. Passed directly to Send-MailMessage as -From
	
    .PARAMETER MailTo
    Email address to send to. Passed directly to Send-MailMessage as -To
	
    .PARAMETER MailServer
    SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer
	  
    .PARAMETER ViewEntireForest
    By default, true. Set the option in Exchange 2007 or 2010 to view all Exchange servers and recipients in the forest.
   
    .PARAMETER ServerFilter
    Use a text based string to filter Exchange Servers by, e.g., NL-* 
    Note the use of the wildcard (*) character to allow for multiple matches.

    .PARAMETER ShowDriveNames
    Include drive names of EDB file path and LOG file folder in database report table
    
    .EXAMPLE
    Generate the HTML report 
    .\Get-ExchangeEnvironmentReport.ps1 -HTMLReport .\report.html

    .EXAMPLE
    Generate am HTML report and send the result as HTML email with attachment to the specified recipient using a dedicated smart host
    .\Get-ExchangeEnvironmentReport.ps1 -HTMReport ExchangeEnvironment.html -SendMail -ViewEntireForet $true -MailFrom roaster@mcsmemail.de -MailTo grillmaster@mcsmemail.de -MailServer relay.mcsmemail.de

    .EXAMPLE
    Generate the HTML report including EDB and LOG drive names
    .\Get-ExchangeEnvironmentReport.ps1 -ShowDriveNames -HTMLReport .\report.html
#>
[CmdletBinding()]
param(
  [parameter(Position=0,Mandatory,HelpMessage='Filename to write HTML report to')][string]$HTMLReport,
  [switch]$SendMail,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = '',
  [bool]$ViewEntireForest=$true,
  [string]$ServerFilter='*',
  [switch]$ShowDriveNames
)

# Warning Limits, adjust as needed
$MinFreeDiskspace = 10 # Mark free space less than this value (%) in red
$MaxDatabaseSize = 250 # Mark database larger than this value (GB) in red

# Default variables
$NotAvailable = 'N/A'

# Set TLS version o TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Sub-Function to Get Database Information. Shorter than expected..
function Get-DatabaseAvailabilityGroupInformation {
  [CmdletBinding()]
  param(
    $DAG
  )

  @{Name = $DAG.Name.ToUpper()
    MemberCount	= $DAG.Servers.Count
    Members = [array]($DAG.Servers | ForEach-Object { $_.Name })
    Databases = @()
  }
}

# Sub-Function to Get Database Information
function Get-DatabaseInformation {
  [CmdletBinding()]
  param(
    $Database,
    $ExchangeEnvironment,
    $Mailboxes,
    $ArchiveMailboxes,
    $E2010
  )
	
  # Circular Logging, Last Full Backup
  if ($Database.CircularLoggingEnabled) { $CircularLoggingEnabled='Yes' } else { $CircularLoggingEnabled = 'No' }
  if ($Database.LastFullBackup) { $LastFullBackup=$Database.LastFullBackup.ToString() } else { $LastFullBackup = 'Not Available' }

  # Drive Letter, GitHub issue #4
  $DriveNameEdb = ''
  try {
    $DriveNameEdb = $Database.EdbFilePath.DriveName
  }
  catch {
    $DriveNameEdb = $NotAvailable
  }

  $DriveNameLog = ''
  try {
    $DriveNameLog = $Database.LogFolderPath.DriveName
  }
  catch {
    $DriveNameLog = $NotAvailable
  }
	
  # Mailbox Average Sizes
  $MailboxStatistics = [array]($ExchangeEnvironment.Servers[$Database.Server.Name].MailboxStatistics | Where-Object {$_.Database -eq $Database.Identity})
  
  if ($MailboxStatistics) {
    [long]$MailboxItemSizeB = 0
    $MailboxStatistics | ForEach-Object{ $MailboxItemSizeB+=$_.TotalItemSizeB }
    [long]$MailboxAverageSize = $MailboxItemSizeB / $MailboxStatistics.Count
  } 
  else {
    $MailboxAverageSize = 0
  }
	
  # Free Disk Space Percentage
  if ($ExchangeEnvironment.Servers[$Database.Server.Name].Disks) {

    foreach ($Disk in $ExchangeEnvironment.Servers[$Database.Server.Name].Disks) {
      if ($Database.EdbFilePath.PathName -like ('{0}*' -f $Disk.Name)) {
        $FreeDatabaseDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
      }
      if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14) {

        if ($Database.LogFolderPath.PathName -like ('{0}*' -f $Disk.Name)) {
          $FreeLogDiskSpace = ($Disk.FreeSpace / $Disk.Capacity) * 100
        }
      } 
      else {
        $StorageGroupDN = $Database.DistinguishedName.Replace(('CN={0},' -f $Database.Name),'')
        $Adsi=[adsi]"LDAP://$($Database.OriginatingServer)/$($StorageGroupDN)"
        if ($Adsi.msExchESEParamLogFilePath -like ('{0}*' -f $Disk.Name)) {
          $FreeLogDiskSpace = $Disk.FreeSpace / $Disk.Capacity * 100
        }
      }
    }
  } 
  else {
    $FreeLogDiskSpace=$null
    $FreeDatabaseDiskSpace=$null
  }
	
  if ($Database.ExchangeVersion.ExchangeBuild.Major -ge 14 -and $E2010) {
    # Exchange 2010 Database Only
    $CopyCount = [int]$Database.Servers.Count

    if ($Database.MasterServerOrAvailabilityGroup.Name -ne $Database.Server.Name) {
      $Copies = [array]($Database.Servers | ForEach-Object { $_.Name })
    } 
    else {
      $Copies = @()
    }
    # Archive Info
    $ArchiveMailboxCount = [int]([array]($ArchiveMailboxes | Where-Object {$_.ArchiveDatabase -eq $Database.Name})).Count

    $ArchiveStatistics = [array]($ArchiveMailboxes | Where-Object {$_.ArchiveDatabase -eq $Database.Name} | Get-MailboxStatistics -Archive )

    if ($ArchiveStatistics) {
      [long]$ArchiveItemSizeB = 0
      $ArchiveStatistics | ForEach-Object{ $ArchiveItemSizeB+=$_.TotalItemSize.Value.ToBytes() }
      [long]$ArchiveAverageSize = $ArchiveItemSizeB / $ArchiveStatistics.Count
    } 
    else {
      $ArchiveAverageSize = 0
    }

    # DB Size / Whitespace Info
    [long]$Size = $Database.DatabaseSize.ToBytes()
    [long]$Whitespace = $Database.AvailableNewMailboxSpace.ToBytes()
    $StorageGroup = $null
		
  } 
  else {
    $ArchiveMailboxCount = 0
    $CopyCount = 0
    $Copies = @()
    # 2003 & 2007, Use WMI (Based on code by Gary Siepser, http://bit.ly/kWWMb3)
    $Size = [long](get-wmiobject -Class cim_datafile -ComputerName $Database.Server.Name -Filter ('name=''' + $Database.edbfilepath.pathname.replace('\','\\') + '''')).filesize
    
    if (!$Size) {
      Write-Warning -Message ('Cannot detect database size via WMI for {0}' -f $Database.Server.Name)
      [long]$Size = 0
      [long]$Whitespace = 0
    } else {
      [long]$MailboxDeletedItemSizeB = 0
      if ($MailboxStatistics) {
        $MailboxStatistics | ForEach-Object{ $MailboxDeletedItemSizeB+=$_.TotalDeletedItemSizeB }
      }

      # Calculate database whitespace
      $Whitespace = $Size - $MailboxItemSizeB - $MailboxDeletedItemSizeB
      if ($Whitespace -lt 0) { $Whitespace = 0 }
    }

    $StorageGroup =$Database.DistinguishedName.Split(',')[1].Replace('CN=','')
  }
	
  @{Name = $Database.Name
    StorageGroup= $StorageGroup
    ActiveOwner	= $Database.Server.Name.ToUpper()
    MailboxCount = [long]([array]($Mailboxes | Where-Object {$_.Database -eq $Database.Identity})).Count
    MailboxAverageSize = $MailboxAverageSize
    ArchiveMailboxCount	= $ArchiveMailboxCount
    ArchiveAverageSize = $ArchiveAverageSize
    CircularLoggingEnabled = $CircularLoggingEnabled
    LastFullBackup = $LastFullBackup
    Size = $Size
    Whitespace = $Whitespace
    Copies = $Copies
    CopyCount = $CopyCount
    FreeLogDiskSpace = $FreeLogDiskSpace
    FreeDatabaseDiskSpace = $FreeDatabaseDiskSpace
    DriveNameEdb = $DriveNameEdb
    DriveNameLog = $DriveNameLog
  }
}

# Sub-Function to get mailbox count per server.
# New in 1.5.2
function Get-ExchangeServerMailboxCount {
  [CmdletBinding()]
  param(
    $Mailboxes,
    $ExchangeServer,
    $Databases
  )
  # The following *should* work, but it doesn't. Apparently, ServerName is not always returned correctly which may be the cause of
  # reports of counts being incorrect
  #([array]($Mailboxes | Where {$_.ServerName -eq $ExchangeServer.Name})).Count
	
  # ..So as a workaround, I'm going to check what databases are assigned to each server and then get the mailbox counts on a per-
  # database basis and return the resulting total. As we already have this information resident in memory it should be cheap, just
  # not as quick.
  $MailboxCount = 0

  foreach ($Database in [array]($Databases | Where-Object {$_.Server -eq $ExchangeServer.Name})) {
    $MailboxCount+=([array]($Mailboxes | Where-Object {$_.Database -eq $Database.Identity})).Count
  }

  $MailboxCount
	
}

# 2021-12-23 Function added to handle empty virtual directory hostname strings (Issue #9)
function Test-vDirHost {
  param(
    $VDirHost
  )

  [string]$Hostname = 'None'

  if ($null -ne $VDirHost) {
    $Hostname = ([string]$VDirHost).Trim()
  }

  $Hostname
}

# Sub-Function to Get Exchange Server information
function Get-ExchangeServerInformation {
  [CmdletBinding()]
  param($E2010,$ExchangeServer,$Mailboxes,$Databases,$Hybrids)
	
  # Set Basic Variables
  $MailboxCount = 0
  $RollupLevel = 0
  $RollupVersion = ''
  $ExtNames = @()
  $IntNames = @()
  $CASArrayName = ''

  # 2019-05-20 TST Added to handle max preferred/active databases per server
  $MaxPrefDatabases = 0
  $MaxActiveDatabases = 0
  $NotSet = '--'
	
  # Get WMI Information: Operatin System
  $tWMI = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue
  
  if ($tWMI) {
    $OSVersion = $tWMI.Caption.Replace('(R)','').Replace('Microsoft ','').Replace('Enterprise','Ent').Replace('Standard','Std').Replace(' Edition','')
    $OSServicePack = $tWMI.CSDVersion
    $RealName = $tWMI.CSName.ToUpper()
  } 
  else {
    Write-Warning -Message ('Cannot detect OS information via WMI for {0}' -f $ExchangeServer.Name)
    $OSVersion = $NotAvailable
    $OSServicePack = $NotAvailable
    $RealName = $ExchangeServer.Name.ToUpper()
  }

  # Get WMI Information: Disk Space
  $tWMI=Get-WmiObject -Query 'Select * from Win32_Volume' -ComputerName $ExchangeServer.Name -ErrorAction SilentlyContinue

  if ($tWMI) {
    $Disks=$tWMI | Select-Object -Property Name,Capacity,FreeSpace | Sort-Object -Property Name
  } 
  else {
    Write-Warning -Message ('Cannot detect OS information via WMI for {0}' -f $ExchangeServer.Name)
    $Disks=$null
  }
	
  # Get Exchange Version
  if ($ExchangeServer.AdminDisplayVersion.Major -eq 6) {
    $ExchangeMajorVersion = [double]('{0}.{1}' -f $ExchangeServer.AdminDisplayVersion.Major, $ExchangeServer.AdminDisplayVersion.Minor)
    $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.FilePatchLevelDescription.Replace('Service Pack ','')
  } 
  elseif ($ExchangeServer.AdminDisplayVersion.Major -eq 15 -and $ExchangeServer.AdminDisplayVersion.Minor -ge 1) {
    $ExchangeMajorVersion = [double]('{0}.{1}' -f $ExchangeServer.AdminDisplayVersion.Major, $ExchangeServer.AdminDisplayVersion.Minor)
    $ExchangeSPLevel = 0
  } 
  else {
    $ExchangeMajorVersion = $ExchangeServer.AdminDisplayVersion.Major
    $ExchangeSPLevel = $ExchangeServer.AdminDisplayVersion.Minor
  }

  # Exchange 2007+
  if ($ExchangeMajorVersion -ge 8)
  {
    # Get Roles
    $MailboxStatistics=$null
    [array]$Roles = $ExchangeServer.ServerRole.ToString().Replace(' ','').Split(',')
    
    # Add Hybrid "Role" for report
    if ($Hybrids -contains $ExchangeServer.Name) {
      $Roles+='Hybrid'
    }

    if ($Roles -contains 'Mailbox') {

      $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
      if ($ExchangeServer.Name.ToUpper() -ne $RealName) {
        $Roles = [array]($Roles | Where-Object {$_ -ne 'Mailbox'})
        $Roles += 'ClusteredMailbox'
      }

      # Get Mailbox Statistics the normal way, return in a consitent format
      # 2019-05-20 TST, try/catch added
      try {
        $MailboxStatistics = Get-MailboxStatistics -Server $ExchangeServer -ErrorAction SilentlyContinue | Select-Object -Property DisplayName,@{Name='TotalItemSizeB';Expression={$_.TotalItemSize.Value.ToBytes()}},@{Name='TotalDeletedItemSizeB';Expression={$_.TotalDeletedItemSize.Value.ToBytes()}},Database
      }
      catch {
        $MailboxStatistics = $null
        Write-Warning -Message ('Cannot get mailbox statistics for server {0}' -f $ExchangeServer)
      }

      if($ExchangeMajorVersion -ge 14) {
        $mailboxServer = Get-MailboxServer -Identity $($ExchangeServer.Name)

        # 2019-05-20 TST Gather max active/max preferred database config
        if($ExchangeMajorVersion -lt 15) {
          # Exchange 2010
          $MaxActiveDatabases = $mailboxServer.MaximumActiveDatabases
        }
        else {
          # Exchange 2013+
          if($null -ne $mailboxServer.MaximumPreferredActiveDatabases) {
            $MaxPrefDatabases = $mailboxServer.MaximumPreferredActiveDatabases
          }
          else {
            $MaxPrefDatabases = $NotSet
          }
          
          if($null -ne $mailboxServer.MaximumActiveDatabases) {
            $MaxActiveDatabases = $mailboxServer.MaximumActiveDatabases
          }
          else {
            $MaxActiveDatabases = $NotSet
          } 
        }
      }
    }

    # Get HTTPS Names (Exchange 2010 only due to time taken to retrieve data)
    # 2019-05-16 TST | Update to support 'Mailbox' role for gathering namespace information
    if (($Roles -contains 'ClientAccess' -and $E2010) -or ($Roles -contains 'Mailbox' -and $E2010))
    {        
      Get-OWAVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host) }

      Get-WebServicesVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host) }
      
      Get-OABVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host) }
      
      Get-ActiveSyncVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host) }
      
      if (Get-Command -Name Get-MAPIVirtualDirectory -ErrorAction SilentlyContinue) {
        Get-MAPIVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host) }
      }
      
      if (Get-Command -Name Get-ClientAccessService -ErrorAction SilentlyContinue) {
        $IntNames+=(Test-vDirHost -VDirHost (Get-ClientAccessService -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host)
      } 
      else {
        # Fallback to use Get-ClientAccessServer cmdlet
        $IntNames+=(Test-vDirHost -VDirHost (Get-ClientAccessServer -Identity $ExchangeServer.Name).AutoDiscoverServiceInternalURI.Host)
      }
            
      if ($ExchangeMajorVersion -ge 14) {
        Get-ECPVirtualDirectory -Server $ExchangeServer -ADPropertiesOnly | ForEach-Object{ $ExtNames+=(Test-vDirHost -VDirHost $_.ExternalURL.Host); $IntNames+=(Test-vDirHost -VDirHost $_.InternalURL.Host); }
      }

      $IntNames = $IntNames | Sort-Object -Unique
      $ExtNames = $ExtNames | Sort-Object -Unique
      $CASArray = Get-ClientAccessArray -Site $ExchangeServer.Site.Name

      if ($CASArray) {
        $CASArrayName = $CASArray.Fqdn
      }
    }

    # Rollup Level / Versions (Thanks to Bhargav Shukla https://bhargavs.com/index.php/2009/12/14/how-do-i-check-update-rollup-version-on-exchange-20xx-server/)
    switch([string]$ExchangeMajorVersion) {
      # Exchange Server 2016 / 2019
      '15.2' {$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\442189DC8B9EA5040962A6BED9EC1F1F\\Patches"}
      '15.1' {$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\442189DC8B9EA5040962A6BED9EC1F1F\\Patches"}
      # Exchange Server 2010 / 2013
      '15' {$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"}
      '14' {$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\AE1D439464EB1B8488741FFA028E291C\\Patches"}
      # Exchange 2007
      default {$RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\UserData\\S-1-5-18\\Products\\461C2B4266EDEF444B864AD6D9E5B613\\Patches"}
    }
    
    # 2019-05-17 Thomas Stensitzki, try/catch added
    try {
      $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name)
    }
    catch {
      $RemoteRegistry = $null
    }

    if ($null -ne $RemoteRegistry) {
        
      $RUKeys = $RemoteRegistry.OpenSubKey($RegKey).GetSubKeyNames() | ForEach-Object {"$RegKey\\$_"}
     
      if ($RUKeys) {
        [array]($RUKeys | ForEach-Object{$RemoteRegistry.OpenSubKey($_).GetValue('DisplayName')}) | `
        ForEach-Object{
          if ($_ -like 'Update Rollup *') {
            $tRU = $_.Split(' ')[2]            
            if ($tRU -like '*-*') { $tRUV=$tRU.Split('-')[1]; $tRU=$tRU.Split('-')[0] } else { $tRUV='' }            
            if ([int]$tRU -ge [int]$RollupLevel) { $RollupLevel=$tRU; $RollupVersion=$tRUV }
          }
        }
      }
    } 
    else {
      Write-Warning -Message ('Cannot detect Rollup Version via Remote Registry for {0}' -f $ExchangeServer.Name)
    }

    # Exchange 2013+ CU or SP Level
    if ($ExchangeMajorVersion -ge 15) {
      $RegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Microsoft Exchange v15"
      # 2019-05-17 Thomas Stensitzki, try/catch added
      try {
        $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ExchangeServer.Name)
      }
      catch {
        $RemoteRegistry = $null
      }
      
      if ($RemoteRegistry) {
        $ExchangeSPLevel = $RemoteRegistry.OpenSubKey($RegKey).GetValue('DisplayName')
        
        if ($ExchangeSPLevel -like '*Service Pack*' -or $ExchangeSPLevel -like '*Cumulative Update*') {
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2013 ','')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2016 ','')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Microsoft Exchange Server 2019 ','')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Service Pack ','SP')
          $ExchangeSPLevel = $ExchangeSPLevel.Replace('Cumulative Update ','CU') 
        } 
        else {
          $ExchangeSPLevel = 0
        }
      } 
      else {
        Write-Warning -Message ('Cannot detect CU/SP via Remote Registry for {0}' -f $ExchangeServer.Name)
      }
    }
  }

  # Exchange 2003
  if ($ExchangeMajorVersion -eq 6.5) {

    # Mailbox Count
    $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases

    # Get Role via WMI
    $tWMI = Get-WMIObject -Class Exchange_Server -Namespace 'root\microsoftexchangev2' -ComputerName $ExchangeServer.Name -Filter "Name='$($ExchangeServer.Name)'"

    if ($tWMI) {
      if ($tWMI.IsFrontEndServer) { $Roles=@('FE') } else { $Roles=@('BE') }
    } 
    else {
      Write-Warning -Message ('Cannot detect Front End/Back End Server information via WMI for {0}' -f $ExchangeServer.Name)
      $Roles+='Unknown'
    }

    # Get Mailbox Statistics using WMI, return in a consistent format
    $tWMI = Get-WMIObject -class Exchange_Mailbox -Namespace ROOT\MicrosoftExchangev2 -ComputerName $ExchangeServer.Name -Filter ("ServerName='$($ExchangeServer.Name)'")
    if ($tWMI)
    {
      $MailboxStatistics = $tWMI | Select-Object -Property @{Name='DisplayName';Expression={$_.MailboxDisplayName}},@{Name='TotalItemSizeB';Expression={$_.Size}},@{Name='TotalDeletedItemSizeB';Expression={$_.DeletedMessageSizeExtended }},@{Name='Database';Expression={((Get-MailboxDatabase -Identity "$($_.ServerName)\$($_.StorageGroupName)\$($_.StoreName)").Identity)}}
    } 
    else {
      Write-Warning -Message ('Cannot retrieve Mailbox Statistics via WMI for {0}' -f $ExchangeServer.Name)
      $MailboxStatistics = $null
    }
  }	

  # Exchange 2000
  if ($ExchangeMajorVersion -eq '6.0')
  {
    # Mailbox Count
    $MailboxCount = Get-ExchangeServerMailboxCount -Mailboxes $Mailboxes -ExchangeServer $ExchangeServer -Databases $Databases
    
    # Get Role via ADSI
    $tADSI=[ADSI]"LDAP://$($ExchangeServer.OriginatingServer)/$($ExchangeServer.DistinguishedName)"

    if ($tADSI) {
      if ($tADSI.ServerRole -eq 1) { $Roles=@('FE') } else { $Roles=@('BE') }
    } 
    else {
      Write-Warning -Message ('Cannot detect Front End/Back End Server information via ADSI for {0}' -f $ExchangeServer.Name)
      $Roles+='Unknown'
    }
    $MailboxStatistics = $null
  }
	
  # Return Hashtable
  @{Name = $ExchangeServer.Name.ToUpper()
    RealName = $RealName
    ExchangeMajorVersion = $ExchangeMajorVersion
    ExchangeSPLevel	= $ExchangeSPLevel
    Edition = $ExchangeServer.Edition
    Mailboxes = $MailboxCount
    OSVersion = $OSVersion;
    OSServicePack = $OSServicePack
    Roles = $Roles
    RollupLevel = $RollupLevel
    RollupVersion = $RollupVersion
    Site = $ExchangeServer.Site.Name
    MailboxStatistics	= $MailboxStatistics
    Disks = $Disks
    IntNames = $IntNames
    ExtNames = $ExtNames
    CASArrayName = $CASArrayName
    MaximumPreferredDatabases = $MaxPrefDatabases
    MaximumActiveDatabases = $MaxActiveDatabases
  }	
}

# Sub Function to Get Totals by Version
function Get-TotalsByVersion {
  [CmdletBinding()]
  param(
    $ExchangeEnvironment
  )

  # Create empty hash table
  $TotalMailboxesByVersion=@{}

  if ($ExchangeEnvironment.Sites) {
    foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {
      foreach ($Server in $Site.Value) {
        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"]) {
          $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)",@{ServerCount=1;MailboxCount=$Server.Mailboxes})
        } 
        else {
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount+=$Server.Mailboxes
        }
      }
    }
  }

  if ($ExchangeEnvironment.Pre2007) {
    foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator()) {
      foreach ($Server in $FakeSite.Value) {
        if (!$TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"]) {
          $TotalMailboxesByVersion.Add("$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)",@{ServerCount=1;MailboxCount=$Server.Mailboxes})
        } 
        else {
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].ServerCount++
          $TotalMailboxesByVersion["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].MailboxCount+=$Server.Mailboxes
        }
      }
    }
  }
  $TotalMailboxesByVersion
}

# Sub Function to Get Totals by Role
function Get-TotalsByRole {
  [CmdletBinding()]
  param(
    $ExchangeEnvironment
  )

  # Add Roles We Always Show
  $TotalServersByRole=@{
    'ClientAccess' = 0
    'HubTransport' = 0
    'UnifiedMessaging' = 0
    'Mailbox' = 0
    'Edge' = 0
  }

  if ($ExchangeEnvironment.Sites) {

    foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {

      foreach ($Server in $Site.Value) {

        foreach ($Role in $Server.Roles) {
          if ($null -eq $TotalServersByRole[$Role]) {
            $TotalServersByRole.Add($Role,1)
          } 
          else {
            $TotalServersByRole[$Role]++
          }
        }
      }
    }
  }

  if ($ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {
		
    foreach ($Server in $ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {
			
      foreach ($Role in $Server.Roles) {
        if ($null -eq $TotalServersByRole[$Role]) {
          $TotalServersByRole.Add($Role,1)
        } 
        else {
          $TotalServersByRole[$Role]++
        }
      }
    }
  }

  $TotalServersByRole
}

# Sub Function to return HTML Table for Sites/Pre 2007
function Get-HtmlOverview {
  [CmdletBinding()]
  param(
    $Servers,
    $ExchangeEnvironment,
    $ExRoleStrings,
    $Pre2007=$False
  )


  if ($Pre2007) {
    $BGColHeader='#880099'
    $BGColSubHeader='#8800CC'
    $Prefix=''
    $IntNamesText=''
    $ExtNamesText=''
    $CASArrayText=''
    $ColClass = 'pre2007overviewcolheader'
    $ColSubClass = 'pre2007overviewcolsubheader'
  } 
  else {
    $BGColHeader='#000099'
    $BGColSubHeader='#0000FF'
    $Prefix='Site:'
    $IntNamesText=''
    $ExtNamesText=''
    $CASArrayText=''
    $IntNames=@()
    $ExtNames=@()
    $CASArrayName=''
    $ColClass = 'overviewcolheader'
    $ColSubClass = 'overviewcolsubheader'

    foreach ($Server in $Servers.Value) {
      $IntNames+=$Server.IntNames
      $ExtNames+=$Server.ExtNames
      $CASArrayName=$Server.CASArrayName
    }

    $IntNames = $IntNames | Sort-Object -Unique
    $ExtNames = $ExtNames | Sort-Object -Unique
    
    $IntNames = [string]::Join(',',$IntNames)
    $ExtNames = [string]::Join(',',$ExtNames)
    $ExtNamesEmptyText = 'At least one of the analysed servers contains an empty ExternalUrl entry.'

    if ($IntNames) {
      
      $ExtNamesText=('External Names: {0}<br />' -f $ExtNames)

      if($ExtNames -notlike '*None*') {
        $IntNamesText=('Internal Names: {0}' -f $IntNames)
      }
      else {
        $IntNamesText=('Internal Names: {0}<br />{1}' -f $IntNames, $ExtNamesEmptyText)
      }
    }

    if ($CASArrayName) {
      $CASArrayText="CAS Array: $($CASArrayName)"
    }
  }

  $Output="<table class='overview'>
    <col width='20%'>
    <col width='20%'>
  <colgroup width='25%'>"

  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object{$Output+="<col width='3%'>"}
  
  $Output+="</colgroup><col width='20%'><col width='20%'>
    <tr class=""$($ColClass)"">
    <th class='overview'>$($Prefix) $($Servers.Key)</th>
    <th class='overview' colspan=""$(($ExchangeEnvironment.TotalServersByRole.Count)+2)"" align=""left"">$($ExtNamesText)$($IntNamesText)</th>
    <th class='overview' align=""center"">$($CASArrayText)</th>
    <th colspan='2'>&nbsp;</th>
  </tr>"
  $TotalMailboxes=0
  $Servers.Value | ForEach-Object{$TotalMailboxes += $_.Mailboxes}
  $Output+="<tr class=""$($ColSubClass)""><th class='overview'>Mailboxes: $($TotalMailboxes)</th>"
  $Output+="<th class='overview'>Exchange Version</th>"
  $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object{$Output+="<th class='overview'>$($ExRoleStrings[$_.Key].Short)</th>"}

  # 2016-04-19 Thomas Stensitzki Pref/Max Databases added
  $Output+="<th class='overview'>Databases Pref/Max</th><th class='overview'>OS Version</th><th class='overview'>OS Service Pack</th></tr>"
    
  $AlternateRow=0
	
  foreach ($Server in $Servers.Value) {
    $Output+='<tr'

    if ($AlternateRow) {
      $Output+=" class='alternaterow'"
      $AlternateRow=0
    } 
    else {
      $AlternateRow=1
    }

    $Output+=('><td>{0}' -f $Server.Name)
    
    if ($Server.RealName -ne $Server.Name) {
      $Output+=(' ({0})' -f $Server.RealName)
    }

    $Output+="</td><td>$($ExVersionStrings["$($Server.ExchangeMajorVersion).$($Server.ExchangeSPLevel)"].Long)"

    if ($Server.RollupLevel -gt 0) {
      $Output+=(' UR{0}' -f $Server.RollupLevel)
      if ($Server.RollupVersion) {
        $Output+=(' {0}' -f $Server.RollupVersion)
      }
    }

    $Output+='</td>'

    $ExchangeEnvironment.TotalServersByRole.GetEnumerator() | Sort-Object -Property Name | ForEach-Object{ 
      $Output+='<td'
      if ($Server.Roles -contains $_.Key) {
        $Output+=" class='roledata'>"
      }
      else {
        $Output+=' />'
      }
      
      if (($_.Key -eq 'ClusteredMailbox' -or $_.Key -eq 'Mailbox' -or $_.Key -eq 'BE') -and $Server.Roles -contains $_.Key) {
        $Output+=('{0}</td>' -f $Server.Mailboxes)
      } 
    }
		
    # 2016-04-19 Thomas Stensitzki Max Databases added	
    if($Server.Roles -contains 'Mailbox') {	
        $Output+="<td align=""center"">$($Server.MaximumPreferredDatabases) / $($Server.MaximumActiveDatabases)</td><td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>"
    }
    else {
        $Output+=('<td align=""center"">Not Applicable</td><td>{0}</td><td>{1}</td></tr>' -f $Server.OSVersion, $Server.OSServicePack)
    }

    # $Output+="<td>$($Server.OSVersion)</td><td>$($Server.OSServicePack)</td></tr>";	
  }

  $Output+='<tr />
  </table><br />'

  $Output
}

# Sub Function to return HTML Table for Databases
function Get-HtmlDatabaseInformationTable {
  [CmdletBinding()]
  param(
    $Databases
  )

  # Only Show Archive Mailbox Columns, Backup Columns and Circ Logging if at least one DB has an Archive mailbox, backed up or Cir Log enabled.
  $ShowArchiveDBs = $False
  $ShowLastFullBackup = $False
  $ShowCircularLogging = $False
  $ShowStorageGroups = $False
  $ShowCopies = $False
  $ShowFreeDatabaseSpace = $False
  $ShowFreeLogDiskSpace = $False

  foreach ($Database in $Databases) {
    if ($Database.ArchiveMailboxCount -gt 0) {
      $ShowArchiveDBs=$True
    }
    if ($Database.LastFullBackup -ne 'Not Available') {
      $ShowLastFullBackup=$True
    }
    if ($Database.CircularLoggingEnabled -eq 'Yes') {
      $ShowCircularLogging=$True
    }
    if ($Database.StorageGroup) {
      $ShowStorageGroups=$True
    }
    if ($Database.CopyCount -gt 0) {
      $ShowCopies=$True
    }
    if ($null -ne $Database.FreeDatabaseDiskSpace) {
      $ShowFreeDatabaseSpace=$true
    }
    if ($null -ne $Database.FreeLogDiskSpace) {
      $ShowFreeLogDiskSpace=$true
    }
  }
		
  $Output="<table class='databases'>"

  #region database table header
  $Output +="<tr class='databases'>
    <th>Server</th>"

  if ($ShowStorageGroups) {
    $Output+='<th>Storage Group</th>'
  }

  $Output+='<th>Database Name</th>
    <th>Standard Mailboxes</th>
  <th>Av. Mailbox Size</th>'

  if ($ShowArchiveDBs) {
    $Output+='<th>Archive Mailboxes</th><th>Av. Archive Size</th>'
  }

  $Output+='<th>DB Size</th><th>DB Whitespace</th>'
  
  if ($ShowFreeDatabaseSpace) {
    $Output+='<th>Database Disk Free</th>'
  }
  if ($ShowFreeLogDiskSpace) {
    $Output+='<th>Log Disk Free</th>'
  }
  if ($ShowLastFullBackup) {
    $Output+='<th>Last Full Backup</th>'
  }
  if ($ShowCircularLogging) {
    $Output+='<th>Circular Logging</th>'
  }
  if ($ShowCopies) {
    $Output+='<th>DB Copies (n)</th>'
  }

  # Drive names, issue #4
  if($ShowDriveNames) {
    $Output+='<th>EDB / LOG</th>'
  }
	
  $Output+='</tr>'
  #endregion

  $AlternateRow=0

  foreach ($Database in $Databases) {
    $Output+='<tr'

    if ($AlternateRow) {
      $Output+=" class='alternaterow'"
      $AlternateRow=0
    } 
    else {
      $AlternateRow=1
    }
		
    # Close open <tr tag
    $Output+=('><td>{0}</td>' -f $Database.ActiveOwner)

    if ($ShowStorageGroups) {
      $Output+=('><td>{0}</td>' -f $Database.StorageGroup)
    }

    $Output+="<td>$($Database.Name)</td>
      <td class='center'>$($Database.MailboxCount)</td>
    <td class='center'>$("{0:N2}" -f ($Database.MailboxAverageSize/1MB)) MB</td>"

    if ($ShowArchiveDBs) {
      $Output+="<td class=""center"">$($Database.ArchiveMailboxCount)</td> 
      <td class='center'>$("{0:N2}" -f ($Database.ArchiveAverageSize/1MB)) MB</td>"
    }

    if([double]($Database.Size/1GB) -le $MaxDatabaseSize) {
      $Output+="<td class=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>"
    }
    else {
      $Output+="<td class=""center alert"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td>"
    }

    $Output+="<td class='center'>$("{0:N2} GB" -f ($Database.Whitespace/1GB))</td>"

    # $Output+="<td align=""center"">$("{0:N2}" -f ($Database.Size/1GB)) GB </td><td class='center'>$("{0:N2}" -f ($Database.Whitespace/1GB)) GB</td>"

    if ($ShowFreeDatabaseSpace) {
      if([double]($Database.FreeDatabaseDiskSpace) -gt $MinFreeDiskspace) { 
        $Output+="<td class='center'>$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
      }
      else {
        $Output+="<td class='center alert'>$("{0:N1}" -f $Database.FreeDatabaseDiskSpace)%</td>"
      }
    }
    if ($ShowFreeLogDiskSpace) {
      if([double]($Database.FreeLogDiskSpace) -gt $MinFreeDiskspace) {
        $Output+="<td class='center'>$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
      }
      else {
        $Output+="<td class='center alert'>$("{0:N1}" -f $Database.FreeLogDiskSpace)%</td>"
      }
    }
    if ($ShowLastFullBackup) {
      $Output+="<td class='center'>$($Database.LastFullBackup)</td>"
    }
    if ($ShowCircularLogging) {
      $Output+="<td class='center'>$($Database.CircularLoggingEnabled)</td>"
    }
    if ($ShowCopies) {
      $Output+="<td>$($Database.Copies | ForEach-Object{$_}) ($($Database.CopyCount))</td>"
    }
    
    # Drive names, issue #4
    if ($ShowDriveNames) {
      $Output+="<td class='center'>$("{0} / {1}" -f $Database.DriveNameEdb, $Database.DriveNameLog)</td>"
    }
    $Output+='</tr>'
  }

  $Output+='</table><br />'

  $Output += '<p class="dagtablefooter">Explanation</p>'
  $Output += ("<p class='dagtablefooter'>Maximum mailbox database size: {0} GB<br/>Minimum free disk space: {1}%</p>" -f $MaxDatabaseSize, $MinFreeDiskspace)
	
  $Output
}

function Get-HtmlReportHeader {
  [CmdletBinding()]
  param(
    $ExchangeEnvironment,
    $Path
  )

  # Labels and stuff
  $LabelTotalServers = 'Total Servers'
  $LabelTotalMailboxes = 'Total Mailboxes'
  $LabelTotalRoles = 'Total Roles'
  $DateLabelFormat = 'yyyy-MM-dd HH:mm'
  $CssFileName = 'EnvironmentReport.css'
  $UseCss = $true

  # Header
  $Output='
    <html>
    <body>
  <title>Exchange Environment Report</title>'

  if($UseCss -and (Test-Path -Path (Join-Path -Path (Split-Path -Parent $Path) -ChildPath $CssFileName))) {
    $Output += "<style type=""text/css"">$(Get-Content -Path (Join-Path -Path (Split-Path -Parent $Path) -ChildPath $CssFileName))</style>"
  }  

  $Output += ("<h2 align=""center"">Exchange Environment Report</h2><h3 align=""center"">Organization: {3}</h3>
      <h4 align=""center"">Generated {0}</h4>
      <table class='header'>
      <tr class='header'>
  <th colspan=""{1}"" class='header'>{2}</th>" -f (Get-Date -Format $DateLabelFormat), $ExchangeEnvironment.TotalMailboxesByVersion.Count, $LabelTotalServers, $ExchangeEnvironment.OrganizationName)

  if ($ExchangeEnvironment.RemoteMailboxes) {
    $Output+=("<th colspan=""{0}"" class='header'>{1}</th>" -f ($ExchangeEnvironment.TotalMailboxesByVersion.Count+2), $LabelTotalMailboxes)
  } 
  else {
    $Output+=("<th colspan=""{0}"" class='header'>{1}</th>" -f ($ExchangeEnvironment.TotalMailboxesByVersion.Count+1), $LabelTotalMailboxes)
  }
  
  $Output+=("<th colspan=""{0}"" class='header'>{1}</th></tr>
  <tr class='subheader'>" -f $ExchangeEnvironment.TotalServersByRole.Count, $LabelTotalRoles)

  # Show Column Headings based on the Exchange versions we have
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object{$Output+="<th class='subheader'>$($ExVersionStrings[$_.Key].Short)</th>"}
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator() | Sort-Object -Property Name | ForEach-Object{$Output+="<th class='subheader'>$($ExVersionStrings[$_.Key].Short)</th>"}

  if ($ExchangeEnvironment.RemoteMailboxes) {
    $Output+="<th class='subheader'>Office 365</th>"
  }

  $Output+="<th class='subheader'>Org</th>"

  # Exchange Server Roles
  $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort-Object -Property Name| ForEach-Object{$Output+="<th class='subheader'>$($ExRoleStrings[$_.Key].Short)</th>"}

  $Output += '</tr>'

  $Output += "<tr class='headerdata'>"

  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort-Object -Property Name| ForEach-Object{$Output+="<td class='headerdata'>$($_.Value.ServerCount)</td>" }
  $ExchangeEnvironment.TotalMailboxesByVersion.GetEnumerator()|Sort-Object -Property Name| ForEach-Object{$Output+="<td class='headerdata'>$($_.Value.MailboxCount)</td>" }

  if ($RemoteMailboxes) {
    $Output+="<td class='headerdata'>$($ExchangeEnvironment.RemoteMailboxes)</td>"
  }

  $Output += "<td class='headerdata'>$($ExchangeEnvironment.TotalMailboxes)</td>"

  $ExchangeEnvironment.TotalServersByRole.GetEnumerator()|Sort-Object -Property Name| ForEach-Object{$Output+="<td class='headerdata'>$($_.Value)</td>"}

  #$Output+="</tr><tr><tr></table><br>"
  $Output += '</tr></table><!-- End --><br />'

  $Output
}

function Get-HtmlDagHeader {
  [CmdletBinding()]
  param (
    $DAG
  )

  # Database Availability Group Header
  $Output+="<table class=""dagsummary"">
    <col width='20%'><col width='10%'><col width='10%'><col width='70%''>
    <tr class=""dagsummary""><th>Database Availability Group Name</th><th>Member Count</th><th># Databases</th>
    <th>Database Availability Group Members</th></tr>
  <tr><td>$($DAG.Name)</td><td>$($DAG.MemberCount)</td><td>$(($DAG.Databases | Measure-Object).Count)</td><td>" 

  $DAG.Members | ForEach-Object { $Output+=('{0} ' -f $_) }

  $Output += '</td></tr></table><br />'

  $Output
}

# Sub Function to neatly update progress
function Show-ProgressBar {
  [CmdletBinding()]
  param(
    [int]$PercentComplete,
    [string]$Status,
    [int]$Stage
  )

  $TotalStages=5
  Write-Progress -Id 1 -Activity 'Get-ExchangeEnvironmentReport' -Status $Status -PercentComplete (($PercentComplete/$TotalStages)+(1/$TotalStages*$Stage*100))
}

# 1. Initial Startup

# 1.0 Check Powershell Version
if ((Get-Host).Version.Major -eq 1) {
  throw 'Powershell Version 1 not supported'
}

# 1.1 Check Exchange Management Shell, attempt to load
if (!(Get-Command -Name Get-ExchangeServer -ErrorAction SilentlyContinue))
{
  # 2019-05-17 Thomas Stensitzki, Support for Exchange Scripts located in non-default locations
  # Use $env:ExchangeInstallPath for Exchange 2010/2013+ installations
  $ExchangeInstallPath = $env:ExchangeInstallPath

  if (($ExchangeInstallPath -eq '') -or ($null -eq $ExchangeInstallPath)) {
    # $env:ExchangeInstallPath not available on Exchange Server 2007 Setups
    try {
      $ExchangeInstallPath = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Exchange\Setup).MsiInstallPath 
    }
    catch {}
  }

  Write-Verbose -Message ('Exchange Install Path: {0}' -f $ExchangeInstallPath)

  $RemoteExchangePath = Join-Path -Path $ExchangeInstallPath -ChildPath 'bin\RemoteExchange.ps1'
  $LocalExchangePath = Join-Path -Path $ExchangeInstallPath -ChildPath 'bin\Exchange.ps1'

  if (Test-Path -Path $RemoteExchangePath) {
    . $RemoteExchangePath 
    Connect-ExchangeServer -auto
  } 
  elseif (Test-Path -Path $LocalExchangePath) {
    Add-PSSnapIn -Name Microsoft.Exchange.Management.PowerShell.Admin
    . $LocalExchangePath
  } 
  else {
    throw 'Exchange Management Shell cannot be loaded'
  }
}

# 1.2 Check if -SendMail parameter set and if so check -MailFrom, -MailTo and -MailServer are set
if ($SendMail)
{
  if (!$MailFrom -or !$MailTo -or !$MailServer)
  {
    throw 'If -SendMail specified, you must also specify -MailFrom, -MailTo and -MailServer'
  }
}

# 1.3 Check Exchange Management Shell Version
if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue))
{
  $E2010 = $false;
  if (Get-ExchangeServer | Where-Object {$_.AdminDisplayVersion.Major -gt 14}) {
    Write-Warning -Message "Exchange 2010 or higher detected. You'll get better results if you run this script from the latest management shell"
  }
}
else{
    
  $E2010 = $true

  # 2019-05-17 Thomas Stensitzki, Support for Exchange 2013+ servers with installed management tools
  $localversion = $localserver = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiProductMajor

  if ($localversion -eq 15) { $E2013 = $true }
}

# 1.4 Check view entire forest if set (by default, true)
if ($E2010) {
  Set-ADServerSettings -ViewEntireForest:$ViewEntireForest
} 
else {
  $global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
}

# 1.5 Initial Variables

# 1.5.1 Hashtable to update with environment data
$ExchangeEnvironment = @{
  Sites = @{}
  Pre2007 = @{}
  Servers = @{}
  DAGs = @()
  NonDAGDatabases = @()
  OrganizationName = ''
}

# 1.5.7 Exchange Major Version String Mapping
$ExMajorVersionStrings = @{
  '6.0' = @{Long='Exchange 2000';Short='E2000'}
  '6.5' = @{Long='Exchange 2003';Short='E2003'}
  '8'   = @{Long='Exchange 2007';Short='E2007'}
  '14'  = @{Long='Exchange 2010';Short='E2010'}
  '15'  = @{Long='Exchange 2013';Short='E2013'}
  '15.1'  = @{Long='Exchange 2016';Short='E2016'}
  '15.2'  = @{Long='Exchange 2019';Short='E2019'} #2019-05-17 TST Exchange Server 2019 added
}

# 1.5.8 Exchange Service Pack String Mapping
$ExSPLevelStrings = @{
  '0' = 'RTM'
  '1' = 'SP1'
  '2' = 'SP2'
  '3' = 'SP3'
  '4' = 'SP4'
  'SP1' = 'SP1'
'SP2' = 'SP2'}

# Add many CUs               
for ($i = 1; $i -le 40; $i++) {
  $ExSPLevelStrings.Add("CU$($i)","CU$($i)");
}

# 1.5.9 Populate Full Mapping using above info
$ExVersionStrings = @{}

foreach ($Major in $ExMajorVersionStrings.GetEnumerator()) {
  foreach ($Minor in $ExSPLevelStrings.GetEnumerator()) {
    $ExVersionStrings.Add("$($Major.Key).$($Minor.Key)",@{Long="$($Major.Value.Long) $($Minor.Value)";Short="$($Major.Value.Short)$($Minor.Value)"})
  }
}
# 1.5.10 Exchange Role String Mapping
$ExRoleStrings = @{'ClusteredMailbox' = @{Short='ClusMBX';Long='CCR/SCC Clustered Mailbox'}
  'Mailbox' = @{Short='MBX';Long='Mailbox'}
  'ClientAccess' = @{Short='CAS';Long='Client Access'}
  'HubTransport'	 = @{Short='HUB';Long='Hub Transport'}
  'UnifiedMessaging' = @{Short='UM';Long='Unified Messaging'}
  'Edge' = @{Short='EDGE';Long='Edge Transport'}
  'FE'	 = @{Short='FE';Long='Front End'}
  'BE'	 = @{Short='BE';Long='Back End'}
  'Hybrid' = @{Short='HYB'; Long='Hybrid'}
  'Coexistence' = @{Short='COEX'; Long='Coexistence'} #2019-05-17 TST Coexistence added
'Unknown' = @{Short='Unknown';Long='Unknown'}}

# 2 Get Relevant Exchange Information Up-Front

# 2.1 Get Server, Exchange and Mailbox Information
Show-ProgressBar -PercentComplete 1 -Status 'Getting Exchange Server List' -Stage 1

$ExchangeServers = [array](Get-ExchangeServer $ServerFilter | Sort-Object Name)
if (!$ExchangeServers) {
  throw ('No Exchange Servers matched by -ServerFilter {0}' -f $ServerFilter)
}

$HybridServers=@()
if (Get-Command -Name Get-HybridConfiguration -ErrorAction SilentlyContinue) {
  $HybridConfig = Get-HybridConfiguration
  $HybridConfig.ReceivingTransportServers | ForEach-Object{ $HybridServers+=$_.Name  }
  $HybridConfig.SendingTransportServers | ForEach-Object{ $HybridServers+=$_.Name  }
  $HybridServers = $HybridServers | Sort-Object -Unique
}

Show-ProgressBar -PercentComplete 10 -Status 'Getting Mailboxes' -Stage 1

$Mailboxes = [array](Get-Mailbox -ResultSize Unlimited) | Where-Object {$_.ServerName -like $ServerFilter}

if ($E2010) { 

  Show-ProgressBar -PercentComplete 60 -Status 'Getting Archive Mailboxes' -Stage 1

  $ArchiveMailboxes = [array](Get-Mailbox -Archive -ResultSize Unlimited) | Where-Object {$_.ServerName -like $ServerFilter}
  
  Show-ProgressBar -PercentComplete 70 -Status 'Getting Remote Mailboxes' -Stage 1

  $RemoteMailboxes = [array](Get-RemoteMailbox -ResultSize Unlimited)
  $ExchangeEnvironment.Add('RemoteMailboxes',$RemoteMailboxes.Count)
  
  Show-ProgressBar -PercentComplete 90 -Status 'Getting Databases' -Stage 1

  if ($E2013) {	
    # 2019-05-17 TST Sorting added
    $Databases = [array](Get-MailboxDatabase -IncludePreExchange2013 -Status) | Sort-Object -Property Name | Where-Object {$_.Server -like $ServerFilter} 
  }
  elseif ($E2010) {	
    # 2019-05-17 TST Sorting added
    $Databases = [array](Get-MailboxDatabase -IncludePreExchange2010 -Status) | Sort-Object -Property Name | Where-Object {$_.Server -like $ServerFilter} 
  }

  $DAGs = [array](Get-DatabaseAvailabilityGroup) | Where-Object {$_.Servers -like $ServerFilter}
} 
else {
  $ArchiveMailboxes = $null
  $ArchiveMailboxStats = $null	
  $DAGs = $null

  Show-ProgressBar -PercentComplete 90 -Status 'Getting Databases' -Stage 1
  $Databases = [array](Get-MailboxDatabase -IncludePreExchange2007 -Status) | Where-Object {$_.Server -like $ServerFilter}
  $ExchangeEnvironment.Add('RemoteMailboxes',0)
}

# 2.3 Populate Information we know
$ExchangeEnvironment.Add('TotalMailboxes',$Mailboxes.Count + $ExchangeEnvironment.RemoteMailboxes)

# 2.4 Organizational Info

$ExchangeEnvironment.OrganizationName = (Get-OrganizationConfig).Name

# 3 Process High-Level Exchange Information

# 3.1 Collect Exchange Server Information
for ($i=0; $i -lt $ExchangeServers.Count; $i++) {
  Show-ProgressBar -PercentComplete ($i/$ExchangeServers.Count*100) -Status 'Getting Exchange Server Information' -Stage 2

  # Get Exchange Info
  $ExSvr = Get-ExchangeServerInformation -E2010 $E2010 -ExchangeServer $ExchangeServers[$i] -Mailboxes $Mailboxes -Databases $Databases -Hybrids $HybridServers
  
  # Add to site or pre-Exchange 2007 list
  if ($ExSvr.Site) {
    # Exchange 2007 or higher
    if (!$ExchangeEnvironment.Sites[$ExSvr.Site]) {
      $ExchangeEnvironment.Sites.Add($ExSvr.Site,@($ExSvr))
    } 
    else {
      $ExchangeEnvironment.Sites[$ExSvr.Site]+=$ExSvr
    }
  } 
  else {
    # Exchange 2003 or lower
    if (!$ExchangeEnvironment.Pre2007['Pre 2007 Servers']) {
      $ExchangeEnvironment.Pre2007.Add('Pre 2007 Servers',@($ExSvr))
    } 
    else {
      $ExchangeEnvironment.Pre2007['Pre 2007 Servers']+=$ExSvr
    }
  }

  # Add to Servers List
  $ExchangeEnvironment.Servers.Add($ExSvr.Name,$ExSvr)
}

# 3.2 Calculate Environment Totals for Version/Role using collected data
Show-ProgressBar -PercentComplete 1 -Status 'Getting Totals' -Stage 3

$ExchangeEnvironment.Add('TotalMailboxesByVersion',(Get-TotalsByVersion -ExchangeEnvironment $ExchangeEnvironment))
$ExchangeEnvironment.Add('TotalServersByRole',(Get-TotalsByRole -ExchangeEnvironment $ExchangeEnvironment))

# 3.4 Populate Environment DAGs
Show-ProgressBar -PercentComplete 5 -Status 'Getting DAG Info' -Stage 3

if ($DAGs) {
  foreach($DAG in $DAGs) {
    $ExchangeEnvironment.DAGs+=(Get-DatabaseAvailabilityGroupInformation -DAG $DAG)
  }
}

# 3.5 Get Database information
Show-ProgressBar -PercentComplete 60 -Status 'Getting Database Info' -Stage 3

for ($i=0; $i -lt $Databases.Count; $i++)
{
  $Database = Get-DatabaseInformation -Database $Databases[$i] -ExchangeEnvironment $ExchangeEnvironment -Mailboxes $Mailboxes -ArchiveMailboxes $ArchiveMailboxes -E2010 $E2010
  $DAGDB = $false
  for ($j=0; $j -lt $ExchangeEnvironment.DAGs.Count; $j++) {
    if ($ExchangeEnvironment.DAGs[$j].Members -contains $Database.ActiveOwner) {
      $DAGDB=$true
      $ExchangeEnvironment.DAGs[$j].Databases += $Database
    }
  }
  if (!$DAGDB) {
    $ExchangeEnvironment.NonDAGDatabases += $Database
  }
}

# 4 Write Information
Show-ProgressBar -PercentComplete 5 -Status 'Writing HTML Report Header' -Stage 4

$Output = Get-HtmlReportHeader -ExchangeEnvironment $ExchangeEnvironment -Path $MyInvocation.MyCommand.Path

# Sites and Servers
Show-ProgressBar -PercentComplete 20 -Status 'Writing HTML Site Information' -Stage 4

foreach ($Site in $ExchangeEnvironment.Sites.GetEnumerator()) {
  $Output+=Get-HtmlOverview -Servers $Site -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings
}

Show-ProgressBar -PercentComplete 40 -Status 'Writing HTML Pre-2007 Information' -Stage 4

foreach ($FakeSite in $ExchangeEnvironment.Pre2007.GetEnumerator()) {
  $Output+=Get-HtmlOverview -Servers $FakeSite -ExchangeEnvironment $ExchangeEnvironment -ExRoleStrings $ExRoleStrings -Pre2007:$true
}

Show-ProgressBar -PercentComplete 60 -Status 'Writing HTML DAG Information' -Stage 4

foreach ($DAG in $ExchangeEnvironment.DAGs) {

  if ($DAG.MemberCount -gt 0) {

    # Get DAG Header
    $Output += Get-HtmlDagHeader -DAG $DAG
		
    # Get Table HTML for DAG databases
    $Output += Get-HtmlDatabaseInformationTable -Databases $DAG.Databases
  }
}

if ($ExchangeEnvironment.NonDAGDatabases.Count) {

  Show-ProgressBar -PercentComplete 80 -Status 'Writing HTML Non-DAG Database Information' -Stage 4
  
  $Output+='<table class="dagsummary">
  <tr class="dagsummarynondag"><th>Mailbox Databases (Non-DAG)</th></table>'

  # Get Table HTML for non-DAG databases
  $Output+=Get-HtmlDatabaseInformationTable -Databases $ExchangeEnvironment.NonDAGDatabases
}


# End
Show-ProgressBar -PercentComplete 90 -Status 'Finishing off..' -Stage 4

$Output+='</body></html>'

# 2019-05-20 TST Updated to ensure script path as storage location
$HtmlReportFullPath = Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath $HTMLReport

$Output | Out-File -FilePath $HtmlReportFullPath -Force -Encoding utf8 


if ($SendMail)
{
  Show-ProgressBar -PercentComplete 95 -Status 'Sending mail message..' -Stage 4

  # 2019-05-17 TST, Changed to .NET send method to work as scheduled job

  $smtpMail = New-Object Net.Mail.SmtpClient($MailServer) 
  
  $smtpMessage = New-Object System.Net.Mail.MailMessage $MailFrom, $MailTo

  if(Test-Path -Path $HtmlReportFullPath) {
    $smtpAttachment = New-Object Net.Mail.Attachment($HtmlReportFullPath, 'text/plain')
    $smtpMessage.Attachments.Add($smtpAttachment)
  }

  $smtpMessage.Subject = 'Exchange Environment Report'
  $smtpMessage.Body = $Output   
  $smtpMessage.IsBodyHtml = $true

  $smtpMail.Send($smtpMessage)

  Return 0
}