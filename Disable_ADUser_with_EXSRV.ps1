<#
.SYNOPSIS
Manages the deactivation process of Active Directory user accounts, including mailbox exports and directory archiving.

.DESCRIPTION
This script provides comprehensive user deactivation functionality in Active Directory environments. It:
- Exports user mailboxes and archives to PST files
- Archives home and Citrix profile directories
- Cleans up AD attributes
- Manages export requests through a SQLite cache database
- Provides detailed logging of all operations

.PARAMETER user
The samAccountName of the user account to be deactivated.

.PARAMETER requiredModules
Array of PowerShell modules required by the script.

.PARAMETER sleepDuration
Time in seconds to wait between mailbox export operations.

.EXAMPLE
.\UserDeactivation-ADManager.ps1 -user "jdoe"
Deactivates the user account "jdoe", exports their mailbox, and archives their directories.

.NOTES
Version:        0.2
Author:         Anonymized
Created:        06.07.2023
Last Updated:   27.12.2024
Requires:       
- Windows PowerShell 5.1 or later
- Active Directory module
- Exchange Management Shell access
- SQLite package
- Network access to file shares and Exchange server

#>

Param(
    [string]$user,
    [string[]]$requiredModules = @(
        'ActiveDirectory',
        'Microsoft.PowerShell.SecretManagement',
        'Microsoft.PowerShell.SecretStore'
    ),
    [int]$sleepDuration = 420
)

$date = (Get-Date).ToString("dd/MM/yyyy")
$logFile = "C:\Temp\Log\proc_$env:computername.log"

if (-Not (Test-Path $logFile)) {
    New-Item -Path $logFile -ItemType File
}

if (-Not (Get-Package -Name System.Data.SQLite -ProviderName NuGet -ErrorAction SilentlyContinue)) {
    Install-Package -Name System.Data.SQLite -Source https://www.nuget.org/api/v2
}

function WriteLog {
    Param(
        [string]$message,
        [string]$level = "INFO"
    )
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "[$level] $stamp - $message"
    Add-Content $logfile -Value $logMessage
}

function Initialize-CacheDatabase {
    $dbPath = "C:\Temp\Log\cache.db"
    if (-Not (Test-Path $dbPath)) {
        $connectionString = "Data Source=$dbPath;Version=3;"
        $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $connection.Open()

        $createTableQuery = @"
        CREATE TABLE Cache (
            User TEXT PRIMARY KEY,
            Timestamp DATETIME,
            Status TEXT
        )
"@

        $command = $connection.CreateCommand()
        $command.CommandText = $createTableQuery
        $command.ExecuteNonQuery()

        $connection.Close()
        WriteLog "Cache database initialized."
    }
}

function Add-CacheEntry {
    param (
        [string]$user,
        [string]$status
    )
    $dbPath = "C:\Temp\Log\cache.db"
    $connectionString = "Data Source=$dbPath;Version=3;"
    $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
    $connection.Open()

    $query = "INSERT OR REPLACE INTO Cache (User, Timestamp, Status) VALUES (@User, @Timestamp, @Status)"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@User", $user)))
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@Timestamp", (Get-Date))))
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@Status", $status)))
    $command.ExecuteNonQuery()

    $connection.Close()
    WriteLog "Cache entry added/updated for user: $user"
}

function Get-CachedUsers {
    $dbPath = "C:\Temp\Log\cache.db"
    $connectionString = "Data Source=$dbPath;Version=3;"
    $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
    $connection.Open()

    $query = "SELECT User FROM Cache"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $reader = $command.ExecuteReader()

    $users = @()
    while ($reader.Read()) {
        $users += $reader["User"]
    }

    $connection.Close()
    return $users
}

function Update-CacheEntry {
    param (
        [string]$user,
        [string]$status
    )
    $dbPath = "C:\Temp\Log\cache.db"
    $connectionString = "Data Source=$dbPath;Version=3;"
    $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
    $connection.Open()

    $query = "UPDATE Cache SET Timestamp = @Timestamp, Status = @Status WHERE User = @User"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@User", $user)))
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@Timestamp", (Get-Date))))
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@Status", $status)))
    $command.ExecuteNonQuery()

    $connection.Close()
    WriteLog "Cache entry updated for user: $user with status: $status"
}

function Remove-CacheEntry {
    param (
        [string]$user
    )
    $dbPath = "C:\Temp\Log\cache.db"
    $connectionString = "Data Source=$dbPath;Version=3;"
    $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
    $connection.Open()

    $query = "DELETE FROM Cache WHERE User = @User"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.Parameters.Add((New-Object System.Data.SQLite.SQLiteParameter("@User", $user)))
    $command.ExecuteNonQuery()

    $connection.Close()
    WriteLog "Cache entry removed for user: $user"
}

function Connect-PSExchSession {
    try {
        if (-not $UserCredential) {
            try {
                $UserCredential = Get-Secret -Name "mySecretCred"
                if (-not $UserCredential) {
                    throw "Credential 'mySecretCred' not found"
                }
            }
            catch {
                WriteLog "Failed to retrieve secret 'mySecretCred': $_" "ERROR"
                exit 1
            }
        }

        WriteLog "Establishing Remote PowerShell Session with MyServerEXCH01.mydomain.local."
        $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://MyServerEXCH01.mydomain.local/PowerShell/ -Authentication Kerberos -Credential $UserCredential
        return $s
    }
    catch {
        WriteLog "Failed to establish PSSession: $_" "ERROR"
        exit 1
    }
}

function Exit-PSSession {
    param($Session)
    WriteLog "Failed to establish PSSession: $_" "ERROR"
    exit 1
}

function Export-Mailbox {
    param($user)
    try {
        $s = Connect-PSExchSession

        $exportPath = "\\server.mydomain.local\MailboxArchive$\Archived\$($date)_$($user).pst"
        $archivePath = "\\server.mydomain.local\MailboxArchive$\Archived\$($date)_$($user)_Archive.pst"
        
        if (-not (Test-Path (Split-Path -Path $exportPath -Parent))) {
            WriteLog "Export path $exportPath does not exist." "ERROR"
            exit
        }

        $state = Get-Mailbox -Identity $user | Select-Object -ExpandProperty ArchiveState
        New-MailboxExportRequest -Mailbox $user -FilePath $exportPath

        if ($state -eq 'Local') {
            if (-not (Test-Path (Split-Path -Path $archivePath -Parent))) {
                WriteLog "Archive path $archivePath does not exist." "ERROR"
                exit
            }
            New-MailboxExportRequest -Mailbox $user -IsArchive -FilePath $archivePath
        }

        WriteLog "Mailbox $user prepared for export."
        Add-CacheEntry -user $user -status "PREPARED"
        
        Exit-PSSession -Session $s
    }
    catch {
        WriteLog "Failed to create Mailbox Export Request: $_" "ERROR"
        exit
    }
}

function Remove-CompletedMailboxExportRequests {
    $s = Connect-PSExchSession
    $cachedUsers = Get-CachedUsers
    $completedRequests = Get-MailboxExportRequest -ResultSize Unlimited | Get-MailboxExportRequestStatistics | Where-Object { $_.PercentComplete -eq 100 -and $cachedUsers -contains $_.SourceAlias }
    if ($null -eq $completedRequests) {
        WriteLog "No completed export requests found in cache."
    } else {
        foreach ($request in $completedRequests) {
            try {
                Get-MailboxExportRequest -Mailbox $request.SourceAlias | Remove-MailboxExportRequest -Confirm:$false
                WriteLog "Export request for mailbox $($request.SourceAlias) has been removed."
                Update-CacheEntry -user $request.SourceAlias -status "REMOVED"
            }
            catch {
                WriteLog "Error removing export request for mailbox $($request.SourceAlias): $_" "ERROR"
            }
        }
    }
    Exit-PSSession -Session $s
}

function Disable-RequestedMailboxes {
    $s = Connect-PSExchSession
    $cachedUsers = Get-CachedUsers | Where-Object { $_.Status -eq "REMOVED" }
    if ($null -eq $cachedUsers) {
        WriteLog "No mailboxes to disable."
    } else {
        foreach ($user in $cachedUsers) {
            if (Get-Mailbox -Identity $user.User) {
                try {
                    Disable-Mailbox -Identity $user.User -Confirm:$false
                    WriteLog "Mailbox $($user.User) has been disabled."
                    Remove-CacheEntry -user $user.User
                }
                catch {
                    WriteLog "Error disabling mailbox $($user.User): $_" "ERROR"
                }
            }
        }
    }
    Exit-PSSession -Session $s
}

WriteLog "------------------------------------------"
WriteLog "Starting mailbox export for user: $user"

if (-Not (Get-Module -ListAvailable -Name 'Microsoft.PowerShell.SecretManagement', 'Microsoft.PowerShell.SecretStore')) {
    try {
        WriteLog "Installing SecretManagement modules..."
        Install-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -Scope CurrentUser -Force
        Register-SecretVault -Name MySecretVault -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
        Set-Secret -Name "mySecretCred" -Secret (Get-Credential)
    }
    catch {
        WriteLog "Failed to install SecretManagement modules: $_" "ERROR"
        exit 1
    }
}

foreach ($module in $requiredModules) {
    try {
        WriteLog "Importing module: $module"
        Import-Module $module -Force
    }
    catch {
        WriteLog "Failed to import module $module : $_" "ERROR"
        exit 1
    }
}

try {
    WriteLog "Starting export of home and Citrix directories."

    $pathMappings = @(
        @{
            Source = "\\mydomain.local\Files\MyDC\User"
            Dest = "\\mydomain.local\Files\MyDC\User\1_Archive"
        },
        @{
            Source = "\\mydomain.local\Files\MyDC\CTX_Profile_Store"
            Dest = "\\mydomain.local\Files\MyDC\CTX_Profile_Store\1_Archive"
        },
        @{
            Source = "\\mydomain.local\Files\MyDC\CTX_Profile_Store_SITE02"
            Dest = "\\mydomain.local\Files\MyDC\CTX_Profile_Store_SITE02\1_Archive"
        }
    )

    foreach ($mapping in $pathMappings) {
        $sourcePath = Join-Path $mapping.Source $user
        $destPath = Join-Path $mapping.Dest "$($date)_$($user)"
        
        if (Test-Path $sourcePath) {
            WriteLog "Moving from $sourcePath to $destPath"
            Move-Item -Path $sourcePath -Destination $destPath -ErrorAction Stop
            WriteLog "Successfully moved $sourcePath"
        }
    }

    WriteLog "Completed export of home and Citrix directories."
}
catch {
    WriteLog "Error exporting directories: $_" "ERROR"
}

$attributesToRemove = @('manager', 'extensionAttribute12', 'extensionAttribute13')

foreach ($attribute in $attributesToRemove) {
    try {
        $theUser = Get-ADUser -Identity $user -Properties $attribute
        if ($null -ne $theUser.$attribute) {
            WriteLog "Attempting to remove AD attribute '$attribute'."
            Set-ADUser -Identity $theUser -Clear $attribute
            WriteLog "AD attribute '$attribute' removed."
        }
    }
    catch {
        WriteLog "Error removing AD attribute '$attribute': $_" "ERROR"
    }
}

Initialize-CacheDatabase
Export-Mailbox -user $user
Start-Sleep -Seconds $sleepDuration
Remove-CompletedMailboxExportRequests
Disable-RequestedMailboxes

WriteLog "Completed mailbox export for user: $user"
WriteLog "------------------------------------------"
