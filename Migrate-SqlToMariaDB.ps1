<#
.SYNOPSIS
    Migrates relational data from a Microsoft SQL Server source database (Dummy) to a
    MariaDB target database (simplimed) with full primary-key, foreign-key and GUID
    remapping so the migrated graph remains relationally consistent.

.DESCRIPTION
    The script implements the migration described in the project mapping CSV
    (table_datafeld_mapping.csv). Source tables (Tabelle_Patienten, Tabelle_PatientenWv,
    Tabelle_Abre, Tabelle_AbreRe, Tabelle_JourBuch, Tabelle_TermPro) are projected into
    the new schema (tbl_tenant_entities, tbl_entity_main, tbl_contact_main,
    tbl_receipt_main, tbl_appoint_main, tbl_documentation_main, tbl_receipt_content,
    tbl_bookkeeping, tbl_appoint_protocoll) in parent-before-child order so foreign
    keys can be resolved through in-memory remap dictionaries.

    Key features

      * AUTO_INCREMENT-based PK generation in MariaDB. No identity values are copied.
      * Per-table hashtables map source integer IDs to newly assigned MariaDB IDs and
        are consulted whenever a child row references a parent.
      * Brand-new GUIDs (System.Guid.NewGuid()) replace every source GuiID so
        downstream systems cannot accidentally collide with the legacy data set.
      * A pre-flight backup (mysqldump) of the simplimed target database is taken by
        default before any write happens. The backup file path is reported at the
        end of the run.
      * Streaming SqlDataReader on the source side and prepared, batched parameterised
        INSERTs on the target side so 100k+ row tables migrate without exhausting
        memory.
      * -WhatIf performs a complete dry run: it reads the source, builds the same id
        maps, validates row eligibility, but performs no writes against MariaDB.
      * Idempotent: when re-run against an empty (or freshly restored) target, it
        produces the same logical state. When re-run against a populated target, it
        simply layers a new batch of rows on top using the next AUTO_INCREMENT range,
        which is also safe for the "future tenants" requirement.

.PARAMETER WhatIf
    Reads from SQL Server and builds the id maps without writing to MariaDB.

.PARAMETER SkipBackup
    Skips the pre-flight mysqldump backup step. Use only when an external backup is
    already in place; otherwise leave on for safety.

.PARAMETER MaxRows
    Optional safety cap. When set to a positive integer N, only the first N source
    rows per table are migrated. Useful for spot-check / smoke tests.

.PARAMETER LogDirectory
    Directory where the timestamped transcript log is written. Default
    C:\Logs\SqlToMariaDB.

.PARAMETER BackupDirectory
    Directory where the mysqldump backup is written. Default
    C:\Backups\simplimed.

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1
    Interactive mode — shows the main menu (Migrate / Verify / Repair / Cleanup / DryRun).

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1 -Mode Migrate
    Full SQL Server -> MariaDB migration with backup, transcript and Step10 verify.

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1 -Mode Verify
    Read-only relational integrity report against the current target.

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1 -Mode Repair
    In-place auto-zero of all lookup orphans plus full verification (no migration,
    no truncation). Safe against populated production databases.

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1 -Mode Cleanup
    TRUNCATE all migrated tables in the target database. Refuses to run unless
    the database name contains "staging" or -Force is supplied (which then asks
    for an additional name confirmation).

.EXAMPLE
    .\Migrate-SqlToMariaDB.ps1 -Mode Migrate -WhatIf
    Dry run — reads source, builds id maps, performs no target writes.

.NOTES
    Author        : SQL/AD Automation
    Last update   : 2026-04-28
    Requires      : PowerShell 5.1, SimplySql 2.x (auto-installed if missing)
                    System.Data.SqlClient (built-in to .NET Framework 4.x)
                    mysqldump.exe on PATH for the optional backup step
#>

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    # Mode selector. When omitted, an interactive menu is shown.
    #   Migrate : full SQL Server -> MariaDB run (Steps 1-9 + Verify)
    #   Verify  : read-only relational integrity report
    #   Repair  : in-place auto-zero of lookup orphans + Verify
    #   Cleanup : TRUNCATE all migrated tables in target (only allowed on
    #             databases whose name contains 'staging' OR with -Force)
    [Parameter()] [ValidateSet('Migrate','Verify','Repair','Cleanup')]
    [string] $Mode,

    [Parameter()] [switch] $SkipBackup,
    [Parameter()] [int]    $MaxRows         = 0,
    [Parameter()] [string] $LogDirectory    = 'C:\Logs\SqlToMariaDB',
    [Parameter()] [string] $BackupDirectory = 'C:\Backups\simplimed',
    [Parameter()] [int]    $BatchSize       = 500,
    # Override the staging-name guard on Cleanup. Without -Force the cleanup
    # only proceeds against databases matching *staging*.
    [Parameter()] [switch] $Force,

    [Parameter()] [string] $SourceConnectionString =
        'Server=192.168.198.1,1433;Database=Dummy;User Id=sa;Password=!nic7774;Encrypt=True;TrustServerCertificate=True;Connection Timeout=15',
    [Parameter()] [string] $TargetConnectionString =
        'Server=192.168.198.1;Port=3306;Database=simplimed;Uid=root;Pwd=!nic7774;Connection Timeout=15;AllowUserVariables=True'
)

#region Initialization ---------------------------------------------------------

$ErrorActionPreference = 'Stop'
$script:StartTime      = Get-Date
$timestamp             = $script:StartTime.ToString('yyyyMMdd_HHmmss')

# Capture WhatIf intent into a script-local boolean and immediately clear the
# automatic preference so that 3rd-party cmdlets (SimplySql) always execute
# their queries. Writes to the target database are gated through $script:DryRun.
$script:DryRun         = [bool]$WhatIfPreference
$WhatIfPreference      = $false

if (-not (Test-Path -LiteralPath $LogDirectory))    { New-Item -ItemType Directory -Path $LogDirectory    -Force -WhatIf:$false | Out-Null }
if (-not (Test-Path -LiteralPath $BackupDirectory)) { New-Item -ItemType Directory -Path $BackupDirectory -Force -WhatIf:$false | Out-Null }

$script:TranscriptPath = Join-Path $LogDirectory ("Migrate-SqlToMariaDB_{0}.log" -f $timestamp)
$script:BackupPath     = Join-Path $BackupDirectory ("simplimed_pre_migration_{0}.sql" -f $timestamp)

Start-Transcript -Path $script:TranscriptPath -Append -WhatIf:$false | Out-Null
$VerbosePreference = 'SilentlyContinue'

# Aggregate counters & error collection -------------------------------------
$script:Stats           = [ordered]@{}
$script:ErrorCollection = [System.Collections.Generic.List[psobject]]::new()
$script:WarnCollection  = [System.Collections.Generic.List[psobject]]::new()

# Cross-reference id maps (source-int-id -> new-int-id) ---------------------
# A "tenant" map is intentionally separate per source patient because a single
# Tabelle_Patienten row can only become a tenant *or* an entity *or* a contact.
$script:Maps = @{
    Tenant     = @{}    # IDP / ID0(VIP)        -> tbl_tenant_entities.fld_id
    Entity     = @{}    # IDM / ID0(employee)   -> tbl_entity_main.fld_id
    Contact    = @{}    # ID0(non-VIP non-emp)  -> tbl_contact_main.fld_id
    Receipt    = @{}    # Tabelle_AbreRe.ID1    -> tbl_receipt_main.fld_id
    Appoint    = @{}    # Tabelle_PatientenWv.ID2 -> tbl_appoint_main.fld_id
    Doc        = @{}    # Tabelle_Abre.ID2 (ID1<10) -> tbl_documentation_main.fld_id
    RcptCont   = @{}    # Tabelle_Abre.ID2 (ID1>10) -> tbl_receipt_content.fld_id
    Booking    = @{}    # Tabelle_JourBuch.ID0  -> tbl_bookkeeping.fld_id
    Protocoll  = @{}    # Tabelle_TermPro.IDA   -> tbl_appoint_protocoll.fld_id
}

# Tenant default fallback - reused for orphaned FKs to satisfy NOT NULL FKs.
$script:DefaultTenantId = 0

#endregion

#region Console UI Helpers -----------------------------------------------------

function Write-Banner {
    param([string]$Title, [string]$Color = 'Cyan')
    $line = '=' * 78
    Write-Host ''
    Write-Host $line -ForegroundColor $Color
    Write-Host (' {0}' -f $Title) -ForegroundColor $Color
    Write-Host $line -ForegroundColor $Color
}

function Write-Section {
    param([string]$Text, [string]$Color = 'Yellow')
    Write-Host ''
    Write-Host ('--- {0} ---' -f $Text) -ForegroundColor $Color
}

function Write-OK   { param([string]$Text) Write-Host (' [ OK ] {0}' -f $Text) -ForegroundColor Green }
function Write-Warn2 { param([string]$Text) Write-Host (' [WARN] {0}' -f $Text) -ForegroundColor Yellow }
function Write-Err  { param([string]$Text) Write-Host (' [FAIL] {0}' -f $Text) -ForegroundColor Red }

function Get-TargetDatabaseName {
    if ($TargetConnectionString -match '(?i)(?:Database|Initial Catalog)\s*=\s*([^;]+)') { return $matches[1].Trim() }
    return '<unknown>'
}

function Show-MainMenu {
    [CmdletBinding()] param()
    while ($true) {
        Clear-Host
        $tgt = Get-TargetDatabaseName
        $line = '=' * 78
        Write-Host $line -ForegroundColor Cyan
        Write-Host ' SQL Server -> MariaDB Migration Toolkit' -ForegroundColor Cyan
        Write-Host $line -ForegroundColor Cyan
        Write-Host ('   Target database  : {0}' -f $tgt)         -ForegroundColor Gray
        Write-Host ('   Source           : SQL Server "Dummy" (192.168.198.1)') -ForegroundColor Gray
        Write-Host ('   Log directory    : {0}' -f $LogDirectory) -ForegroundColor Gray
        Write-Host ''
        Write-Host '   [1]  Full migration   (SQL Server  ->  MariaDB, Steps 1-9 + Verify)' -ForegroundColor White
        Write-Host '   [2]  Verify only      (read-only relational integrity report)'      -ForegroundColor White
        Write-Host '   [3]  Repair only      (auto-zero lookup orphans, in-place + Verify)' -ForegroundColor White
        Write-Host '   [4]  Cleanup target   (TRUNCATE migrated tables — staging only)'    -ForegroundColor DarkYellow
        Write-Host '   [5]  Migrate (DryRun) (read source, build maps, no writes)'         -ForegroundColor DarkGray
        Write-Host '   [Q]  Quit'                                                          -ForegroundColor DarkGray
        Write-Host ''
        $sel = Read-Host '   Select'
        switch ($sel) {
            '1' { return @{ Mode = 'Migrate'; DryRun = $false } }
            '2' { return @{ Mode = 'Verify';  DryRun = $false } }
            '3' { return @{ Mode = 'Repair';  DryRun = $false } }
            '4' { return @{ Mode = 'Cleanup'; DryRun = $false } }
            '5' { return @{ Mode = 'Migrate'; DryRun = $true  } }
            'Q' { return @{ Mode = 'Quit';    DryRun = $false } }
            'q' { return @{ Mode = 'Quit';    DryRun = $false } }
            default { Write-Warn2 ("Unknown selection '{0}'." -f $sel); Start-Sleep -Milliseconds 800 }
        }
    }
}

#endregion

#region Module bootstrap -------------------------------------------------------

function Initialize-Dependencies {
    [CmdletBinding()] param()
    Write-Section 'Bootstrapping dependencies'
    if (-not (Get-Module -ListAvailable -Name SimplySql)) {
        Write-Warn2 'SimplySql module not found. Attempting per-user install (PSGallery)...'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
            }
            if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            }
            Install-Module -Name SimplySql -Scope CurrentUser -Force -AllowClobber
        } catch {
            throw "Cannot install SimplySql module. Install it manually with 'Install-Module SimplySql' and retry. Inner: $($_.Exception.Message)"
        }
    }
    Import-Module SimplySql -ErrorAction Stop
    Write-OK ("SimplySql v{0} loaded." -f (Get-Module SimplySql).Version)
}

#endregion

#region Connection management -------------------------------------------------

function Open-Connections {
    [CmdletBinding()] param()
    Write-Section 'Opening database connections'
    try {
        Open-SqlConnection -ConnectionName 'src' -ConnectionString $SourceConnectionString -ErrorAction Stop
        $sqlVer = Invoke-SqlScalar -ConnectionName 'src' -Query 'SELECT @@VERSION'
        if ([string]::IsNullOrEmpty([string]$sqlVer)) { throw 'SQL Server returned an empty version string.' }
        Write-OK ("SQL Server connected: {0}" -f (([string]$sqlVer) -split "`n")[0].Trim())
    } catch {
        throw "Cannot connect to source SQL Server. $($_.Exception.Message)"
    }
    try {
        Open-MySqlConnection -ConnectionName 'tgt' -ConnectionString $TargetConnectionString -ErrorAction Stop
        $mariaVer = Invoke-SqlScalar -ConnectionName 'tgt' -Query 'SELECT VERSION()'
        Write-OK ("MariaDB connected:    {0}" -f $mariaVer)
        # Bulk-import session optimisations — restored in Close-Connections
        if (-not $script:DryRun) {
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION foreign_key_checks = 0' | Out-Null
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION unique_checks      = 0' | Out-Null
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION sql_log_bin        = 0' | Out-Null
            # innodb_flush_log_at_trx_commit is GLOBAL-only in this MariaDB version
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET GLOBAL  innodb_flush_log_at_trx_commit = 0' | Out-Null
            Write-OK 'MariaDB bulk-import flags set (fk_checks=0, unique=0, binlog=0, innodb_flush=GLOBAL 0).'
        }
    } catch {
        throw "Cannot connect to target MariaDB. $($_.Exception.Message)"
    }
}

function Close-Connections {
    try {
        if (Test-SqlConnection -ConnectionName 'tgt') {
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION foreign_key_checks = 1' | Out-Null
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION unique_checks      = 1' | Out-Null
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET SESSION sql_log_bin        = 1' | Out-Null
            Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET GLOBAL  innodb_flush_log_at_trx_commit = 1' | Out-Null
        }
    } catch { }
    foreach ($n in 'src','tgt') {
        try { if (Test-SqlConnection -ConnectionName $n) { Close-SqlConnection -ConnectionName $n } } catch { }
    }
}

#endregion

#region Pre-flight backup -----------------------------------------------------

function Invoke-MariaBackup {
    [CmdletBinding()] param()
    if ($SkipBackup) { Write-Warn2 'Backup skipped on request (-SkipBackup).'; return }
    Write-Section 'Pre-flight: mysqldump backup of simplimed'
    $dump    = $null
    $dumpCmd = Get-Command mysqldump.exe -ErrorAction SilentlyContinue
    if ($dumpCmd) { $dump = $dumpCmd.Source }
    if (-not $dump) {
        Write-Warn2 'mysqldump.exe not found on PATH. Falling back to a SQL-level table-by-table backup...'
        Backup-MariaTablesViaSql
        return
    }
    if (-not $script:DryRun) {
        $args = @(
            '--host=192.168.198.1','--port=3306','--user=root','--password=!nic7774',
            '--single-transaction','--routines','--triggers','--events',
            '--default-character-set=utf8mb4','simplimed'
        )
        Write-Host (' Running mysqldump -> {0}' -f $script:BackupPath) -ForegroundColor DarkCyan
        & $dump @args 2>$null | Out-File -FilePath $script:BackupPath -Encoding UTF8
        if ($LASTEXITCODE -ne 0) {
            throw "mysqldump exited with code $LASTEXITCODE."
        }
        Write-OK ("Backup written: {0} ({1:N0} bytes)" -f $script:BackupPath, (Get-Item $script:BackupPath).Length)
    }
}

function Backup-MariaTablesViaSql {
    # Last-resort logical snapshot when mysqldump is missing. We capture the schema
    # and a JSON dump of any non-empty rows so a restore script can be derived.
    $tables = @('tbl_tenant_entities','tbl_entity_main','tbl_contact_main','tbl_appoint_main',
                'tbl_documentation_main','tbl_receipt_content','tbl_receipt_main',
                'tbl_bookkeeping','tbl_appoint_protocoll')
    $obj = [ordered]@{}
    foreach ($t in $tables) {
        try {
            $rows = Invoke-SqlQuery -ConnectionName 'tgt' -Query "SELECT * FROM $t"
            $obj[$t] = $rows
        } catch {
            $obj[$t] = @{ error = $_.Exception.Message }
        }
    }
    $jsonPath = [IO.Path]::ChangeExtension($script:BackupPath, '.json')
    $obj | ConvertTo-Json -Depth 6 -Compress | Out-File -FilePath $jsonPath -Encoding UTF8
    Write-OK ("Fallback JSON backup written: {0}" -f $jsonPath)
    $script:BackupPath = $jsonPath
}

#endregion

#region Helper utilities ------------------------------------------------------

function ConvertTo-DbValue {
    <#
        Normalises a value coming out of SqlDataReader into something the
        MySql ADO.NET driver will accept. PowerShell-shaped DBNulls become $null
        (which the driver translates to NULL); empty strings stay as empty
        strings; booleans become 0/1 because the target uses bit(1); dates are
        passed through as DateTime.
    #>
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $null }
    if ($Value -is [bool])      { if ($Value) { return 1 } else { return 0 } }
    return $Value
}

function ConvertTo-DbString {
    # Returns a string safely truncated to MaxLen characters.
    # If source is DBNull/null and -Required is set, returns '' (for NOT NULL columns).
    # Otherwise returns $null for DBNull/null.
    param($Value, [int]$MaxLen, [switch]$Required)
    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        if ($Required) { return '' }
        return $null
    }
    $s = [string]$Value
    if ($s.Length -gt $MaxLen) { return $s.Substring(0, $MaxLen) }
    return $s
}

function ConvertTo-SmallInt {
    # Clamp into MariaDB signed smallint range (-32768..32767).
    # DBNull/null becomes 0.
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 0 }
    $n = 0
    if (-not [int]::TryParse([string]$Value, [ref]$n)) { return 0 }
    if ($n -gt 32767)   { return 32767 }
    if ($n -lt -32768)  { return -32768 }
    return $n
}

function ConvertTo-USmallInt {
    # Clamp into MariaDB unsigned smallint range (0..65535).
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 0 }
    $n = 0
    if (-not [int]::TryParse([string]$Value, [ref]$n)) { return 0 }
    if ($n -gt 65535) { return 65535 }
    if ($n -lt 0)     { return 0 }
    return $n
}

function ConvertTo-UTinyInt {
    # Clamp into MariaDB unsigned tinyint range (0..255).
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 0 }
    $n = 0
    if (-not [int]::TryParse([string]$Value, [ref]$n)) { return 0 }
    if ($n -gt 255) { return 255 }
    if ($n -lt 0)   { return 0 }
    return $n
}

function ConvertTo-BitInt {
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 0 }
    if ($Value -is [bool]) { if ($Value) { return 1 } else { return 0 } }
    if ($Value -is [int] -or $Value -is [int16]) { if ($Value -ne 0) { return 1 } else { return 0 } }
    if ([int]::TryParse([string]$Value, [ref]([int]0))) { if ([int]$Value -ne 0) { return 1 } else { return 0 } }
    return 0
}

function Get-MappedId {
    param(
        [Parameter(Mandatory)] [hashtable] $Map,
        $Key,
        [int] $Default = 0
    )
    if ($null -eq $Key -or $Key -is [System.DBNull]) { return $Default }
    if (-not ($Key -is [int] -or $Key -is [long])) {
        if (-not [int]::TryParse([string]$Key, [ref]([int]0))) { return $Default }
        $Key = [int]$Key
    }
    if ($Key -le 0) { return $Default }
    if ($Map.ContainsKey($Key)) { return [int]$Map[$Key] }
    return $Default
}

# ---------------------------------------------------------------------------
# Lookup-master sets — pre-loaded once at start so each row insert can
# validate its raw foreign-key value (1:1 copied from source) against the
# actual simplimed master tables without per-row round-trips. If a referenced
# master id does not exist, Resolve-LookupId returns 0 (= sentinel "no
# reference") so the target row stays consistent and the FK never points into
# nothing. We also track distinct dropped values for audit logging.
$script:LookupSets       = @{}
$script:LookupDropped    = @{}   # 'tbl_xxx' -> HashSet[int] of source-ids zeroed

function Initialize-LookupSets {
    [CmdletBinding()] param()
    Write-Section 'Pre-loading lookup master tables for FK validation'
    $masters = @(
        'tbl_currencies',
        'tbl_contact_payments',
        'tbl_receipt_dispatch',
        'tbl_services_main',
        'tbl_contact_marital',
        'tbl_contact_gender',
        'tbl_contact_insurance',
        'tbl_receipt_type',
        'tbl_appoint_location',
        'tbl_appoint_status',
        'tbl_appoint_priority',
        'tbl_documentation_type',
        'tbl_banking'
    )
    foreach ($t in $masters) {
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
        try {
            $cnt = [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query "SELECT COUNT(*) FROM $t")
            if ($cnt -gt 0) {
                $rows = Invoke-SqlQuery -ConnectionName 'tgt' -Query "SELECT fld_id FROM $t" -WarningAction SilentlyContinue
                foreach ($r in $rows) { [void]$set.Add([int]$r.fld_id) }
            }
            Write-Host (' [ OK ] {0,-28} {1,5} master rows' -f $t, $set.Count) -ForegroundColor DarkGray
        } catch {
            Write-Warn2 ("Master table '{0}' not loadable ({1}); FKs to it will pass-through unvalidated." -f $t, $_.Exception.Message)
        }
        $script:LookupSets[$t]    = $set
        $script:LookupDropped[$t] = New-Object 'System.Collections.Generic.HashSet[int]'
    }
}

function Resolve-LookupId {
    # Returns the raw value if the corresponding fld_id exists in the master
    # table, otherwise 0 (sentinel "no reference"). Tracks distinct dropped
    # ids per master table for the post-run audit log. Master tables that
    # could not be loaded simply pass through.
    param([string] $Master, $Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 0 }
    $n = 0
    if (-not [int]::TryParse([string]$Value, [ref]$n)) { return 0 }
    if ($n -le 0) { return 0 }
    $set = $script:LookupSets[$Master]
    if ($null -eq $set) { return $n }   # master not loadable -> pass through
    if ($set.Count -eq 0) {
        # Empty master -> remember dropped id, return 0
        [void]$script:LookupDropped[$Master].Add($n)
        return 0
    }
    if ($set.Contains($n)) { return $n }
    [void]$script:LookupDropped[$Master].Add($n)
    return 0
}

function New-Guid2 { return [System.Guid]::NewGuid().ToString() }

function Add-Warning {
    param([string]$Table, [string]$Reason, $SourceKey)
    $script:WarnCollection.Add([pscustomobject]@{
        Table = $Table; Reason = $Reason; Key = $SourceKey
    })
}

function Add-MigError {
    param([string]$Table, [string]$Reason, $SourceKey, [System.Exception]$Exception)
    $script:ErrorCollection.Add([pscustomobject]@{
        Table = $Table; Reason = $Reason; Key = $SourceKey
        Message = if ($Exception) { $Exception.Message } else { '' }
    })
}

function Initialize-Stats {
    param([string]$Table)
    if (-not $script:Stats.Contains($Table)) {
        $script:Stats[$Table] = [pscustomobject]@{
            Table = $Table; Read = 0; Inserted = 0; Skipped = 0; Warnings = 0; Errors = 0
        }
    }
}

function Update-Stats {
    param(
        [string]$Table,
        [int]$Read = 0, [int]$Inserted = 0, [int]$Skipped = 0,
        [int]$Warnings = 0, [int]$Errors = 0
    )
    Initialize-Stats -Table $Table
    $s = $script:Stats[$Table]
    $s.Read     += $Read
    $s.Inserted += $Inserted
    $s.Skipped  += $Skipped
    $s.Warnings += $Warnings
    $s.Errors   += $Errors
}

# Batch-transaction state — one open transaction is held across BatchSize inserts
# to amortise per-commit overhead.  Commit-PendingBatch must be called at the
# end of each step function and in the finally block.
$script:TxOpen      = $false
$script:TxBatchCnt  = 0

function Open-InsertBatch {
    if (-not $script:DryRun -and -not $script:TxOpen) {
        Start-SqlTransaction -ConnectionName 'tgt' | Out-Null
        $script:TxOpen = $true
    }
}

function Commit-InsertBatch {
    param([switch]$Force)
    if ($script:DryRun) { return }
    if ($script:TxOpen -and ($Force -or $script:TxBatchCnt -ge $BatchSize)) {
        Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null
        $script:TxOpen = $false; $script:TxBatchCnt = 0
        Start-SqlTransaction -ConnectionName 'tgt' | Out-Null
        $script:TxOpen = $true
    }
}

function Commit-PendingBatch {
    if ($script:DryRun -or -not $script:TxOpen) { return }
    Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null
    $script:TxOpen = $false; $script:TxBatchCnt = 0
}

function Invoke-TargetInsertReturningId {
    # Uses MariaDB's INSERT ... RETURNING fld_id to get the new PK in a single
    # round-trip instead of the old INSERT + SELECT LAST_INSERT_ID() pattern.
    param(
        [Parameter(Mandatory)] [string]    $Sql,
        [Parameter(Mandatory)] [hashtable] $Parameters
    )
    $sqlR    = ($Sql.TrimEnd()) + ' RETURNING fld_id'
    $attempt = 0
    while ($true) {
        try {
            $row = Invoke-SqlQuery -ConnectionName 'tgt' -Query $sqlR -Parameters $Parameters -WhatIf:$false
            $script:TxBatchCnt++
            Commit-InsertBatch
            return [int]$row.fld_id
        } catch {
            $attempt++
            if ($attempt -ge 3) { throw }
            Start-Sleep -Milliseconds (200 * [math]::Pow(2, $attempt))
        }
    }
}

function Get-SourceReader {
    <#
        Returns an open SqlDataReader against a *dedicated* SqlConnection so we
        do not materialise large result sets in memory and do not interfere
        with SimplySql's primary connection (which we still use for the
        scalar count queries). The caller must call Close-SourceReader.
    #>
    param([Parameter(Mandatory)] [string] $Query)

    Add-Type -AssemblyName System.Data -ErrorAction SilentlyContinue
    $conn = New-Object System.Data.SqlClient.SqlConnection $SourceConnectionString
    $conn.Open()
    $cmd                = $conn.CreateCommand()
    $cmd.CommandText    = $Query
    $cmd.CommandTimeout = 0
    $reader             = $cmd.ExecuteReader()
    return [pscustomobject]@{ Reader = $reader; Command = $cmd; Connection = $conn }
}

function Close-SourceReader {
    param([Parameter(Mandatory)] [psobject] $H)
    try { if ($H.Reader -and -not $H.Reader.IsClosed) { $H.Reader.Close() } } catch { }
    try { if ($H.Command) { $H.Command.Dispose() } } catch { }
    try { if ($H.Connection) { $H.Connection.Close(); $H.Connection.Dispose() } } catch { }
}

# NOTE: We intentionally read the SqlDataReader inline in each step function.
# Wrapping the reader access in a function or script block exhibits a PS 5.1
# quirk where the Reader object's underlying state machine appears to skip the
# first record on second-and-subsequent invocations (likely an interaction with
# how PowerShell resolves PSCustomObject property bags vs. IEnumerable). The
# inline pattern is verbose but reliable.

#endregion

#region Migration steps -------------------------------------------------------

function Step1_Migrate-Tenants {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_tenant_entities'
    $tableSrc = 'Tabelle_Patienten'
    Write-Section ("Step 1/9: {0}  ->  {1}  (VIP=1, Passiv=0)" -f $tableSrc, $tableTgt)
    Initialize-Stats -Table $tableTgt

    $countQuery = "SELECT COUNT(*) FROM [Tabelle_Patienten] WHERE VIP=1 AND Passiv=0"
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'No source tenants match the filter.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID0, GuiID, Anrede, Titel, Name, Vorname, [Straße] AS Strasse, PLZ, Ort, Land,
       R_Firma1, Bemerkung, Passiv
FROM [Tabelle_Patienten]
WHERE VIP=1 AND Passiv=0
ORDER BY ID0
"@
    Open-InsertBatch
    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        $i = 0
        $insertSql = @"
INSERT INTO tbl_tenant_entities
  (fld_guiid, fld_salutation, fld_title, fld_surname, fld_firstname,
   fld_street, fld_zipcode, fld_city, fld_country,
   fld_organisation, fld_comment, fld_passive)
VALUES
  (@guiid, @sal, @tit, @sur, @fir,
   @str, @zip, @cit, @cou,
   @org, @com, @pas)
"@
        while ($rdr.Read()) {
            $i++
            $row = [ordered]@{}
            for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            Update-Stats -Table $tableTgt -Read 1
            $oldId = [int]$row.ID0
            $params = @{
                guiid = (New-Guid2)
                sal   = ConvertTo-DbString $row.Anrede   -MaxLen 50
                tit   = ConvertTo-DbString $row.Titel    -MaxLen 100
                sur   = ConvertTo-DbString $row.Name     -MaxLen 200
                fir   = ConvertTo-DbString $row.Vorname  -MaxLen 200
                str   = ConvertTo-DbString $row.Strasse  -MaxLen 200
                zip   = ConvertTo-DbString $row.PLZ      -MaxLen 50
                cit   = ConvertTo-DbString $row.Ort      -MaxLen 250
                cou   = ConvertTo-DbString $row.Land     -MaxLen 150
                org   = ConvertTo-DbString $row.R_Firma1 -MaxLen 150 -Required
                com   = ConvertTo-DbString $row.Bemerkung -MaxLen 800
                pas   = ConvertTo-BitInt $row.Passiv
            }
            try {
                if (-not $script:DryRun) {
                    $newId = Invoke-TargetInsertReturningId -Sql $insertSql -Parameters $params
                    $script:Maps.Tenant[$oldId] = $newId
                    if ($script:DefaultTenantId -eq 0) { $script:DefaultTenantId = $newId }
                    Update-Stats -Table $tableTgt -Inserted 1
                } else {
                    # WhatIf: still build a fake mapping so downstream FK lookups exist
                    $script:Maps.Tenant[$oldId] = -1 * $oldId
                    if ($script:DefaultTenantId -eq 0) { $script:DefaultTenantId = -1 * $oldId }
                    Update-Stats -Table $tableTgt -Inserted 1
                }
            } catch {
                Add-MigError -Table $tableTgt -Reason 'INSERT failed' -SourceKey $oldId -Exception $_.Exception
                Update-Stats -Table $tableTgt -Errors 1
            }
            if (($i % 25) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
    } finally {
        Close-SourceReader $h
        Commit-PendingBatch
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} tenants mapped." -f $script:Maps.Tenant.Count)
}

function Step2_Migrate-Entities {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_entity_main'
    $tableSrc = 'Tabelle_Patienten'
    Write-Section ("Step 2/9: {0}  ->  {1}  (VIP=0, Mitarbeiter=1, Passiv=0)" -f $tableSrc, $tableTgt)
    Initialize-Stats -Table $tableTgt

    $countQuery = "SELECT COUNT(*) FROM [Tabelle_Patienten] WHERE VIP=0 AND Mitarbeiter=1 AND Passiv=0"
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'No source employees match the filter.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID0, GuiID, IDP, Name, Vorname, Ort, R_Firma1,
       Telefon5, Telefon6, Telefon1, Telefon2, Telefon4, Internet, Firma2,
       Grafik, Em_User, Em_Pass, Kontoinhaber, IBAN, BIC, Gesperrt, Passiv
FROM [Tabelle_Patienten]
WHERE VIP=0 AND Mitarbeiter=1 AND Passiv=0
ORDER BY ID0
"@
    Open-InsertBatch
    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        $i = 0
        $insertSql = @"
INSERT INTO tbl_entity_main
  (fld_guiid, fld_tenant_id, fld_colorcode,
   fld_surname, fld_firstname, fld_city, fld_organisation,
   fld_email_private, fld_email_busines,
   fld_phone_home, fld_phone_business, fld_phone_mobile,
   fld_website, fld_caption, fld_username, fld_password, fld_photofile,
   fld_account_holder, fld_IBAN, fld_BIC, fld_hidecalendar, fld_passive)
VALUES
  (@guiid, @ten, '#cccccc',
   @sur, @fir, @cit, @org,
   @ep, @eb,
   @ph, @pb, @pm,
   @web, @cap, @usr, @pwd, @pho,
   @ah, @iban, @bic, @hide, @pas)
"@
        while ($rdr.Read()) {
            $i++
            Update-Stats -Table $tableTgt -Read 1
            $row    = [ordered]@{}; for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            $oldId  = [int]$row.ID0
            $oldTen = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newTen = Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No tenant mapping for IDP=$oldTen" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            $params = @{
                guiid = (New-Guid2)
                ten   = $newTen
                sur   = ConvertTo-DbString $row.Name         -MaxLen 200 -Required
                fir   = ConvertTo-DbString $row.Vorname      -MaxLen 200 -Required
                cit   = ConvertTo-DbString $row.Ort          -MaxLen 250
                org   = ConvertTo-DbString $row.R_Firma1     -MaxLen 150
                ep    = ConvertTo-DbString $row.Telefon5     -MaxLen 100
                eb    = ConvertTo-DbString $row.Telefon6     -MaxLen 100
                ph    = ConvertTo-DbString $row.Telefon1     -MaxLen 100
                pb    = ConvertTo-DbString $row.Telefon2     -MaxLen 100
                pm    = ConvertTo-DbString $row.Telefon4     -MaxLen 100
                web   = ConvertTo-DbString $row.Internet     -MaxLen 100
                cap   = ConvertTo-DbString $row.Firma2       -MaxLen 200 -Required
                usr   = ConvertTo-DbString $row.Em_User      -MaxLen 100 -Required
                pwd   = ConvertTo-DbString $row.Em_Pass      -MaxLen 100 -Required
                pho   = ConvertTo-DbString $row.Grafik       -MaxLen 100
                ah    = ConvertTo-DbString $row.Kontoinhaber -MaxLen 150
                iban  = ConvertTo-DbString $row.IBAN         -MaxLen 30
                bic   = ConvertTo-DbString $row.BIC          -MaxLen 20
                hide  = ConvertTo-BitInt $row.Gesperrt
                pas   = ConvertTo-BitInt $row.Passiv
            }
            try {
                if (-not $script:DryRun) {
                    $newId = Invoke-TargetInsertReturningId -Sql $insertSql -Parameters $params
                    $script:Maps.Entity[$oldId] = $newId
                    Update-Stats -Table $tableTgt -Inserted 1
                } else {
                    $script:Maps.Entity[$oldId] = -1 * $oldId
                    Update-Stats -Table $tableTgt -Inserted 1
                }
            } catch {
                Add-MigError -Table $tableTgt -Reason 'INSERT failed' -SourceKey $oldId -Exception $_.Exception
                Update-Stats -Table $tableTgt -Errors 1
            }
            if (($i % 25) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
    } finally {
        Close-SourceReader $h
        Commit-PendingBatch
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} entities mapped." -f $script:Maps.Entity.Count)
}

function Step3_Migrate-Contacts {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_contact_main'
    Write-Section ("Step 3/9: Tabelle_Patienten  ->  {0}  (VIP=0, Mitarbeiter=0)" -f $tableTgt)
    Initialize-Stats -Table $tableTgt

    $countQuery = "SELECT COUNT(*) FROM [Tabelle_Patienten] WHERE VIP=0 AND Mitarbeiter=0"
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'No source contacts match the filter.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID0, GuiID, IDP, IDZ, [Währung] AS Waehrung, Versand, ID3,
       Familienstand, Geschlecht, Mandant, Versichertenart AS Ableitung, IDKurz,
       Anrede, Titel, Name, Vorname, [Straße] AS Strasse, PLZ, Ort, Land,
       Firma1, Telefon5, Telefon1, Telefon2, Telefon4, Internet, Grafik, Bemerkung,
       R_Briefanrede, Anschrift, Geboren, Datum, [Geändert] AS Geaendert,
       R_Anrede, R_Titel, [R_Straße] AS R_Strasse, R_Vorname, R_Name,
       R_PLZ, R_Ort, R_Land, R_Firma1,
       Kontoinhaber, IBAN, BIC, Beruf, [Größe] AS Groesse, Gewicht,
       Diagnose, Hinweis, Kartennummer, Muttersprache, Kopien, Rabatt,
       Mailing, Passiv
FROM [Tabelle_Patienten]
WHERE VIP=0 AND Mitarbeiter=0
ORDER BY ID0
"@
    Open-InsertBatch
    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        $i = 0
        $insertSql = @"
INSERT INTO tbl_contact_main
  (fld_guiid, fld_tenant_id, fld_payment_id, fld_currencies_id, fld_dispatch_id, fld_service_id,
   fld_marital_id, fld_gender_id, fld_insurance_id, fld_customerid, fld_displayname,
   fld_salutation, fld_title, fld_surname, fld_firstname, fld_street, fld_zipcode, fld_city, fld_country,
   fld_organisation, fld_email_private, fld_email_busines,
   fld_phone_home, fld_phone_business, fld_phone_mobile, fld_website, fld_photofile, fld_comment,
   fld_letter_salutation, fld_address, fld_birthday, fld_date_add, fld_date_edit,
   fld_salutation2, fld_title2, fld_firstname2, fld_surname2, fld_street2, fld_zipcode2, fld_city2, fld_country2, fld_organisation2,
   fld_account_holder, fld_IBAN, fld_BIC, fld_profession, fld_size, fld_weight,
   fld_medical_history, fld_notes, fld_card_number, fld_native_language,
   fld_copies, fld_credit_balance, fld_mailing, fld_passive)
VALUES
  (@guiid, @ten, @pay, @cur, @dis, @ser,
   @mar, @gen, @ins, @cust, @disp,
   @sal, @tit, @sur, @fir, @str, @zip, @cit, @cou,
   @org, @ep, @eb,
   @ph, @pb, @pm, @web, @photo, @com,
   @lsal, @addr, @birth, @dadd, @dedit,
   @sal2, @tit2, @fir2, @sur2, @str2, @zip2, @cit2, @cou2, @org2,
   @ah, @iban, @bic, @prof, @sz, @wt,
   @mh, @notes, @card, @lang,
   @copies, @credit, @mail, @pas)
"@
        $idGenders   = @{}     # map for marital, gender etc. share the new map: keep simple
        while ($rdr.Read()) {
            $i++
            Update-Stats -Table $tableTgt -Read 1
            $row     = [ordered]@{}; for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            $oldId   = [int]$row.ID0
            $oldTen  = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newTen  = Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No tenant mapping for IDP=$oldTen" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            # Lookup FKs: validate against the simplimed master tables loaded
            # by Initialize-LookupSets. Unknown ids are coerced to 0 (= "no
            # reference") so the target never holds a dangling pointer.
            $payId = Resolve-LookupId -Master 'tbl_contact_payments'  -Value $row.IDZ
            $curId = Resolve-LookupId -Master 'tbl_currencies'        -Value $row.Waehrung
            $disId = Resolve-LookupId -Master 'tbl_receipt_dispatch'  -Value $row.Versand
            $serId = Resolve-LookupId -Master 'tbl_services_main'     -Value $row.ID3
            $marId = Resolve-LookupId -Master 'tbl_contact_marital'   -Value $row.Familienstand
            $genId = Resolve-LookupId -Master 'tbl_contact_gender'    -Value $row.Geschlecht
            $insId = Resolve-LookupId -Master 'tbl_contact_insurance' -Value $row.Ableitung

            $now = Get-Date
            $params = @{
                guiid  = (New-Guid2)
                ten    = $newTen
                pay    = $payId
                cur    = $curId
                dis    = $disId
                ser    = $serId
                mar    = $marId
                gen    = $genId
                ins    = $insId
                cust   = ConvertTo-DbString $row.Mandant       -MaxLen 50
                disp   = ConvertTo-DbString $row.IDKurz        -MaxLen 150
                sal    = ConvertTo-DbString $row.Anrede        -MaxLen 50
                tit    = ConvertTo-DbString $row.Titel         -MaxLen 100
                sur    = ConvertTo-DbString $row.Name          -MaxLen 200
                fir    = ConvertTo-DbString $row.Vorname       -MaxLen 200
                str    = ConvertTo-DbString $row.Strasse       -MaxLen 200
                zip    = ConvertTo-DbString $row.PLZ           -MaxLen 50
                cit    = ConvertTo-DbString $row.Ort           -MaxLen 250
                cou    = ConvertTo-DbString $row.Land          -MaxLen 150
                org    = ConvertTo-DbString $row.Firma1        -MaxLen 150
                ep     = ConvertTo-DbString $row.Telefon5      -MaxLen 100
                eb     = ConvertTo-DbString $row.Telefon5      -MaxLen 100
                ph     = ConvertTo-DbString $row.Telefon1      -MaxLen 100
                pb     = ConvertTo-DbString $row.Telefon2      -MaxLen 100
                pm     = ConvertTo-DbString $row.Telefon4      -MaxLen 100
                web    = ConvertTo-DbString $row.Internet      -MaxLen 100
                photo  = ConvertTo-DbString $row.Grafik        -MaxLen 100
                com    = ConvertTo-DbString $row.Bemerkung     -MaxLen 800
                lsal   = ConvertTo-DbString $row.R_Briefanrede -MaxLen 100
                addr   = ConvertTo-DbString $row.Anschrift     -MaxLen 200
                birth  = if ($row.Geboren -is [System.DBNull]) { $null } else { ([datetime]$row.Geboren).Date }
                dadd   = if ($row.Datum   -is [System.DBNull]) { $now }  else { [datetime]$row.Datum }
                dedit  = if ($row.Geaendert -is [System.DBNull]) { $now } else { [datetime]$row.Geaendert }
                sal2   = ConvertTo-DbString $row.R_Anrede      -MaxLen 50
                tit2   = ConvertTo-DbString $row.R_Titel       -MaxLen 100
                fir2   = ConvertTo-DbString $row.R_Strasse     -MaxLen 200
                sur2   = ConvertTo-DbString $row.R_Vorname     -MaxLen 200
                str2   = ConvertTo-DbString $row.R_Name        -MaxLen 200
                zip2   = ConvertTo-DbString $row.R_PLZ         -MaxLen 50
                cit2   = ConvertTo-DbString $row.R_Ort         -MaxLen 250
                cou2   = ConvertTo-DbString $row.R_Land        -MaxLen 150
                org2   = ConvertTo-DbString $row.R_Firma1      -MaxLen 150
                ah     = ConvertTo-DbString $row.Kontoinhaber  -MaxLen 150
                iban   = ConvertTo-DbString $row.IBAN          -MaxLen 30
                bic    = ConvertTo-DbString $row.BIC           -MaxLen 20
                prof   = ConvertTo-DbString $row.Beruf         -MaxLen 100
                sz     = ConvertTo-DbString $row.Groesse       -MaxLen 20
                wt     = ConvertTo-DbString $row.Gewicht       -MaxLen 20
                mh     = ConvertTo-DbString $row.Diagnose      -MaxLen 2000
                notes  = ConvertTo-DbString $row.Hinweis       -MaxLen 2000
                card   = ConvertTo-DbString $row.Kartennummer  -MaxLen 50
                lang   = ConvertTo-DbString $row.Muttersprache -MaxLen 50
                copies = if ($row.Kopien -is [System.DBNull]) { 1 } else { [int]$row.Kopien }
                credit = if ($row.Rabatt -is [System.DBNull]) { 0.0 } else { [decimal]$row.Rabatt }
                mail   = ConvertTo-BitInt $row.Mailing
                pas    = ConvertTo-BitInt $row.Passiv
            }
            try {
                if (-not $script:DryRun) {
                    $newId = Invoke-TargetInsertReturningId -Sql $insertSql -Parameters $params
                    $script:Maps.Contact[$oldId] = $newId
                    Update-Stats -Table $tableTgt -Inserted 1
                } else {
                    $script:Maps.Contact[$oldId] = -1 * $oldId
                    Update-Stats -Table $tableTgt -Inserted 1
                }
            } catch {
                Add-MigError -Table $tableTgt -Reason 'INSERT failed' -SourceKey $oldId -Exception $_.Exception
                Update-Stats -Table $tableTgt -Errors 1
            }
            if (($i % 100) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
    } finally {
        Close-SourceReader $h
        Commit-PendingBatch
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} contacts mapped." -f $script:Maps.Contact.Count)
}

function Step4_Migrate-Receipts {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_receipt_main'
    Write-Section "Step 4/9: Tabelle_AbreRe  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_AbreRe]'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'Tabelle_AbreRe is empty.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID1, GuiID, ID0, IDP, IDZ, [Währung] AS Waehrung, Versand, [Type] AS Typ,
       GutNr, RechNr, GuStr, Datum, Fallig, Berichtdatum,
       GesBetrag, Bezahlt, Rabatt, Steuer, Kopie, Selekt, Storniert
FROM [Tabelle_AbreRe]
ORDER BY ID1
"@
    Open-InsertBatch
    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        $i = 0
        $insertSql = @"
INSERT INTO tbl_receipt_main
  (fld_guiid, fld_contact_id, fld_tenant_id, fld_payment_id, fld_currencies_id, fld_dispatch_id, fld_type_id,
   fld_credit_note_id, fld_receipt_number, fld_credit_note_number,
   fld_receipt_date, fld_due_date, fld_date_add, fld_date_changed, fld_date_published,
   fld_receipt_amount, fld_amount_paid, fld_credit_balance, fld_tax_rate,
   fld_copies, fld_completed, fld_passive)
VALUES
  (@guiid, @con, @ten, @pay, @cur, @dis, @typ,
   @cnId, @rno, @cstr,
   @rdate, @ddate, @dadd, @dch, @dpub,
   @amt, @paid, @cb, @tx,
   @copies, @done, @pas)
"@
        while ($rdr.Read()) {
            $i++
            Update-Stats -Table $tableTgt -Read 1
            $row    = [ordered]@{}; for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            $oldId  = [int]$row.ID1
            $oldTen = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newTen = Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No tenant mapping for IDP=$oldTen" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            $oldCon = if ($row.ID0 -is [System.DBNull]) { 0 } else { [int]$row.ID0 }
            $newCon = Get-MappedId -Map $script:Maps.Contact -Key $oldCon -Default 0
            if ($newCon -eq 0 -and $oldCon -ne 0) {
                Add-Warning -Table $tableTgt -Reason "Receipt contact ID0=$oldCon not in contact map (orphan FK)" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1
            }
            $payId = Resolve-LookupId -Master 'tbl_contact_payments' -Value $row.IDZ
            $curId = Resolve-LookupId -Master 'tbl_currencies'       -Value $row.Waehrung
            $disId = Resolve-LookupId -Master 'tbl_receipt_dispatch' -Value $row.Versand
            $typId = Resolve-LookupId -Master 'tbl_receipt_type'     -Value $row.Typ
            # fld_credit_note_id references another receipt — that lives in the
            # script's Receipt map (post-step-4 it does not yet exist for the
            # current row). Keep the source value, will be validated in Step10.
            $cnId  = if ($row.GutNr -is [System.DBNull]) { 0 } else { [math]::Max(0,[int]$row.GutNr) }

            $now = Get-Date
            $rdate = if ($row.Datum   -is [System.DBNull]) { $now.Date } else { ([datetime]$row.Datum).Date }
            $ddate = if ($row.Fallig  -is [System.DBNull]) { $rdate    } else { ([datetime]$row.Fallig).Date }
            $dpub  = if ($row.Berichtdatum -is [System.DBNull]) { $now } else { [datetime]$row.Berichtdatum }
            $params = @{
                guiid = (New-Guid2)
                con   = $newCon
                ten   = $newTen
                pay   = $payId
                cur   = $curId
                dis   = $disId
                typ   = $typId
                cnId  = $cnId
                rno   = if ($row.RechNr -is [System.DBNull] -or [string]::IsNullOrEmpty([string]$row.RechNr)) { '0' } else { $tmp=[string]$row.RechNr; if($tmp.Length -gt 25){$tmp.Substring(0,25)}else{$tmp} }
                cstr  = if ($row.GuStr  -is [System.DBNull] -or [string]::IsNullOrEmpty([string]$row.GuStr))  { '0' } else { $tmp=[string]$row.GuStr;  if($tmp.Length -gt 25){$tmp.Substring(0,25)}else{$tmp} }
                rdate = $rdate
                ddate = $ddate
                dadd  = $now
                dch   = $now
                dpub  = $dpub
                amt   = if ($row.GesBetrag -is [System.DBNull]) { 0.0 } else { [decimal]$row.GesBetrag }
                paid  = if ($row.Bezahlt   -is [System.DBNull]) { 0.0 } else { [decimal]$row.Bezahlt }
                cb    = if ($row.Rabatt    -is [System.DBNull]) { 0.0 } else { [decimal]$row.Rabatt }
                tx    = if ($row.Steuer    -is [System.DBNull]) { 0.0 } else { [decimal]$row.Steuer }
                copies= if ($row.Kopie -is [System.DBNull]) { 1 } else { ConvertTo-SmallInt $row.Kopie }
                done  = ConvertTo-BitInt $row.Selekt
                pas   = ConvertTo-BitInt $row.Storniert
            }
            try {
                if (-not $script:DryRun) {
                    $newId = Invoke-TargetInsertReturningId -Sql $insertSql -Parameters $params
                    $script:Maps.Receipt[$oldId] = $newId
                    Update-Stats -Table $tableTgt -Inserted 1
                } else {
                    $script:Maps.Receipt[$oldId] = -1 * $oldId
                    Update-Stats -Table $tableTgt -Inserted 1
                }
            } catch {
                Add-MigError -Table $tableTgt -Reason 'INSERT failed' -SourceKey $oldId -Exception $_.Exception
                Update-Stats -Table $tableTgt -Errors 1
            }
            if (($i % 50) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
    } finally {
        Close-SourceReader $h
        Commit-PendingBatch
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} receipts mapped." -f $script:Maps.Receipt.Count)
}

function Step5_Migrate-Appointments {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_appoint_main'
    Write-Section "Step 5/9: Tabelle_PatientenWv  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_PatientenWv]'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'Tabelle_PatientenWv is empty.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID2, GuiID, DAVID, IDP, IDM, ID0, IDR,
       Farbtyp, [Priorität] AS Prio, MasTer,
       Datum, Change, VonDat, BisDat, DAVDate, OnlBook,
       Farbe, IDKurz, Kommentar, DAVChange, Erinnerung, OnlTer, Erledigt, Passiv
FROM [Tabelle_PatientenWv]
ORDER BY ID2
"@
    Open-InsertBatch
    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        $i = 0
        $insertSql = @"
INSERT INTO tbl_appoint_main
  (fld_guiid, fld_dav_guiid, fld_tenant_id, fld_entity_id, fld_contact_id, fld_location_id,
   fld_status_id, fld_priority_id, fld_serial_number,
   fld_date_add, fld_date_edit, fld_date_start, fld_date_end, fld_dav_date, fld_bookdate,
   fld_colorcode, fld_subject, fld_comment, fld_email,
   fld_dav_change, fld_reminder, fld_online, fld_finished, fld_passive)
VALUES
  (@guiid, @dav, @ten, @ent, @con, @loc,
   @sta, @pri, @serial,
   @dadd, @dedit, @dstart, @dend, @dav_dt, @book,
   @color, @subj, @com, '',
   @davchg, @rem, @onl, @fin, @pas)
"@
        while ($rdr.Read()) {
            $i++
            Update-Stats -Table $tableTgt -Read 1
            $row    = [ordered]@{}; for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            $oldId  = [int]$row.ID2
            $oldTen = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newTen = Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No tenant mapping for IDP=$oldTen" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            $oldEnt = if ($row.IDM -is [System.DBNull]) { 0 } else { [int]$row.IDM }
            $newEnt = Get-MappedId -Map $script:Maps.Entity  -Key $oldEnt -Default 0
            $oldCon = if ($row.ID0 -is [System.DBNull]) { 0 } else { [int]$row.ID0 }
            $newCon = Get-MappedId -Map $script:Maps.Contact -Key $oldCon -Default 0
            # Lookup FKs validated against simplimed master tables.
            $newLoc = Resolve-LookupId -Master 'tbl_appoint_location' -Value $row.IDR
            $rawSta = if ($row.Farbtyp -is [System.DBNull]) { 1 } else { [math]::Max(1,[int]$row.Farbtyp) }
            $newSta = Resolve-LookupId -Master 'tbl_appoint_status'   -Value $rawSta
            if ($newSta -eq 0) { $newSta = 1 }   # status NOT NULL DEFAULT 1
            $rawPri = if ($row.Prio -is [System.DBNull]) { 2 } else { [math]::Max(1,[int]$row.Prio) }
            $newPri = Resolve-LookupId -Master 'tbl_appoint_priority' -Value $rawPri
            if ($newPri -eq 0) { $newPri = 2 }   # priority NOT NULL DEFAULT 2

            $now    = Get-Date
            $params = @{
                guiid  = (New-Guid2)
                dav    = (New-Guid2)
                ten    = $newTen
                ent    = $newEnt
                con    = $newCon
                loc    = $newLoc
                sta    = $newSta
                pri    = $newPri
                serial = if ($row.MasTer -is [System.DBNull]) { 0 } else { [math]::Max(0,[int]$row.MasTer) }
                dadd   = if ($row.Datum    -is [System.DBNull]) { $now } else { [datetime]$row.Datum }
                dedit  = if ($row.Change   -is [System.DBNull]) { $now } else { [datetime]$row.Change }
                dstart = if ($row.VonDat   -is [System.DBNull]) { $now } else { [datetime]$row.VonDat }
                dend   = if ($row.BisDat   -is [System.DBNull]) { $now } else { [datetime]$row.BisDat }
                dav_dt = if ($row.DAVDate  -is [System.DBNull]) { $null } else { [datetime]$row.DAVDate }
                book   = if ($row.OnlBook  -is [System.DBNull]) { $null } else { [datetime]$row.OnlBook }
                color  = if ([string]::IsNullOrEmpty([string](ConvertTo-DbValue $row.Farbe))) { '#cccccc' } else {
                            $tmp = [string]$row.Farbe
                            if ($tmp -notmatch '^#') { $tmp = '#' + $tmp }
                            if ($tmp.Length -ne 7) { '#cccccc' } else { $tmp }
                         }
                subj   = ConvertTo-DbString $row.IDKurz    -MaxLen 250
                com    = ConvertTo-DbString $row.Kommentar -MaxLen 800
                davchg = ConvertTo-BitInt $row.DAVChange
                rem    = ConvertTo-BitInt $row.Erinnerung
                onl    = ConvertTo-BitInt $row.OnlTer
                fin    = ConvertTo-BitInt $row.Erledigt
                pas    = ConvertTo-BitInt $row.Passiv
            }
            try {
                if (-not $script:DryRun) {
                    $newId = Invoke-TargetInsertReturningId -Sql $insertSql -Parameters $params
                    $script:Maps.Appoint[$oldId] = $newId
                    Update-Stats -Table $tableTgt -Inserted 1
                } else {
                    $script:Maps.Appoint[$oldId] = -1 * $oldId
                    Update-Stats -Table $tableTgt -Inserted 1
                }
            } catch {
                Add-MigError -Table $tableTgt -Reason 'INSERT failed' -SourceKey $oldId -Exception $_.Exception
                Update-Stats -Table $tableTgt -Errors 1
            }
            if (($i % 200) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
    } finally {
        Close-SourceReader $h
        Commit-PendingBatch
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} appointments mapped." -f $script:Maps.Appoint.Count)
}

function Step6_Migrate-Documentation {
    [CmdletBinding()] param()
    $tableTgt  = 'tbl_documentation_main'
    $batchMax  = 150  # 150 rows × 18 params = 2700
    Write-Section "Step 6/9: Tabelle_Abre (ID1<10)  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_Abre] WHERE ID1 < 10'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'No Tabelle_Abre rows with ID1<10.'; return }

    $top           = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $longTextIdSet = @(21..30) + @(103,107,108)
    $query = @"
SELECT $($top)ID2, IDP, IDM, ID1, ID0,
       Sorter, Datum, GONr, Kommentar, x AS XCnt, Druckdatum,
       Multi, Betrag, GesBetrag, Steuer, Lock, Gedruckt, Storniert
FROM [Tabelle_Abre]
WHERE ID1 < 10
ORDER BY ID2
"@
    $cols  = 'fld_guiid,fld_tenant_id,fld_typ_id,fld_contact_id,fld_entity_id,fld_sorter,fld_docdate,fld_identifier,fld_longtext,fld_number,fld_time,fld_multiplier,fld_unit_price,fld_total_price,fld_tax,fld_lock,fld_print,fld_passive'
    $batch = [System.Collections.Generic.List[hashtable]]::new($batchMax)

    function Flush-DocBatch { param($rows)
        if ($rows.Count -eq 0) { return }
        $v = [System.Text.StringBuilder]::new(); $p = @{}
        for ($k = 0; $k -lt $rows.Count; $k++) {
            if ($k) { [void]$v.Append(',') }
            [void]$v.Append("(@g$k,@t$k,@y$k,@c$k,@e$k,@so$k,@dd$k,@id$k,@lt$k,@n$k,@tm$k,@mu$k,@up$k,@tp$k,@tx$k,@lk$k,@pr$k,@ps$k)")
            $r=$rows[$k]; $p["g$k"]=$r.g;$p["t$k"]=$r.t;$p["y$k"]=$r.y;$p["c$k"]=$r.c;$p["e$k"]=$r.e
            $p["so$k"]=$r.so;$p["dd$k"]=$r.dd;$p["id$k"]=$r.id;$p["lt$k"]=$r.lt;$p["n$k"]=$r.n
            $p["tm$k"]=$r.tm;$p["mu$k"]=$r.mu;$p["up$k"]=$r.up;$p["tp$k"]=$r.tp;$p["tx$k"]=$r.tx
            $p["lk$k"]=$r.lk;$p["pr$k"]=$r.pr;$p["ps$k"]=$r.ps
        }
        Invoke-SqlUpdate -ConnectionName 'tgt' -Query "INSERT INTO tbl_documentation_main ($cols) VALUES $($v.ToString())" -Parameters $p -WhatIf:$false | Out-Null
    }

    $h = Get-SourceReader -Query $query; $rdr = $h.Reader
    try {
        if (-not $script:DryRun) { Start-SqlTransaction -ConnectionName 'tgt' | Out-Null }
        $i = 0
        while ($rdr.Read()) {
            $i++; Update-Stats -Table $tableTgt -Read 1
            $row   = [ordered]@{}; for ($f=0;$f -lt $rdr.FieldCount;$f++){$row[$rdr.GetName($f)]=$rdr.GetValue($f)}
            $oldId = [int]$row.ID2
            $oldTen= if($row.IDP -is [System.DBNull]){0}else{[int]$row.IDP}
            $newTen= Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) { Add-Warning -Table $tableTgt -Reason "No tenant IDP=$oldTen" -SourceKey $oldId; Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1; continue }
            $rawTyp = if($row.ID1 -is [System.DBNull]){0}else{[int]$row.ID1}
            $typ    = Resolve-LookupId -Master 'tbl_documentation_type' -Value $rawTyp
            $oldCon = if($row.ID0 -is [System.DBNull]){0}else{[int]$row.ID0}
            $oldEnt = if($row.IDM -is [System.DBNull]){0}else{[int]$row.IDM}
            $batch.Add(@{
                g  = (New-Guid2)
                t  = $newTen
                y  = $typ
                c  = Get-MappedId -Map $script:Maps.Contact -Key $oldCon -Default 0
                e  = Get-MappedId -Map $script:Maps.Entity  -Key $oldEnt -Default 0
                so = if($row.Sorter -is [System.DBNull]){0}else{[int]$row.Sorter}
                dd = if($row.Datum  -is [System.DBNull]){[datetime]::Now}else{[datetime]$row.Datum}
                id = ConvertTo-DbString $row.GONr -MaxLen 10 -Required
                lt = if($longTextIdSet -contains $rawTyp){[string](ConvertTo-DbValue $row.Kommentar)}else{$null}
                n  = if($row.XCnt -is [System.DBNull]){1}else{ConvertTo-SmallInt $row.XCnt}
                tm = if($row.Druckdatum -is [System.DBNull]){0}else{try{ConvertTo-SmallInt (([datetime]$row.Druckdatum).Hour*60+([datetime]$row.Druckdatum).Minute)}catch{0}}
                mu = if($row.Multi    -is [System.DBNull]){0.0}else{[decimal]$row.Multi}
                up = if($row.Betrag   -is [System.DBNull]){0.0}else{[decimal]$row.Betrag}
                tp = if($row.GesBetrag-is [System.DBNull]){0.0}else{[decimal]$row.GesBetrag}
                tx = if($row.Steuer   -is [System.DBNull]){0.0}else{[decimal]$row.Steuer}
                lk = ConvertTo-BitInt $row.Lock
                pr = ConvertTo-BitInt $row.Gedruckt
                ps = ConvertTo-BitInt $row.Storniert
            })
            if ($batch.Count -ge $batchMax) {
                if (-not $script:DryRun) { Flush-DocBatch $batch }
                Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear()
                if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt'|Out-Null; Start-SqlTransaction -ConnectionName 'tgt'|Out-Null }
            }
            if (($i % 5000) -eq 0 -or $i -eq $total) { Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i,$total) -PercentComplete (([double]$i/[double]$total)*100) }
        }
        if ($batch.Count -gt 0) { if (-not $script:DryRun) { Flush-DocBatch $batch }; Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear() }
        if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null }
    } finally {
        Close-SourceReader $h
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    $script:Maps.Doc = @{}
    Write-OK ("{0} documentation rows inserted." -f ($script:Stats[$tableTgt].Inserted))
}

function Step7_Migrate-ReceiptContent {
    # Multi-row batch INSERTs (150 rows/query) — RcptCont map has no downstream use.
    [CmdletBinding()] param()
    $tableTgt = 'tbl_receipt_content'
    $batchMax = 150  # 150 rows × 19 params = 2850 — within driver limits
    Write-Section "Step 7/9: Tabelle_Abre (ID1>10)  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_Abre] WHERE ID1 > 10'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'No Tabelle_Abre rows with ID1>10.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)ID2, Kommentar2, IDP, IDR, IDM, ID1, ID4,
       Sorter, x AS XCnt, Datum, GONr, IDKurz,
       Multi, Betrag, GesBetrag, Steuer, Analog, Lock, Gedruckt, Storniert
FROM [Tabelle_Abre]
WHERE ID1 > 10
ORDER BY ID2
"@
    $cols  = 'fld_guiid,fld_tenant_id,fld_receipt_id,fld_type_id,fld_entity_id,fld_service_id,fld_sorter,fld_number,fld_docdate,fld_identifier,fld_service_text,fld_multiplier,fld_unit_price,fld_total_price,fld_tax_rate,fld_analog_flag,fld_lock,fld_print,fld_passive'
    $batch = [System.Collections.Generic.List[hashtable]]::new($batchMax)

    function Flush-RcptContBatch { param($rows)
        if ($rows.Count -eq 0) { return }
        $v = [System.Text.StringBuilder]::new(); $p = @{}
        for ($k = 0; $k -lt $rows.Count; $k++) {
            if ($k) { [void]$v.Append(',') }
            [void]$v.Append("(@g$k,@t$k,@r$k,@y$k,@e$k,@sv$k,@so$k,@n$k,@dd$k,@id$k,@st$k,@mu$k,@up$k,@tp$k,@tx$k,@an$k,@lk$k,@pr$k,@ps$k)")
            $r=$rows[$k]
            $p["g$k"]=$r.g;   $p["t$k"]=$r.t;   $p["r$k"]=$r.r;   $p["y$k"]=$r.y;   $p["e$k"]=$r.e
            $p["sv$k"]=$r.sv; $p["so$k"]=$r.so; $p["n$k"]=$r.n;   $p["dd$k"]=$r.dd; $p["id$k"]=$r.id
            $p["st$k"]=$r.st; $p["mu$k"]=$r.mu; $p["up$k"]=$r.up; $p["tp$k"]=$r.tp; $p["tx$k"]=$r.tx
            $p["an$k"]=$r.an; $p["lk$k"]=$r.lk; $p["pr$k"]=$r.pr; $p["ps$k"]=$r.ps
        }
        Invoke-SqlUpdate -ConnectionName 'tgt' -Query "INSERT INTO tbl_receipt_content ($cols) VALUES $($v.ToString())" -Parameters $p -WhatIf:$false | Out-Null
    }

    $h = Get-SourceReader -Query $query; $rdr = $h.Reader
    try {
        if (-not $script:DryRun) { Start-SqlTransaction -ConnectionName 'tgt' | Out-Null }
        $i = 0
        while ($rdr.Read()) {
            $i++; Update-Stats -Table $tableTgt -Read 1
            $row = [ordered]@{}; for ($f=0;$f -lt $rdr.FieldCount;$f++){$row[$rdr.GetName($f)]=$rdr.GetValue($f)}
            $oldId = [int]$row.ID2
            $oldTen = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newTen = Get-MappedId -Map $script:Maps.Tenant -Key $oldTen -Default $script:DefaultTenantId
            if ($newTen -eq 0) { Add-Warning -Table $tableTgt -Reason "No tenant mapping for IDP=$oldTen" -SourceKey $oldId; Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1; continue }
            $oldRcp = if ($row.IDR -is [System.DBNull]) { 0 } else { [int]$row.IDR }
            $newRcp = Get-MappedId -Map $script:Maps.Receipt -Key $oldRcp -Default 0
            if ($newRcp -eq 0) { Add-Warning -Table $tableTgt -Reason "No receipt mapping for IDR=$oldRcp - skipping" -SourceKey $oldId; Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1; continue }
            $oldEnt = if ($row.IDM -is [System.DBNull]) { 0 } else { [int]$row.IDM }
            $newEnt = Get-MappedId -Map $script:Maps.Entity -Key $oldEnt -Default 0
            # Lookup FKs validated against simplimed master tables.
            $serId  = Resolve-LookupId -Master 'tbl_services_main'      -Value $row.ID4
            $typId  = Resolve-LookupId -Master 'tbl_documentation_type' -Value $row.ID1

            $batch.Add(@{
                g  = (New-Guid2)
                t  = $newTen
                r  = $newRcp
                y  = $typId
                e  = $newEnt
                sv = $serId
                so = ConvertTo-SmallInt $row.Sorter
                n  = if ($row.XCnt -is [System.DBNull]) { 1 } else { ConvertTo-SmallInt $row.XCnt }
                dd = if ($row.Datum  -is [System.DBNull]) { Get-Date } else { [datetime]$row.Datum }
                id = ConvertTo-DbString $row.GONr   -MaxLen 10  -Required
                st = ConvertTo-DbString $row.IDKurz -MaxLen 250 -Required
                mu = if ($row.Multi     -is [System.DBNull]) { 0.0 } else { [decimal]$row.Multi }
                up = if ($row.Betrag    -is [System.DBNull]) { 0.0 } else { [decimal]$row.Betrag }
                tp = if ($row.GesBetrag -is [System.DBNull]) { 0.0 } else { [decimal]$row.GesBetrag }
                tx = if ($row.Steuer    -is [System.DBNull]) { 0.0 } else { [decimal]$row.Steuer }
                an = ConvertTo-BitInt $row.Analog
                lk = ConvertTo-BitInt $row.Lock
                pr = ConvertTo-BitInt $row.Gedruckt
                ps = ConvertTo-BitInt $row.Storniert
            })

            if ($batch.Count -ge $batchMax) {
                if (-not $script:DryRun) { Flush-RcptContBatch $batch }
                Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear()
                if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt'|Out-Null; Start-SqlTransaction -ConnectionName 'tgt'|Out-Null }
            }
            if (($i % 5000) -eq 0 -or $i -eq $total) { Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i,$total) -PercentComplete (([double]$i/[double]$total)*100) }
        }
        if ($batch.Count -gt 0) { if (-not $script:DryRun) { Flush-RcptContBatch $batch }; Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear() }
        if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null }
    } finally {
        Close-SourceReader $h
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} receipt-content rows inserted." -f ($script:Stats[$tableTgt].Inserted))
    $script:Maps.RcptCont = @{}
    $script:Maps.Receipt  = @{}   # Receipt map not needed after Step 7
}

function Step8_Migrate-Bookkeeping {
    [CmdletBinding()] param()
    $tableTgt = 'tbl_bookkeeping'
    Write-Section "Step 8/9: Tabelle_JourBuch  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_JourBuch]'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'Tabelle_JourBuch is empty.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    # Note: source table does NOT have an IDZ column. The mapping spec says
    # `IDZ -> fld_tenant_id` but no such column exists on Tabelle_JourBuch. We
    # therefore derive the tenant id from the linked patient's tenant chain
    # (IDP -> patient -> patient.IDP -> tenant) by joining Tabelle_Patienten.
    $query = @"
SELECT $($top)b.ID0, b.GuiID, b.IDK AS IDk, b.IDG, b.IDB, b.IDP, b.IDR, b.IDM, b.[Währung] AS Waehrung,
       b.RechNr, b.IDKurz, b.IDGegen, b.Buchtext, b.Kommentar, b.Datei,
       b.Ausgabe, b.Einnahme, b.Steuer, b.Privat AS Abziehbar,
       b.TSEStrSig, b.TSEString, b.TSELog AS TSESign, b.TSEZahl,
       b.Datum, b.Berichtdatum, b.Ermittlung, b.Anfang, b.Selekt, b.Drucken,
       b.Serie, b.Doppelt, b.Zuordnung, b.Lock, b.Storniert,
       p.IDP AS PatientTenant
FROM [Tabelle_JourBuch] b
LEFT JOIN [Tabelle_Patienten] p ON p.ID0 = b.IDP
ORDER BY b.ID0
"@
    # Multi-row batch INSERTs (100 rows/query) — Booking map has no downstream use.
    $batchMax = 100  # 100 rows × 35 params = 3500 — within driver limits
    $cols  = 'fld_guiid,fld_tenant_id,fld_account_key,fld_counter_account_key,fld_bank_id,fld_contact_id,fld_receipt_id,fld_entity_id,fld_currency_id,fld_invoice_no,fld_short_name,fld_counter_short,fld_booking_text,fld_comment,fld_filename,fld_receipt_no,fld_expense,fld_income,fld_tax_rate,fld_private_share,fld_tse_sig_string,fld_tse_string,fld_tse_log,fld_tse_count,fld_booking_date,fld_report_date,fld_determined,fld_opening,fld_selected,fld_print,fld_series,fld_duplicate,fld_assigned,fld_locked,fld_passive'
    $batch = [System.Collections.Generic.List[hashtable]]::new($batchMax)

    function Flush-BookBatch { param($rows)
        if ($rows.Count -eq 0) { return }
        $v = [System.Text.StringBuilder]::new(); $p = @{}
        for ($k = 0; $k -lt $rows.Count; $k++) {
            if ($k) { [void]$v.Append(',') }
            [void]$v.Append("(@g$k,@tn$k,@ak$k,@ck$k,@bk$k,@cn$k,@rc$k,@en$k,@cu$k,@iv$k,@sn$k,@cs$k,@bt$k,@cm$k,@fn$k,@rn$k,@ex$k,@ic$k,@tx$k,@pv$k,@ts$k,@tr$k,@tl$k,@tc$k,@bd$k,@rd$k,@dt$k,@op$k,@sl$k,@pr$k,@sr$k,@dp$k,@ag$k,@lk$k,@ps$k)")
            $r=$rows[$k]
            $p["g$k"]=$r.g;   $p["tn$k"]=$r.t;   $p["ak$k"]=$r.ak; $p["ck$k"]=$r.cak; $p["bk$k"]=$r.bk
            $p["cn$k"]=$r.cn; $p["rc$k"]=$r.rc;  $p["en$k"]=$r.en; $p["cu$k"]=$r.cu;  $p["iv$k"]=$r.iv
            $p["sn$k"]=$r.sn; $p["cs$k"]=$r.cs;  $p["bt$k"]=$r.bt; $p["cm$k"]=$r.cm;  $p["fn$k"]=$r.fn
            $p["rn$k"]=$r.rn; $p["ex$k"]=$r.ex;  $p["ic$k"]=$r.inc;$p["tx$k"]=$r.tx;  $p["pv$k"]=$r.pv
            $p["ts$k"]=$r.ts; $p["tr$k"]=$r.tr;  $p["tl$k"]=$r.tl; $p["tc$k"]=$r.tc;  $p["bd$k"]=$r.bd
            $p["rd$k"]=$r.rd; $p["dt$k"]=$r.dt;  $p["op$k"]=$r.op; $p["sl$k"]=$r.sl;  $p["pr$k"]=$r.pr
            $p["sr$k"]=$r.sr; $p["dp$k"]=$r.dp;  $p["ag$k"]=$r.asg;$p["lk$k"]=$r.lk;  $p["ps$k"]=$r.ps
        }
        Invoke-SqlUpdate -ConnectionName 'tgt' -Query "INSERT INTO tbl_bookkeeping ($cols) VALUES $($v.ToString())" -Parameters $p -WhatIf:$false | Out-Null
    }

    $h = Get-SourceReader -Query $query; $rdr = $h.Reader
    try {
        if (-not $script:DryRun) { Start-SqlTransaction -ConnectionName 'tgt' | Out-Null }
        $i = 0
        while ($rdr.Read()) {
            $i++; Update-Stats -Table $tableTgt -Read 1
            $row = [ordered]@{}; for ($f=0;$f -lt $rdr.FieldCount;$f++){$row[$rdr.GetName($f)]=$rdr.GetValue($f)}
            $oldId = [int]$row.ID0

            $tenantSourceKey = if ($row.PatientTenant -is [System.DBNull]) { 0 } else { [int]$row.PatientTenant }
            $newTen = Get-MappedId -Map $script:Maps.Tenant -Key $tenantSourceKey -Default $script:DefaultTenantId
            if ($newTen -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No tenant resolvable from contact chain (IDP=$($row.IDP))" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            $oldCon = if ($row.IDP -is [System.DBNull]) { 0 } else { [int]$row.IDP }
            $newCon = Get-MappedId -Map $script:Maps.Contact -Key $oldCon -Default 0
            $oldEnt = if ($row.IDM -is [System.DBNull]) { 0 } else { [int]$row.IDM }
            $newEnt = Get-MappedId -Map $script:Maps.Entity  -Key $oldEnt -Default 0
            $oldRcp = if ($row.IDR -is [System.DBNull]) { 0 } else { [int]$row.IDR }
            $newRcp = Get-MappedId -Map $script:Maps.Receipt -Key $oldRcp -Default 0
            $bankId = Resolve-LookupId -Master 'tbl_banking' -Value $row.IDB
            $bdate  = if ($row.Datum -is [System.DBNull]) { (Get-Date).Date } else { ([datetime]$row.Datum).Date }
            $rdate  = if ($row.Berichtdatum -is [System.DBNull]) { $null } else { ([datetime]$row.Berichtdatum).Date }

            $batch.Add(@{
                g   = (New-Guid2)
                t   = $newTen
                ak  = ConvertTo-DbString $row.IDk       -MaxLen 10 -Required
                cak = ConvertTo-DbString $row.IDG       -MaxLen 10 -Required
                bk  = $bankId
                cn  = $newCon
                rc  = $newRcp
                en  = $newEnt
                cu  = Resolve-LookupId -Master 'tbl_currencies' -Value $row.Waehrung
                iv  = ConvertTo-DbString $row.RechNr    -MaxLen 50  -Required
                sn  = ConvertTo-DbString $row.IDKurz    -MaxLen 100 -Required
                cs  = ConvertTo-DbString $row.IDGegen   -MaxLen 100
                bt  = ConvertTo-DbString $row.Buchtext  -MaxLen 200 -Required
                cm  = ConvertTo-DbString $row.Kommentar -MaxLen 250
                fn  = ConvertTo-DbString $row.Datei     -MaxLen 100
                rn  = $oldRcp
                ex  = if ($row.Ausgabe  -is [System.DBNull]) { 0.0 } else { [decimal]$row.Ausgabe }
                inc = if ($row.Einnahme -is [System.DBNull]) { 0.0 } else { [decimal]$row.Einnahme }
                tx  = if ($row.Steuer   -is [System.DBNull]) { 0.0 } else { [decimal]$row.Steuer }
                pv  = if ($row.Abziehbar-is [System.DBNull]) { 0.0 } else { [decimal]$row.Abziehbar }
                ts  = ConvertTo-DbString $row.TSEStrSig -MaxLen 200
                tr  = ConvertTo-DbString $row.TSEString -MaxLen 255
                tl  = ConvertTo-DbString $row.TSESign   -MaxLen 100
                tc  = if ($row.TSEZahl -is [System.DBNull]) { $null } else { [int]$row.TSEZahl }
                bd  = $bdate
                rd  = $rdate
                dt  = ConvertTo-BitInt $row.Ermittlung
                op  = ConvertTo-BitInt $row.Anfang
                sl  = ConvertTo-BitInt $row.Selekt
                pr  = ConvertTo-BitInt $row.Drucken
                sr  = ConvertTo-BitInt $row.Serie
                dp  = ConvertTo-BitInt $row.Doppelt
                asg = ConvertTo-BitInt $row.Zuordnung
                lk  = ConvertTo-BitInt $row.Lock
                ps  = ConvertTo-BitInt $row.Storniert
            })

            if ($batch.Count -ge $batchMax) {
                if (-not $script:DryRun) { Flush-BookBatch $batch }
                Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear()
                if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt'|Out-Null; Start-SqlTransaction -ConnectionName 'tgt'|Out-Null }
            }
            if (($i % 2000) -eq 0 -or $i -eq $total) { Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i,$total) -PercentComplete (([double]$i/[double]$total)*100) }
        }
        if ($batch.Count -gt 0) { if (-not $script:DryRun) { Flush-BookBatch $batch }; Update-Stats -Table $tableTgt -Inserted $batch.Count; $batch.Clear() }
        if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null }
    } finally {
        Close-SourceReader $h
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    Write-OK ("{0} bookkeeping rows inserted." -f ($script:Stats[$tableTgt].Inserted))
    $script:Maps.Booking = @{}
    $script:Maps.Contact = @{}   # Contact map not needed after Step 8
    # NOTE: Tenant + Entity maps are still needed by Step 9 — cleared there instead
    [System.GC]::Collect()
}

function Step9_Migrate-Protocoll {
    # Uses multi-row batch INSERTs (200 rows/query) instead of one INSERT per row.
    # The Protocoll map has no downstream consumers so RETURNING is unnecessary.
    # This reduces 142k round-trips to ~710, making Step 9 as fast as Steps 5-8.
    [CmdletBinding()] param()
    $tableTgt = 'tbl_appoint_protocoll'
    $batchMax = 200   # 200 rows × 7 params = 1400 params — well within driver limits
    Write-Section "Step 9/9: Tabelle_TermPro  ->  $tableTgt"
    Initialize-Stats -Table $tableTgt

    $countQuery = 'SELECT COUNT(*) FROM [Tabelle_TermPro]'
    $total      = [int](Invoke-SqlScalar -ConnectionName 'src' -Query $countQuery)
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $total = $MaxRows }
    if ($total -eq 0) { Write-Warn2 'Tabelle_TermPro is empty.'; return }

    $top   = if ($MaxRows -gt 0) { "TOP ($MaxRows) " } else { '' }
    $query = @"
SELECT $($top)t.IDA, t.ID2, t.Datum, t.IDKurz, t.Selekt,
       w.IDP AS ApptTenant, w.IDM AS ApptEntity
FROM [Tabelle_TermPro] t
LEFT JOIN [Tabelle_PatientenWv] w ON w.ID2 = t.ID2
ORDER BY t.IDA
"@
    $cols  = 'fld_guiid, fld_tenant_id, fld_appointment_id, fld_entity_id, fld_date_log, fld_logtext, fld_select'
    $batch = [System.Collections.Generic.List[hashtable]]::new($batchMax)

    function Flush-ProtocollBatch {
        param([System.Collections.Generic.List[hashtable]]$rows)
        if ($rows.Count -eq 0) { return }
        $vals   = [System.Text.StringBuilder]::new()
        $params = @{}
        for ($k = 0; $k -lt $rows.Count; $k++) {
            if ($k -gt 0) { [void]$vals.Append(',') }
            [void]$vals.Append("(@g$k,@t$k,@a$k,@e$k,@d$k,@l$k,@s$k)")
            $r = $rows[$k]
            $params["g$k"] = $r.g; $params["t$k"] = $r.t; $params["a$k"] = $r.a
            $params["e$k"] = $r.e; $params["d$k"] = $r.d; $params["l$k"] = $r.l
            $params["s$k"] = $r.s
        }
        $sql = "INSERT INTO tbl_appoint_protocoll ($cols) VALUES $($vals.ToString())"
        Invoke-SqlUpdate -ConnectionName 'tgt' -Query $sql -Parameters $params -WhatIf:$false | Out-Null
    }

    $h = Get-SourceReader -Query $query
    $rdr = $h.Reader
    try {
        if (-not $script:DryRun) { Start-SqlTransaction -ConnectionName 'tgt' | Out-Null }
        $i = 0
        while ($rdr.Read()) {
            $i++
            Update-Stats -Table $tableTgt -Read 1
            $row    = [ordered]@{}
            for ($f = 0; $f -lt $rdr.FieldCount; $f++) { $row[$rdr.GetName($f)] = $rdr.GetValue($f) }
            $oldId  = [int]$row.IDA
            $oldApt = if ($row.ID2 -is [System.DBNull]) { 0 } else { [int]$row.ID2 }
            $newApt = Get-MappedId -Map $script:Maps.Appoint -Key $oldApt -Default 0
            if ($newApt -eq 0) {
                Add-Warning -Table $tableTgt -Reason "No appointment mapping for ID2=$oldApt - skipping" -SourceKey $oldId
                Update-Stats -Table $tableTgt -Warnings 1 -Skipped 1
                continue
            }
            $tenSrc = if ($row.ApptTenant -is [System.DBNull]) { 0 } else { [int]$row.ApptTenant }
            $entSrc = if ($row.ApptEntity -is [System.DBNull]) { 0 } else { [int]$row.ApptEntity }
            $batch.Add(@{
                g = (New-Guid2)
                t = Get-MappedId -Map $script:Maps.Tenant -Key $tenSrc -Default $script:DefaultTenantId
                a = $newApt
                e = Get-MappedId -Map $script:Maps.Entity -Key $entSrc -Default 0
                d = if ($row.Datum -is [System.DBNull]) { [datetime]::Now } else { [datetime]$row.Datum }
                l = ConvertTo-DbString $row.IDKurz -MaxLen 250
                s = ConvertTo-BitInt $row.Selekt
            })

            if ($batch.Count -ge $batchMax) {
                if (-not $script:DryRun) { Flush-ProtocollBatch $batch }
                Update-Stats -Table $tableTgt -Inserted $batch.Count
                $batch.Clear()
                if (-not $script:DryRun) {
                    Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null
                    Start-SqlTransaction    -ConnectionName 'tgt' | Out-Null
                }
            }
            if (($i % 5000) -eq 0 -or $i -eq $total) {
                Write-Progress -Id 1 -Activity "Migrating $tableTgt" -Status ("{0:N0} / {1:N0}" -f $i, $total) -PercentComplete (([double]$i / [double]$total) * 100)
            }
        }
        if ($batch.Count -gt 0) {
            if (-not $script:DryRun) { Flush-ProtocollBatch $batch }
            Update-Stats -Table $tableTgt -Inserted $batch.Count
            $batch.Clear()
        }
        if (-not $script:DryRun) { Complete-SqlTransaction -ConnectionName 'tgt' | Out-Null }
    } finally {
        Close-SourceReader $h
        Write-Progress -Id 1 -Completed -Activity "Migrating $tableTgt"
    }
    # Free maps no longer needed
    $script:Maps.Appoint = @{}
    $script:Maps.Tenant  = @{}
    $script:Maps.Entity  = @{}
    [System.GC]::Collect()
    Write-OK ("{0} protocoll rows inserted." -f ($script:Stats[$tableTgt].Inserted))
}

#endregion

#region Verification ----------------------------------------------------------

# Hard-coded relation catalog. The migration uses logical references only — MariaDB
# has no FK constraints on these tables — so we self-validate at the end of every
# run. Any orphan in a STRONG relation aborts the migration; orphans in LOOKUP
# relations (1:1 source IDs not part of the script's ID maps) are reported as
# warnings so the operator can decide whether to remap or zero them.

$script:StrongRelations = @(
    @{ Child='tbl_entity_main';        ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_contact_main';       ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_receipt_main';       ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_receipt_main';       ChildCol='fld_contact_id';     Parent='tbl_contact_main'    },
    @{ Child='tbl_appoint_main';       ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_appoint_main';       ChildCol='fld_entity_id';      Parent='tbl_entity_main'     },
    @{ Child='tbl_appoint_main';       ChildCol='fld_contact_id';     Parent='tbl_contact_main'    },
    @{ Child='tbl_documentation_main'; ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_documentation_main'; ChildCol='fld_contact_id';     Parent='tbl_contact_main'    },
    @{ Child='tbl_documentation_main'; ChildCol='fld_entity_id';      Parent='tbl_entity_main'     },
    @{ Child='tbl_receipt_content';    ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_receipt_content';    ChildCol='fld_receipt_id';     Parent='tbl_receipt_main'    },
    @{ Child='tbl_receipt_content';    ChildCol='fld_entity_id';      Parent='tbl_entity_main'     },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_contact_id';     Parent='tbl_contact_main'    },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_receipt_id';     Parent='tbl_receipt_main'    },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_entity_id';      Parent='tbl_entity_main'     },
    @{ Child='tbl_appoint_protocoll';  ChildCol='fld_tenant_id';      Parent='tbl_tenant_entities' },
    @{ Child='tbl_appoint_protocoll';  ChildCol='fld_appointment_id'; Parent='tbl_appoint_main'    },
    @{ Child='tbl_appoint_protocoll';  ChildCol='fld_entity_id';      Parent='tbl_entity_main'     }
)

$script:LookupRelations = @(
    @{ Child='tbl_contact_main';       ChildCol='fld_currencies_id';  Parent='tbl_currencies'        },
    @{ Child='tbl_contact_main';       ChildCol='fld_payment_id';     Parent='tbl_contact_payments'  },
    @{ Child='tbl_contact_main';       ChildCol='fld_dispatch_id';    Parent='tbl_receipt_dispatch'  },
    @{ Child='tbl_contact_main';       ChildCol='fld_service_id';     Parent='tbl_services_main'     },
    @{ Child='tbl_contact_main';       ChildCol='fld_marital_id';     Parent='tbl_contact_marital'   },
    @{ Child='tbl_contact_main';       ChildCol='fld_gender_id';      Parent='tbl_contact_gender'    },
    @{ Child='tbl_contact_main';       ChildCol='fld_insurance_id';   Parent='tbl_contact_insurance' },
    @{ Child='tbl_receipt_main';       ChildCol='fld_currencies_id';  Parent='tbl_currencies'        },
    @{ Child='tbl_receipt_main';       ChildCol='fld_payment_id';     Parent='tbl_contact_payments'  },
    @{ Child='tbl_receipt_main';       ChildCol='fld_dispatch_id';    Parent='tbl_receipt_dispatch'  },
    @{ Child='tbl_receipt_main';       ChildCol='fld_type_id';        Parent='tbl_receipt_type'      },
    @{ Child='tbl_appoint_main';       ChildCol='fld_location_id';    Parent='tbl_appoint_location'  },
    @{ Child='tbl_appoint_main';       ChildCol='fld_status_id';      Parent='tbl_appoint_status'    },
    @{ Child='tbl_appoint_main';       ChildCol='fld_priority_id';    Parent='tbl_appoint_priority'  },
    @{ Child='tbl_documentation_main'; ChildCol='fld_typ_id';         Parent='tbl_documentation_type'},
    @{ Child='tbl_receipt_content';    ChildCol='fld_type_id';        Parent='tbl_documentation_type'},
    @{ Child='tbl_receipt_content';    ChildCol='fld_service_id';     Parent='tbl_services_main'     },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_currency_id';    Parent='tbl_currencies'        },
    @{ Child='tbl_bookkeeping';        ChildCol='fld_bank_id';        Parent='tbl_banking'           }
)

function Test-TargetTableExists {
    param([string]$Table)
    $n = [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query @"
SELECT COUNT(*) FROM information_schema.tables
WHERE table_schema = DATABASE() AND table_name = @t
"@ -Parameters @{ t = $Table })
    return $n -gt 0
}

function Test-TargetColumnExists {
    param([string]$Table, [string]$Column)
    $n = [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query @"
SELECT COUNT(*) FROM information_schema.columns
WHERE table_schema = DATABASE() AND table_name = @t AND column_name = @c
"@ -Parameters @{ t = $Table; c = $Column })
    return $n -gt 0
}

function Get-RelationOrphans {
    param([string]$Child, [string]$ChildCol, [string]$Parent)
    return [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query @"
SELECT COUNT(*) FROM $Child ch
WHERE ch.$ChildCol IS NOT NULL AND ch.$ChildCol <> 0
  AND NOT EXISTS (SELECT 1 FROM $Parent pr WHERE pr.fld_id = ch.$ChildCol)
"@)
}

function Invoke-CleanupTarget {
    # TRUNCATE all migrated tables in the current target database. Only allowed
    # against databases whose name contains "staging" — production data is
    # protected by name. -Force overrides this guard for explicit recovery.
    [CmdletBinding()] param()
    $db = Get-TargetDatabaseName
    Write-Banner ("Cleanup target database: {0}" -f $db) 'Yellow'

    if ($db -notmatch '(?i)staging' -and -not $Force) {
        Write-Err ("Refusing to TRUNCATE database '{0}' — name does not contain 'staging'." -f $db)
        Write-Host '   Use -Force to override (production data WILL be lost).' -ForegroundColor Red
        return
    }
    if ($Force -and $db -notmatch '(?i)staging') {
        Write-Warn2 ("-Force used against non-staging database '{0}'. Last chance:" -f $db)
        $confirm = Read-Host ("Type the literal database name '{0}' to confirm" -f $db)
        if ($confirm -ne $db) {
            Write-Err 'Confirmation mismatch — aborting cleanup.'
            return
        }
    }

    $tables = @(
        'tbl_appoint_protocoll',
        'tbl_receipt_content',
        'tbl_bookkeeping',
        'tbl_documentation_main',
        'tbl_appoint_main',
        'tbl_receipt_main',
        'tbl_contact_main',
        'tbl_entity_main',
        'tbl_tenant_entities'
    )
    Write-Section 'Truncating migrated tables (children first)'
    Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET FOREIGN_KEY_CHECKS = 0' | Out-Null
    foreach ($t in $tables) {
        Invoke-SqlUpdate -ConnectionName 'tgt' -Query "TRUNCATE TABLE $t" | Out-Null
        Write-Host ("   Truncated {0}" -f $t) -ForegroundColor DarkGray
    }
    Invoke-SqlUpdate -ConnectionName 'tgt' -Query 'SET FOREIGN_KEY_CHECKS = 1' | Out-Null
    Write-OK 'All migrated tables truncated.'

    Write-Section 'Row counts after cleanup'
    foreach ($t in $tables) {
        $n = [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query "SELECT COUNT(*) FROM $t")
        Write-Host ("   {0,-30} {1,8:N0}" -f $t, $n) -ForegroundColor Gray
    }
}

function Repair-LookupOrphans {
    # In-place UPDATE that zeroes any lookup FK whose value does not exist
    # in the corresponding simplimed master table. Safe for re-runs and for
    # databases that already contain mixed (legacy + freshly imported) data:
    # only orphan rows are touched, valid references remain intact.
    [CmdletBinding()] param()
    Write-Section 'Step 10a: Repair existing lookup orphans (auto-zero, in-place)'
    Initialize-Stats -Table 'repair_lookup_orphans'

    $totalFixed = 0
    foreach ($r in $script:LookupRelations) {
        if (-not (Test-TargetColumnExists -Table $r.Child -Column $r.ChildCol)) { continue }
        if (-not (Test-TargetTableExists  -Table $r.Parent))                    { continue }

        # Cheap orphan probe first; if 0, skip the audit query and the UPDATE.
        $orphanCount = [int](Invoke-SqlScalar -ConnectionName 'tgt' -Query @"
SELECT COUNT(*) FROM $($r.Child) ch
LEFT JOIN $($r.Parent) pr ON pr.fld_id = ch.$($r.ChildCol)
WHERE ch.$($r.ChildCol) <> 0 AND pr.fld_id IS NULL
"@)
        if ($orphanCount -eq 0) { continue }

        $sample = Invoke-SqlQuery -ConnectionName 'tgt' -WarningAction SilentlyContinue -Query @"
SELECT DISTINCT ch.$($r.ChildCol) AS OrphanId
FROM $($r.Child) ch
LEFT JOIN $($r.Parent) pr ON pr.fld_id = ch.$($r.ChildCol)
WHERE ch.$($r.ChildCol) <> 0 AND pr.fld_id IS NULL
ORDER BY ch.$($r.ChildCol)
LIMIT 50
"@
        $sampleIds = ($sample | ForEach-Object { $_.OrphanId }) -join ','

        $affected = Invoke-SqlUpdate -ConnectionName 'tgt' -Query @"
UPDATE $($r.Child) ch
LEFT JOIN $($r.Parent) pr ON pr.fld_id = ch.$($r.ChildCol)
SET ch.$($r.ChildCol) = 0
WHERE ch.$($r.ChildCol) <> 0 AND pr.fld_id IS NULL
"@
        $totalFixed += $affected
        Write-Host (' [FIX] {0,-22} {1,-22} -> {2,-22} {3,8:N0} rows zeroed (orphan ids: {4})' -f $r.Child, $r.ChildCol, $r.Parent, $affected, $sampleIds) -ForegroundColor Yellow
        Update-Stats -Table 'repair_lookup_orphans' -Inserted $affected
        Add-Warning -Table $r.Child -Reason ("Auto-zeroed {0} orphan rows in {1} (orphan ids: {2})" -f $affected, $r.ChildCol, $sampleIds) -SourceKey '<repair>'
    }

    if ($totalFixed -eq 0) {
        Write-OK 'No lookup orphans found — nothing to repair.'
    } else {
        Write-OK ('Repaired {0:N0} lookup orphan row(s) in total.' -f $totalFixed)
    }

    # Migration-time skipped values that the Resolve-LookupId helper recorded
    $hasDropped = $false
    foreach ($k in $script:LookupDropped.Keys) { if ($script:LookupDropped[$k].Count -gt 0) { $hasDropped = $true; break } }
    if ($hasDropped) {
        Write-Host ''
        Write-Host '  Lookup ids dropped during this run (set to 0 because master row absent):' -ForegroundColor DarkYellow
        foreach ($k in $script:LookupDropped.Keys) {
            $set = $script:LookupDropped[$k]
            if ($set.Count -gt 0) {
                $list = ($set | Sort-Object | Select-Object -First 25) -join ','
                $more = if ($set.Count -gt 25) { (' (+{0} more)' -f ($set.Count - 25)) } else { '' }
                Write-Host ('    {0,-28} {1,5} distinct id(s): {2}{3}' -f $k, $set.Count, $list, $more) -ForegroundColor DarkGray
            }
        }
    }
}

function Step10_Verify-Relations {
    [CmdletBinding()] param()
    Write-Section 'Step 10/10: Verify relational integrity (no native FKs)'
    Initialize-Stats -Table 'verify_relations'

    if ($script:DryRun) {
        Write-Warn2 'DryRun: skipping relation verification.'
        return
    }

    # First pass: in-place UPDATE that zeroes lookup orphans against the
    # already-loaded master tables. This makes the verification deterministic
    # — any remaining orphan after this pass is a true integrity violation.
    Repair-LookupOrphans

    $rep         = New-Object System.Collections.Generic.List[pscustomobject]
    $strongFails = 0

    Write-Progress -Id 1 -Activity 'Verifying relations' -Status 'Strong relations' -PercentComplete 0
    $i = 0; $n = $script:StrongRelations.Count
    foreach ($r in $script:StrongRelations) {
        $i++
        Write-Progress -Id 1 -Activity 'Verifying relations' -Status ("Strong  {0}.{1} -> {2}" -f $r.Child, $r.ChildCol, $r.Parent) -PercentComplete (([double]$i / [double]($n + $script:LookupRelations.Count)) * 100)
        if (-not (Test-TargetColumnExists -Table $r.Child -Column $r.ChildCol)) {
            $rep.Add([pscustomobject]@{ Kind='STRONG'; Child=$r.Child; Column=$r.ChildCol; Parent=$r.Parent; Orphans='-'; Status='COL_MISSING' }); continue
        }
        $orphans = Get-RelationOrphans -Child $r.Child -ChildCol $r.ChildCol -Parent $r.Parent
        $status  = if ($orphans -eq 0) { 'OK' } else { 'FAIL' }
        $rep.Add([pscustomobject]@{ Kind='STRONG'; Child=$r.Child; Column=$r.ChildCol; Parent=$r.Parent; Orphans=$orphans; Status=$status })
        if ($orphans -gt 0) {
            $strongFails++
            Add-MigError -Table $r.Child -Reason "FK integrity violated: $orphans rows in $($r.ChildCol) point to missing $($r.Parent).fld_id" -SourceKey '<rel>' -Exception (New-Object System.Exception 'Strong relation orphans')
            Update-Stats -Table 'verify_relations' -Errors 1
        }
    }

    foreach ($r in $script:LookupRelations) {
        $i++
        Write-Progress -Id 1 -Activity 'Verifying relations' -Status ("Lookup  {0}.{1} -> {2}" -f $r.Child, $r.ChildCol, $r.Parent) -PercentComplete (([double]$i / [double]($n + $script:LookupRelations.Count)) * 100)
        if (-not (Test-TargetColumnExists -Table $r.Child -Column $r.ChildCol)) {
            $rep.Add([pscustomobject]@{ Kind='LOOKUP'; Child=$r.Child; Column=$r.ChildCol; Parent=$r.Parent; Orphans='-'; Status='COL_MISSING' }); continue
        }
        if (-not (Test-TargetTableExists -Table $r.Parent)) {
            $rep.Add([pscustomobject]@{ Kind='LOOKUP'; Child=$r.Child; Column=$r.ChildCol; Parent=$r.Parent; Orphans='-'; Status='TBL_MISSING' }); continue
        }
        $orphans = Get-RelationOrphans -Child $r.Child -ChildCol $r.ChildCol -Parent $r.Parent
        $status  = if ($orphans -eq 0) { 'OK' } else { 'WARN' }
        $rep.Add([pscustomobject]@{ Kind='LOOKUP'; Child=$r.Child; Column=$r.ChildCol; Parent=$r.Parent; Orphans=$orphans; Status=$status })
        if ($orphans -gt 0) {
            Add-Warning -Table $r.Child -Reason "Lookup orphan: $orphans rows in $($r.ChildCol) -> $($r.Parent)" -SourceKey '<rel>'
            Update-Stats -Table 'verify_relations' -Warnings 1
        }
    }

    Write-Progress -Id 1 -Completed -Activity 'Verifying relations'

    Write-Host ''
    Write-Host '  Strong (script-managed) relations:' -ForegroundColor Cyan
    $rep | Where-Object Kind -eq 'STRONG' | Format-Table Child, Column, Parent, Orphans, Status -AutoSize | Out-String | Write-Host
    Write-Host '  Lookup (1:1 copied) relations:' -ForegroundColor DarkCyan
    $rep | Where-Object Kind -eq 'LOOKUP' | Format-Table Child, Column, Parent, Orphans, Status -AutoSize | Out-String | Write-Host

    if ($strongFails -eq 0) {
        Write-OK 'All strong relations validated (0 orphans).'
    } else {
        Write-Err ("{0} strong relation(s) FAILED — migration produced FK integrity violations." -f $strongFails)
        throw "Relation integrity check failed: $strongFails strong relation(s) have orphan rows."
    }
}

#endregion

#region Summary ----------------------------------------------------------------

function Write-FinalSummary {
    $end      = Get-Date
    $duration = $end - $script:StartTime
    Write-Banner 'Migration Summary' 'Green'
    Write-Host ("  Started   : {0}" -f $script:StartTime) -ForegroundColor Gray
    Write-Host ("  Finished  : {0}" -f $end)              -ForegroundColor Gray
    Write-Host ("  Duration  : {0:hh\:mm\:ss}" -f $duration) -ForegroundColor Gray
    Write-Host ("  Backup    : {0}" -f $script:BackupPath) -ForegroundColor Gray
    Write-Host ("  Transcript: {0}" -f $script:TranscriptPath) -ForegroundColor Gray
    Write-Host ('  DryRun    : {0}' -f $script:DryRun)     -ForegroundColor Gray
    Write-Host ''
    $rows = $script:Stats.GetEnumerator() | ForEach-Object { $_.Value }
    if ($rows.Count -gt 0) {
        $rows | Format-Table Table, Read, Inserted, Skipped, Warnings, Errors -AutoSize | Out-String | Write-Host
    }
    Write-Host ('  Total Read     : {0:N0}' -f (($rows | Measure-Object Read     -Sum).Sum)) -ForegroundColor Cyan
    Write-Host ('  Total Inserted : {0:N0}' -f (($rows | Measure-Object Inserted -Sum).Sum)) -ForegroundColor Green
    Write-Host ('  Total Skipped  : {0:N0}' -f (($rows | Measure-Object Skipped  -Sum).Sum)) -ForegroundColor Yellow
    Write-Host ('  Total Warnings : {0:N0}' -f $script:WarnCollection.Count) -ForegroundColor Yellow
    Write-Host ('  Total Errors   : {0:N0}' -f $script:ErrorCollection.Count) -ForegroundColor $(if ($script:ErrorCollection.Count -gt 0) { 'Red' } else { 'Green' })
    Write-Host ''
    if ($script:ErrorCollection.Count -gt 0) {
        Write-Host '  First 10 errors:' -ForegroundColor Red
        $script:ErrorCollection | Select-Object -First 10 | Format-Table Table, Key, Reason, Message -Wrap | Out-String | Write-Host
    }
    Write-Host '  ID maps (counts):' -ForegroundColor DarkCyan
    foreach ($k in $script:Maps.Keys) {
        Write-Host ("    {0,-12} {1,8:N0}" -f $k, $script:Maps[$k].Count)
    }
}

#endregion

#region Main -------------------------------------------------------------------

try {
    # Resolve effective mode: parameter -> menu (interactive) -> exit
    if ([string]::IsNullOrEmpty($Mode)) {
        $sel             = Show-MainMenu
        $effectiveMode   = $sel.Mode
        $script:DryRun   = [bool]$sel.DryRun
    } else {
        $effectiveMode   = $Mode
    }

    if ($effectiveMode -eq 'Quit') {
        Write-Host '   bye.' -ForegroundColor DarkGray
        Stop-Transcript | Out-Null
        return
    }

    Write-Banner ("SQL Server -> MariaDB Toolkit  -  Mode: {0}" -f $effectiveMode) 'Cyan'
    Write-Host ('  Target database   : {0}' -f (Get-TargetDatabaseName))     -ForegroundColor Gray
    Write-Host ('  DryRun            : {0}' -f $script:DryRun)               -ForegroundColor Gray
    Write-Host ('  MaxRows cap       : {0}' -f $MaxRows)                     -ForegroundColor Gray
    Write-Host ('  Skip backup       : {0}' -f $SkipBackup)                  -ForegroundColor Gray
    Write-Host ('  Transcript        : {0}' -f $script:TranscriptPath)       -ForegroundColor Gray

    Initialize-Dependencies
    Open-Connections

    switch ($effectiveMode) {
        'Migrate' {
            Initialize-LookupSets
            if (-not $script:DryRun) { Invoke-MariaBackup }
            else                     { Write-Warn2 'DryRun: skipping backup.' }
            Step1_Migrate-Tenants
            Step2_Migrate-Entities
            Step3_Migrate-Contacts
            Step4_Migrate-Receipts
            Step5_Migrate-Appointments
            Step6_Migrate-Documentation
            Step7_Migrate-ReceiptContent
            Step8_Migrate-Bookkeeping
            Step9_Migrate-Protocoll
            Step10_Verify-Relations
            Write-FinalSummary
        }
        'Verify' {
            Initialize-LookupSets
            Step10_Verify-Relations
            Write-FinalSummary
        }
        'Repair' {
            Initialize-LookupSets
            Step10_Verify-Relations  # Repair-LookupOrphans runs automatically inside
            Write-FinalSummary
        }
        'Cleanup' {
            Invoke-CleanupTarget
        }
    }
}
catch {
    Write-Err ("FATAL: {0}" -f $_.Exception.Message)
    Write-Err ("       {0}" -f $_.ScriptStackTrace)
    $script:ErrorCollection.Add([pscustomobject]@{ Table = '<global>'; Reason = 'fatal'; Key = ''; Message = $_.Exception.Message })
    try { Write-FinalSummary } catch { }
    throw
}
finally {
    try { Write-Progress -Id 1 -Completed -Activity 'Migration' } catch { }
    try { Commit-PendingBatch } catch { }
    Close-Connections
    Stop-Transcript | Out-Null
}

#endregion
