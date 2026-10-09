<#
.SYNOPSIS
    Compare folder pairs with WinMerge and output reports and a summary (production use).

.DESCRIPTION
    - Compares multiple folder pairs defined in a JSON config file
    - Outputs HTML / CSV reports via WinMerge (minimized, non-interactive)
    - Detects differences by SHA256 hash (own implementation, file-level error handling)
    - Per-file errors are recorded (Status=Error + reason) and processing continues
    - Unreadable folders are recorded (Status=FolderError) and the job becomes PARTIAL_ERROR
    - Log file, summary CSV, per-job file list CSV, timeout, cleanup of old reports
    - Exit code: 0=no differences / 1=differences found / 2=error

.PARAMETER ConfigPath
    Path of the JSON config. Default: compare-config.json in the script folder.

.PARAMETER VerboseFileLog
    Write one log line per file (useful to find the file that fails).

.PARAMETER WhatIf
    Dry run (WinMerge is not launched, old reports are not removed;
    hash check and file list still run).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\tools\Invoke-FolderCompare.ps1 -VerboseFileLog

.NOTES
    Version : 1.2.0 (2026-10-09)
      1.2.0 - Recover from abandoned mutex (previous run was killed)
            - Unreadable folders => Status=FolderError, Result=PARTIAL_ERROR, exit code 2
      1.1.0 - Own SHA256 implementation, per-file error handling, file list CSV
    Target  : Windows 11 / Windows PowerShell 5.1 (recommended) / PowerShell 7
    Requires: WinMerge 2.16+ (/noninteractive support)
    This file is ASCII only, so it is not affected by file encoding (BOM / Shift_JIS).
    The config JSON may contain Japanese; it is read explicitly as UTF-8.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ConfigPath = '',
    [switch]$VerboseFileLog
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '1.2.0'

# ------------------------------------------------------------
# Resolve script folder (not in param block; $PSScriptRoot may be
# empty inside param() defaults on Windows PowerShell 5.1)
# ------------------------------------------------------------
$script:ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($script:ScriptDir) -and $MyInvocation.MyCommand.Path) {
    $script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrEmpty($script:ScriptDir)) {
    $script:ScriptDir = (Get-Location).ProviderPath
}
if ([string]::IsNullOrEmpty($ConfigPath)) {
    $ConfigPath = Join-Path $script:ScriptDir 'compare-config.json'
}

# ============================================================
# Common functions
# ============================================================
$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1,-5}] {2}' -f (Get-Date), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false
    }
}

function Resolve-WinMergePath {
    param([string]$Configured)
    $candidates = @()
    if ($Configured) { $candidates += $Configured }
    if ($env:ProgramFiles)        { $candidates += (Join-Path $env:ProgramFiles 'WinMerge\WinMergeU.exe') }
    if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'WinMerge\WinMergeU.exe') }
    if ($env:LOCALAPPDATA)        { $candidates += (Join-Path $env:LOCALAPPDATA 'Programs\WinMerge\WinMergeU.exe') }
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) { return $c }
    }
    $cmd = Get-Command 'WinMergeU.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    throw 'WinMergeU.exe not found. Set winMergePath in the config file.'
}

function ConvertTo-SafeName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = '[' + [Regex]::Escape($invalid) + ']'
    return ($Name -replace $pattern, '_')
}

function Quote-Arg {
    param([string]$Value)
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Test-ExcludedPath {
    param([string]$RelativePath, [string]$FileName, [string[]]$ExcludePatterns)
    foreach ($p in $ExcludePatterns) {
        if ($RelativePath -like $p -or $FileName -like $p) { return $true }
    }
    return $false
}

function Get-FileSha256 {
    <#
        Own SHA256 implementation (replaces Get-FileHash).
        - Opens with FileShare.ReadWrite so files opened in Excel/Word can be read
        - Throws a clear exception (caller records it per file)
    #>
    param([string]$Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $bytes = $sha.ComputeHash($fs)
        return ([BitConverter]::ToString($bytes) -replace '-', '')
    }
    finally {
        if ($fs) { $fs.Dispose() }
        $sha.Dispose()
    }
}

function Get-RelativeOrFull {
    # Convert a full path under Root to a relative path (for readability)
    param([string]$Root, [string]$Path)
    if ($Path -and $Path.StartsWith($Root, [StringComparison]::OrdinalIgnoreCase)) {
        $rel = $Path.Substring($Root.Length).TrimStart('\')
        if ($rel -eq '') { return '.' }
        return $rel
    }
    return $Path
}

function Get-FileMap {
    <#
        Build a map: lower-case relative path -> { Relative, File }
        Folders are not included (files only).
        Folders that cannot be read are returned in FolderErrors
        (v1.2.0: the caller treats them as errors, not just warnings).
    #>
    param(
        [string]$Root,
        [string]$Side,
        [string[]]$IncludePatterns,
        [string[]]$ExcludePatterns,
        [bool]$Recurse
    )
    $map = @{}
    $folderErrors = New-Object System.Collections.Generic.List[object]
    $rootFull = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('\')

    $enumErrors = @()
    $files = @(Get-ChildItem -LiteralPath $rootFull -File -Recurse:$Recurse -Force `
                -ErrorAction SilentlyContinue -ErrorVariable +enumErrors)

    foreach ($e in $enumErrors) {
        $target = [string]$e.TargetObject
        $relTarget = Get-RelativeOrFull -Root $rootFull -Path $target
        $msg = $e.Exception.Message
        $folderErrors.Add([pscustomobject]@{ Side = $Side; Path = $relTarget; Message = $msg })
        Write-Log ('  Cannot read folder ({0}): {1} : {2}' -f $Side, $target, $msg) 'ERROR'
    }

    foreach ($f in $files) {
        $rel = $f.FullName.Substring($rootFull.Length).TrimStart('\')
        if ($IncludePatterns -and $IncludePatterns.Count -gt 0) {
            $hit = $false
            foreach ($ip in $IncludePatterns) { if ($f.Name -like $ip) { $hit = $true; break } }
            if (-not $hit) { continue }
        }
        if ($ExcludePatterns -and (Test-ExcludedPath -RelativePath $rel -FileName $f.Name -ExcludePatterns $ExcludePatterns)) { continue }
        $map[$rel.ToLowerInvariant()] = [pscustomobject]@{ Relative = $rel; File = $f }
    }

    return [pscustomobject]@{ Map = $map; FolderErrors = $folderErrors }
}

function New-DetailRow {
    param($Rel, $Status, $LeftFile, $RightFile, [string]$Message = '')
    [pscustomobject]@{
        Path       = $Rel
        Status     = $Status
        LeftSize   = if ($LeftFile)  { $LeftFile.Length }  else { '' }
        RightSize  = if ($RightFile) { $RightFile.Length } else { '' }
        LeftDate   = if ($LeftFile)  { $LeftFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') }  else { '' }
        RightDate  = if ($RightFile) { $RightFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        Message    = $Message
    }
}

function Compare-FolderByHash {
    <#
        Difference check independent of WinMerge language/report format.
        Different size => Different; same size => compare SHA256.
        A failure on one file does NOT stop the job (Status=Error).
        An unreadable folder is recorded as Status=FolderError.
    #>
    param(
        [string]$Left, [string]$Right,
        [string[]]$IncludePatterns, [string[]]$ExcludePatterns,
        [bool]$Recurse,
        [bool]$FileLog
    )
    $lr = Get-FileMap -Root $Left  -Side 'Left'  -IncludePatterns $IncludePatterns -ExcludePatterns $ExcludePatterns -Recurse $Recurse
    $rr = Get-FileMap -Root $Right -Side 'Right' -IncludePatterns $IncludePatterns -ExcludePatterns $ExcludePatterns -Recurse $Recurse
    $l = $lr.Map; $r = $rr.Map
    Write-Log ('  Files: Left {0} / Right {1}' -f $l.Count, $r.Count)

    $details = New-Object System.Collections.Generic.List[object]

    # Folder errors first (so they are visible at the top of the file list)
    foreach ($fe in (@($lr.FolderErrors) + @($rr.FolderErrors))) {
        $details.Add((New-DetailRow $fe.Path 'FolderError' $null $null ('[{0}] {1}' -f $fe.Side, $fe.Message)))
    }

    $keys = @(@($l.Keys) + @($r.Keys) | Sort-Object -Unique)

    foreach ($k in $keys) {
        $lf = if ($l.ContainsKey($k)) { $l[$k].File } else { $null }
        $rf = if ($r.ContainsKey($k)) { $r[$k].File } else { $null }
        $rel = if ($lf) { $l[$k].Relative } else { $r[$k].Relative }

        try {
            if ($lf -and -not $rf)      { $row = New-DetailRow $rel 'LeftOnly'  $lf $null }
            elseif ($rf -and -not $lf)  { $row = New-DetailRow $rel 'RightOnly' $null $rf }
            elseif ($lf.Length -ne $rf.Length) { $row = New-DetailRow $rel 'Different' $lf $rf }
            else {
                $lh = Get-FileSha256 -Path $lf.FullName
                $rh = Get-FileSha256 -Path $rf.FullName
                $st = if ($lh -eq $rh) { 'Same' } else { 'Different' }
                $row = New-DetailRow $rel $st $lf $rf
            }
        }
        catch {
            $msg = $_.Exception.Message
            if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
            $row = New-DetailRow $rel 'Error' $lf $rf $msg
            Write-Log ('  File ERROR: {0} : {1}' -f $rel, $msg) 'WARN'
        }
        $details.Add($row)
        if ($FileLog) { Write-Log ('    {0,-11} {1}' -f $row.Status, $rel) }
    }

    $fileRows = @($details | Where-Object { $_.Status -ne 'FolderError' })
    return [pscustomobject]@{
        Total        = $fileRows.Count
        Same         = @($fileRows | Where-Object { $_.Status -eq 'Same' }).Count
        Different    = @($fileRows | Where-Object { $_.Status -eq 'Different' }).Count
        LeftOnly     = @($fileRows | Where-Object { $_.Status -eq 'LeftOnly' }).Count
        RightOnly    = @($fileRows | Where-Object { $_.Status -eq 'RightOnly' }).Count
        Errors       = @($fileRows | Where-Object { $_.Status -eq 'Error' }).Count
        FolderErrors = @($details  | Where-Object { $_.Status -eq 'FolderError' }).Count
        Details      = $details
    }
}

function Invoke-WinMergeReport {
    param(
        [string]$WinMerge,
        [string]$Left, [string]$Right,
        [string]$LeftDesc, [string]$RightDesc,
        [string]$ReportPath,
        [string]$Filter,
        [bool]$Recurse,
        [int]$TimeoutSec
    )
    $argList = New-Object System.Collections.Generic.List[string]
    if ($Recurse) { $argList.Add('/r') }
    $argList.AddRange([string[]]@('/u', '/minimize', '/noninteractive', '/noprefs'))
    if ($Filter) { $argList.Add('/f'); $argList.Add((Quote-Arg $Filter)) }
    $argList.Add('/dl'); $argList.Add((Quote-Arg $LeftDesc))
    $argList.Add('/dr'); $argList.Add((Quote-Arg $RightDesc))
    $argList.Add((Quote-Arg $Left))
    $argList.Add((Quote-Arg $Right))
    $argList.Add('/or'); $argList.Add((Quote-Arg $ReportPath))

    Write-Log ('  WinMerge run: {0} {1}' -f $WinMerge, ($argList -join ' '))

    $proc = Start-Process -FilePath $WinMerge -ArgumentList $argList.ToArray() -PassThru -WindowStyle Minimized
    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        try { $proc.Kill() } catch { }
        throw ('WinMerge timed out ({0} sec): {1}' -f $TimeoutSec, $ReportPath)
    }
    if (-not (Test-Path -LiteralPath $ReportPath)) {
        throw ('Report was not created: {0} (check WinMerge version/arguments)' -f $ReportPath)
    }
}

function Remove-OldReports {
    param([string]$Dir, [int]$RetentionDays)
    if ($RetentionDays -le 0) { return }
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Dir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' -and $_.LastWriteTime -lt $limit } |
        ForEach-Object {
            Write-Log ('Removed old report folder: {0}' -f $_.FullName)
            Remove-Item -LiteralPath $_.FullName -Recurse -Force
        }
}

function Get-Prop {
    param($Obj, [string]$Name, $Default)
    if ($null -ne $Obj -and $Obj.PSObject.Properties[$Name] -and $null -ne $Obj.$Name -and "$($Obj.$Name)" -ne '') {
        return $Obj.$Name
    }
    return $Default
}

function Enter-SingleInstance {
    <#
        Acquire the global mutex.
        v1.2.0: If the previous run was killed (Task Scheduler stop, Ctrl+C,
        process kill, PC shutdown), the mutex is "abandoned". WaitOne() then
        throws AbandonedMutexException although ownership IS acquired.
        This is treated as a successful acquisition with a warning.
    #>
    $m = New-Object System.Threading.Mutex($false, 'Global\Invoke-FolderCompare')
    $acquired = $false
    try {
        $acquired = $m.WaitOne(0)
    }
    catch [System.Threading.AbandonedMutexException] {
        $acquired = $true
        Write-Log 'Previous run did not finish normally (abandoned lock). Lock was recovered and processing continues.' 'WARN'
    }
    if (-not $acquired) {
        $m.Dispose()
        throw 'Another compare process is already running.'
    }
    return $m
}

# ============================================================
# Main
# ============================================================
$exitCode = 0
$mutex = $null
$runStamp = '{0:yyyyMMdd_HHmmss}' -f (Get-Date)

try {
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw ('Config file not found: {0}' -f $ConfigPath)
    }
    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $outputRoot    = [string](Get-Prop $config 'outputRoot' (Join-Path $script:ScriptDir 'reports'))
    $retentionDays = [int](Get-Prop $config 'retentionDays' 30)
    $timeoutSec    = [int](Get-Prop $config 'timeoutSec' 1800)
    $formats       = @(Get-Prop $config 'reportFormats' @('html', 'csv'))
    $useHashCheck  = [bool](Get-Prop $config 'hashCheck' $true)

    $runDir = Join-Path $outputRoot $runStamp
    New-Item -ItemType Directory -Path $runDir -Force -WhatIf:$false | Out-Null
    $script:LogFile = Join-Path $runDir 'compare.log'

    Write-Log ('===== Folder compare START ({0}) =====' -f $runStamp)
    Write-Log ('Script    : {0} (ver {1})' -f $MyInvocation.MyCommand.Path, $script:ScriptVersion)
    Write-Log ('Config    : {0}' -f $ConfigPath)
    Write-Log ('Output    : {0}' -f $runDir)
    Write-Log ('PowerShell: {0}' -f $PSVersionTable.PSVersion)
    if ($WhatIfPreference) { Write-Log 'Mode      : WhatIf (WinMerge is not launched)' 'WARN' }

    $winMerge = Resolve-WinMergePath -Configured ([string](Get-Prop $config 'winMergePath' ''))
    $ver = (Get-Item -LiteralPath $winMerge).VersionInfo.ProductVersion
    Write-Log ('WinMerge  : {0} (ver {1})' -f $winMerge, $ver)

    $mutex = Enter-SingleInstance

    $summary = New-Object System.Collections.Generic.List[object]

    foreach ($job in @($config.jobs)) {
        $name      = [string]$job.name
        $safeName  = ConvertTo-SafeName $name
        $recurse   = [bool](Get-Prop $job 'recurse' $true)
        $filter    = [string](Get-Prop $job 'filter' '')
        $include   = @(Get-Prop $job 'include' @())
        $exclude   = @(Get-Prop $job 'exclude' @())
        $leftDesc  = [string](Get-Prop $job 'leftDesc'  $job.left)
        $rightDesc = [string](Get-Prop $job 'rightDesc' $job.right)

        $row = [ordered]@{
            Job = $name; Left = $job.left; Right = $job.right
            Result = ''; Total = ''; Same = ''; Different = ''; LeftOnly = ''; RightOnly = ''
            Errors = ''; FolderErrors = ''
            Reports = ''; FileList = ''; Message = ''
        }
        Write-Log ('--- [{0}] start' -f $name)

        try {
            foreach ($p in @($job.left, $job.right)) {
                if ([string]::IsNullOrEmpty($p) -or -not (Test-Path -LiteralPath $p -PathType Container)) {
                    throw ('Folder not found: {0}' -f $p)
                }
            }

            # 1) WinMerge reports (a failure here does not stop the hash check)
            $reports = @()
            foreach ($fmt in $formats) {
                $reportPath = Join-Path $runDir ('{0}.{1}' -f $safeName, ([string]$fmt).ToLowerInvariant())
                if ($PSCmdlet.ShouldProcess($reportPath, 'WinMerge report output')) {
                    try {
                        Invoke-WinMergeReport -WinMerge $winMerge -Left $job.left -Right $job.right `
                            -LeftDesc $leftDesc -RightDesc $rightDesc -ReportPath $reportPath `
                            -Filter $filter -Recurse $recurse -TimeoutSec $timeoutSec
                        $reports += (Split-Path $reportPath -Leaf)
                        Write-Log ('  Report: {0}' -f $reportPath)
                    }
                    catch {
                        Write-Log ('  WinMerge ERROR: {0}' -f $_.Exception.Message) 'ERROR'
                        $row.Message = 'WinMerge: ' + $_.Exception.Message
                        $exitCode = 2
                    }
                }
            }
            $row.Reports = $reports -join ';'

            # 2) Hash-based difference check (all files listed, errors per file / folder)
            if ($useHashCheck) {
                $r = Compare-FolderByHash -Left $job.left -Right $job.right `
                        -IncludePatterns $include -ExcludePatterns $exclude -Recurse $recurse `
                        -FileLog ([bool]$VerboseFileLog)
                $row.Total = $r.Total; $row.Same = $r.Same; $row.Different = $r.Different
                $row.LeftOnly = $r.LeftOnly; $row.RightOnly = $r.RightOnly
                $row.Errors = $r.Errors; $row.FolderErrors = $r.FolderErrors

                $fileList = Join-Path $runDir ('{0}_files.csv' -f $safeName)
                $r.Details | Export-Csv -LiteralPath $fileList -NoTypeInformation -Encoding UTF8 -WhatIf:$false
                $row.FileList = Split-Path $fileList -Leaf

                if (($r.Errors + $r.FolderErrors) -gt 0) {
                    # v1.2.0: unreadable folders are also treated as errors
                    $row.Result = 'PARTIAL_ERROR'
                    $exitCode = 2
                    if ($r.FolderErrors -gt 0) {
                        $folderMsg = ('{0} folder(s) could not be read; files in them were NOT compared' -f $r.FolderErrors)
                        $row.Message = if ($row.Message) { $row.Message + ' / ' + $folderMsg } else { $folderMsg }
                    }
                } elseif (($r.Different + $r.LeftOnly + $r.RightOnly) -gt 0) {
                    $row.Result = 'DIFF'
                    if ($exitCode -lt 1) { $exitCode = 1 }
                } else {
                    $row.Result = 'SAME'
                }
                Write-Log ('  Result: {0} (Total {1} / Same {2} / Different {3} / LeftOnly {4} / RightOnly {5} / Error {6} / FolderError {7})' -f `
                    $row.Result, $r.Total, $r.Same, $r.Different, $r.LeftOnly, $r.RightOnly, $r.Errors, $r.FolderErrors)
            } else {
                $row.Result = 'REPORT_ONLY'
            }
        }
        catch {
            $row.Result = 'ERROR'
            $row.Message = $_.Exception.Message
            $exitCode = 2
            Write-Log ('  [{0}] ERROR: {1}' -f $name, $_.Exception.Message) 'ERROR'
            Write-Log ('  at: {0}' -f $_.InvocationInfo.PositionMessage) 'ERROR'
        }
        $summary.Add([pscustomobject]$row)
    }

    $summaryPath = Join-Path $runDir 'summary.csv'
    $summary | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false
    Write-Log ('Summary   : {0}' -f $summaryPath)

    Remove-OldReports -Dir $outputRoot -RetentionDays $retentionDays
}
catch {
    $exitCode = 2
    Write-Log ('FATAL: {0}' -f $_.Exception.Message) 'ERROR'
}
finally {
    if ($mutex) {
        try { $mutex.ReleaseMutex() } catch { }
        $mutex.Dispose()
    }
    $resultText = @{ 0 = 'No differences'; 1 = 'Differences found'; 2 = 'Error occurred' }[$exitCode]
    Write-Log ('===== Folder compare END: {0} (ExitCode={1}) =====' -f $resultText, $exitCode)
}

exit $exitCode
