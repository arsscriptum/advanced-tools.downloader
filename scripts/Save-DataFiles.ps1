#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Save-DataFiles.ps1                                                           ║
#║   Bath DOwnload using BITS                                                     ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <codegp@icloud.com>                                         ║
#║   Code licensed under the GNU GPL v3.0. See the LICENSE file for details.      ║
#╚════════════════════════════════════════════════════════════════════════════════╝


[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [bool]$LogToConsoleEnabled = $True,
    [Parameter(Mandatory = $false)]
    [bool]$LogToFileEnabled = $True
)

$ENV:DownloadLogsToFile = if ($LogToFileEnabled) { 1 } else { 0 }
$ENV:DownloadLogsToConsole = if ($LogToConsoleEnabled) { 1 } else { 0 }

$IsLegacy = ($PSVersionTable.PSVersion.Major -eq 5)
if ($IsLegacy) {
    # No need to load mscorlib; but if you want to:
    Add-Type -AssemblyName "mscorlib"
}

function Convert-Bytes {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [double]$Bytes,

        [string]$Suffix = ""
    )

    switch ($Bytes) {
        { $_ -ge 1TB } { return "{0:N2} TB$Suffix" -f ($Bytes / 1TB) }
        { $_ -ge 1GB } { return "{0:N2} GB$Suffix" -f ($Bytes / 1GB) }
        { $_ -ge 1MB } { return "{0:N2} MB$Suffix" -f ($Bytes / 1MB) }
        { $_ -ge 1KB } { return "{0:N2} KB$Suffix" -f ($Bytes / 1KB) }
        default { return "{0:N2} B$Suffix" -f $Bytes }
    }
}

function Update-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [pscustomobject]$JobStats
    )
    process {
        $CurrentStats = Get-GlobalJobsStats

        $jobCount = $CurrentStats.TotalJobs
        $minTime = $CurrentStats.MinTransferTimeSec
        $maxTime = $CurrentStats.MaxTransferTimeSec

        $totalFiles = $CurrentStats.TotalFilesTransferred
        $totalBytes = $CurrentStats.TotalBytesTransferred
        $totalTime = $CurrentStats.CurrentTotalTransferTime

        $totalFiles += $JobStats.TotalFiles
        $totalBytes += $JobStats.DownloadSize
        $totalTime += $JobStats.DownloadTime

        $globalAverageSpeedBps = [math]::Round($totalBytes / $totalTime, 2)
         


        # Human readable speed
        $globalAverageHumanSpeed = Convert-Bytes -Bytes $globalAverageSpeedBps -Suffix "/s"

        if (($minTime -eq $null) -or ($JobStats.DownloadTime -lt $minTime)) {
            $minTime = $JobStats.DownloadTime
        }
        if (($maxTime -eq $null) -or ($JobStats.DownloadTime -gt $maxTime)) {
            $maxTime = $JobStats.DownloadTime
        }

        $jobCount = $jobCount + 1 
        
        $avgTime = if ($jobCount -gt 0) { $totalTime / $jobCount } else { 0 }

        $stats = [pscustomobject]@{
            TotalJobs = $jobCount
            TotalFilesTransferred = $totalFiles
            TotalBytesTransferred = $totalBytes
            HumanReadableTotalSize = Convert-Bytes -Bytes $totalBytes
            AverageSpeed_Bps = [math]::Round($globalAverageSpeedBps, 2)
            HumanReadableAvgSpeed = Convert-Bytes -Bytes $globalAverageSpeedBps -Suffix "/s"
            AverageTransferTimeSec = [math]::Round($avgTime, 2)
            MinTransferTimeSec = $minTime
            MaxTransferTimeSec = $maxTime
            CurrentTotalTransferTime = $totalTime
        }

        # Save stats as JSON to temp file
        $path = Join-Path $env:TEMP "GlobalJobStats.json"
        $stats | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
    }
}

function Reset-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $stats = [pscustomobject]@{
        TotalJobs = 0
        TotalFilesTransferred = 0
        TotalBytesTransferred = 0
        HumanReadableTotalSize = Convert-Bytes -Bytes 0
        AverageSpeed_Bps = [math]::Round(0, 2)
        HumanReadableAvgSpeed = Convert-Bytes -Bytes 0 -Suffix "/s"
        AverageTransferTimeSec = [math]::Round(0, 2)
        MinTransferTimeSec = 0
        MaxTransferTimeSec = 0
        CurrentTotalTransferTime = 0
    }

    # Save stats as JSON to temp file
    $path = Join-Path $env:TEMP "GlobalJobStats.json"
    $stats | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
}


function Get-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $path = Join-Path $env:TEMP "GlobalJobStats.json"

    if (-not (Test-Path $path)) {
        Write-Warning "No global job stats file found at $path."
        return $null
    }

    $json = Get-Content -Path $path -Raw -Encoding UTF8
    $stats = $json | ConvertFrom-Json
    return $stats
}


function Measure-JobStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "BITS Job Id (GUID)")]
        [guid]$JobId
    )

    # Fetch the job object
    $job = Get-BitsTransfer -JobId $JobId -ErrorAction Stop

    # --- Sanity checks ---
    if ($job.JobState -ne 'Transferred') {
        throw "BITS job state is not 'Transferred'. Current state: $($job.JobState)"
    }

    if ($job.BytesTransferred -ne $job.BytesTotal) {
        throw "BytesTransferred ($($job.BytesTransferred)) does not equal BytesTotal ($($job.BytesTotal))."
    }

    if ($job.TransferCompletionTime -le $job.CreationTime) {
        throw "TransferCompletionTime is not after CreationTime."
    }

    # Compute time delta in seconds
    $durationSec = ($job.TransferCompletionTime - $job.CreationTime).TotalSeconds
    if ($durationSec -le 0) {
        throw "Calculated download time is zero or negative. Cannot compute speed."
    }
    Write-Verbose "[Measure-JobStats] BytesTransferred -> $($job.BytesTransferred)"
    Write-Verbose "[Measure-JobStats] durationSec -> $durationSec"


    # Compute speed
    $speedBps = [math]::Round($job.BytesTransferred / $durationSec, 2)
    


    # Human readable speed
    $humanSpeed = Convert-Bytes -Bytes $speedBps -Suffix "/s"

    # File path(s) - BITS supports multiple files in a job
    $localFilePath = $job.DisplayName

    $numberFilesTransfered = $job.FilesTransferred


    # Return stats
    $o = [pscustomobject]@{
        Speed = $speedBps
        HumanReadableSpeed = $humanSpeed
        DownloadTime = [math]::Round($durationSec, 2)
        DownloadSize = $job.BytesTransferred
        FilePath = $localFilePath
        TotalFiles = $numberFilesTransfered
        TotalJobs = 1
    }

    return $o
}

function Write-JobTransferStatsLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline)]
        [pscustomobject]$JobStats
    )

    process {
        if ($null -eq $JobStats) {
            Write-Host "No job statistics provided." -ForegroundColor Yellow
            return
        }

        Write-Host "===== BITS JOB TRANSFER STATISTICS =====" -ForegroundColor Cyan
        Write-Host ("File(s) Path           : {0}" -f $JobStats.FilePath) -ForegroundColor White
        Write-Host ("Download Size (bytes)  : {0}" -f $JobStats.DownloadSize) -ForegroundColor White
        Write-Host ("Download Time (sec)    : {0}" -f $JobStats.DownloadTime) -ForegroundColor White
        Write-Host ("Speed (Bytes/sec)      : {0}" -f $JobStats.Speed) -ForegroundColor White
        Write-Host ("Human-Readable Speed   : {0}" -f $JobStats.HumanReadableSpeed) -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
    }
}


function Write-GlobalTransferStatsLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $stats = Get-GlobalJobsStats

    if ($null -eq $stats) {
        Write-Host "No global transfer statistics available." -ForegroundColor Yellow
        return
    }

    Write-Host "===== GLOBAL BITS TRANSFER STATISTICS =====" -ForegroundColor Cyan
    Write-Host ("Total Jobs              : {0}" -f $stats.TotalJobs) -ForegroundColor White
    Write-Host ("Total Files Transferred : {0}" -f $stats.TotalFilesTransferred) -ForegroundColor White
    Write-Host ("Total Bytes Transferred : {0} bytes" -f $stats.TotalBytesTransferred) -ForegroundColor White
    Write-Host ("Human-Readable Size     : {0}" -f $stats.HumanReadableTotalSize) -ForegroundColor Green
    Write-Host ("Average Speed (B/s)     : {0}" -f $stats.AverageSpeed_Bps) -ForegroundColor White
    Write-Host ("Human-Readable Avg Spd  : {0}" -f $stats.HumanReadableAvgSpeed) -ForegroundColor Green
    Write-Host ("Average Transfer Time   : {0} seconds" -f $stats.AverageTransferTimeSec) -ForegroundColor White
    Write-Host ("Min Transfer Time       : {0} seconds" -f $stats.MinTransferTimeSec) -ForegroundColor White
    Write-Host ("Max Transfer Time       : {0} seconds" -f $stats.MaxTransferTimeSec) -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
}


function Write-DownloadLogs {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        $Object,

        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor,
        [switch]$NoNewline
    )

    # Compose the text exactly as Write-Host would output it
    $text = -join ($Object | ForEach-Object {
            if ($_ -is [System.Management.Automation.PSObject]) {
                $_.ToString()
            } else {
                $_
            }
        })

    # Write to console if enabled
    if ($ENV:DownloadLogsToConsole -eq "1") {
        $params = @{}
        if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
            $params.ForegroundColor = $ForegroundColor
        }
        if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
            $params.BackgroundColor = $BackgroundColor
        }
        if ($NoNewline) {
            $params.NoNewline = $true
        }
        Write-Host @params $text
    }

    # Write to file if enabled
    if ($ENV:DownloadLogsToFile -eq "1") {
        $logFile = Join-Path -Path $ENV:TEMP -ChildPath "DownloadLogs.log"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logLine = "$timestamp $text"
        Add-Content -Path $logFile -Value $logLine
    }
}


#  “define” KeyValuePair in PowerShell

function New-TmpDirectory {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param()

    # Get system temp path
    $tempPath = [System.IO.Path]::GetTempPath()

    # Generate a unique folder name
    $folderName = [System.IO.Path]::GetRandomFileName()

    # Combine to full path
    $fullPath = Join-Path -Path $tempPath -ChildPath $folderName

    # Create the directory
    $item = New-Item -ItemType Directory -Path $fullPath -Force

    return $item
}
function New-TmpFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param()

    # Use .NET method
    $tempFile = [System.IO.Path]::GetTempFileName()

    $item = New-Item -ItemType File -Path $tempFile -Force

    return $item
}

function Save-DataFiles {

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Base URL where the data is located")]
        [string]$BaseURL,
        [Parameter(Position = 1, Mandatory = $true, HelpMessage = "list of files, appended to baseurl")]
        [string[]]$Files,
        [Parameter(Position = 2, Mandatory = $false, HelpMessage = "Path to the folder containing split parts")]
        [string]$Destination,
        [Parameter(Mandatory = $false, HelpMessage = "how much files to download per batch")]
        [int]$BatchSize = 35,
        [Parameter(Mandatory = $false, HelpMessage = "how much files to download per batch")]
        [ValidateSet('Foreground', 'High', 'Low', 'Normal')]
        [string]$Priority = 'Normal'
    )

    # Define a shorthand type alias
    import-module BitsTransfer -Force
    get-BitsTransfer | Remove-BitsTransfer -Confirm:$false

    $BitsJobListType = [System.Collections.Generic.List[string]]::new()
    $UrlListType = [System.Collections.Generic.List[[System.Collections.Generic.KeyValuePair[string,bool]]]]::new()
    $ShouldExit = $False

    $UrlList = $UrlListType::new()
    $TmpDir = $Destination
    Write-DownloadLogs "[Save-DataFiles] Initialize Files Count $($Files.Count) " -f DarkYellow

    foreach ($f in $Files) {
        $RemoteUrl = '{0}/{1}' -f $BaseURL, $f
        Write-DownloadLogs "  add $RemoteUrl " -f DarkYellow
        $UrlList.Add([System.Collections.Generic.KeyValuePair[string,bool]]::new($RemoteUrl, $False))
    }

    $TotalNumberFiles = $UrlList.Count

    # How many jobs you want to run concurrently

    $CurrentState = 'Idle'

    $completedDownloadsPath = (New-TmpFile).FullName
    $ENV:CompletedDownloadsPath = $completedDownloadsPath
    $runningJobs = $BitsJobListType::new()
    $completedJobs = $BitsJobListType::new()
    $suspendedJobs = $BitsJobListType::new()
    function Get-CompletedDownloads {
        [CmdletBinding()]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        if (-not (Test-Path -LiteralPath $path)) {
            return 0
        }

        $content = Get-Content -LiteralPath $path -ErrorAction Stop
        if ([int]::TryParse($content, [ref]$null)) {
            return [int]$content
        } else {
            throw "Invalid data in $path. Expected an integer."
        }
    }

    function Add-CompletedDownloads {
        [CmdletBinding(SupportsShouldProcess)]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        $current = 0
        if (Test-Path -LiteralPath $path) {
            $current = Get-CompletedDownloads
        }

        $new = $current + 1

        if ($PSCmdlet.ShouldProcess("File: $path", "Write value $new")) {
            Set-Content -LiteralPath $path -Value $new
        }
    }

    function Reset-CompletedDownloads {
        [CmdletBinding(SupportsShouldProcess)]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        if ($PSCmdlet.ShouldProcess("File: $path", "Reset to 0")) {
            Set-Content -LiteralPath $path -Value '0'
        }
    }

    Reset-CompletedDownloads
    Reset-GlobalJobsStats
    $TotalStartedJobs = 0
    $StartedJobsInBatch = 0
    $ProcessedBatch = 0
    $EstimatedNumberOfBatches = [math]::Round(($TotalNumberFiles / $BatchSize),[System.MidpointRounding]::ToPositiveInfinity)

    $StatsEmpty = $True

    while (!$ShouldExit) {
        Start-Sleep -Milliseconds 10
        if ($runningJobs.Count -eq 0) {

            #if ($StatsEmpty -eq $False) {
            #    Write-GlobalTransferStatsLog
            #}
            $ProcessedBatch = $ProcessedBatch + 1
            $StartedJobsInBatch = 0
            Write-DownloadLogs " ★★★ Start Download in a new transfer group. This is Batch no $ProcessedBatch ★★★" -f Red

            $BatchDone = $False
            $globalId = $TotalStartedJobs

            while ($BatchDone -eq $False) {
                $c = 0
                $next = $Null
                while ($next -eq $Null) {
                    if ($False -eq ($UrlList[$c].Value)) {
                        $next = $UrlList[$c]
                    } else {
                        $c++
                        $next = $Null
                        if ($c -gt $UrlList.Count) {

                            $BatchDone = $True
                            break;
                        }
                    }
                }
                $u = $UrlList[$c]
                $IsReady = $u.Value -eq $False
                $url = $UrlList[$c].Key
                if ($IsReady) {
                    $fileName = $url.Replace("$BaseURL/", '').Replace('/', '\')
                    $dest = Join-Path $TmpDir $fileName
                    $globalId = $globalId + 1
                    Write-Verbose "[$globalId] downloading $url and saving to $dest"

                    $desc = '{0}/{1}|{2}/{3}|{4}|{5}' -f $ProcessedBatch,$EstimatedNumberOfBatches,$globalId,$TotalNumberFiles, $dest, $url
                    Write-Verbose "$desc"
                    # Safety checks
                    if (-not $url) { throw "URL is null or empty!" }
                    if (-not $dest) { throw "Destination path is null or empty!" }
                    if (-not $desc) { $desc = "BITS Transfer" }

                    # Ensure the target folder exists
                    $destDir = Split-Path -Parent $dest
                    if (-not (Test-Path $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                    }
                    
                    $arguments = @{
                        Source = "$url"
                        Description = "$desc"
                        Destination = "$dest"
                        TransferType = "Download"
                        Asynchronous = $True
                        DisplayName = "$dest"
                        Priority = $Priority
                        RetryTimeout = 60
                        RetryInterval = 60
                    }
                    $StartedJobsInBatch = $StartedJobsInBatch + 1
                    
                    $res = Start-BitsTransfer @arguments
                    $SmallGuid = $res.JobId.Guid.Substring(14,4)
                    $TotalStartedJobs++
                    $NewJobGuid = $res.JobId.GUID
                    $log = " 🡲 Start-BitsTransfer [{0}] - {1} out of {2} in this batch, {3} out of {4} in total." -f $SmallGuid,$StartedJobsInBatch,$BatchSize,$globalId,$TotalNumberFiles
                    Write-DownloadLogs "$log" -f Blue
                    $UrlList[$c] = [System.Collections.Generic.KeyValuePair[string,bool]]::new($UrlList[$c].Key, $True)
                    $runningJobs.Add($NewJobGuid)
                    if (($runningJobs.Count) -ge $BatchSize) {
                        $BatchDone = $True
                    }
                }
            }
        }

        Write-DownloadLogs "⚠️ There are currently $($runningJobs.Count) active transfers... Waiting for jobs to complete." -f DarkGreen
        while ($runningJobs.Count -gt 0) {
            Get-BitsTransfer | % {
                $JobGuid = $_.JobId.GUID
                $jobptr = Get-BitsTransfer -JobId $JobGuid -ErrorAction Stop
                $SmallGuid = $jobptr.JobId.Guid.Substring(14,4)
                $state = $jobptr.JobState
                $jDesc = $jobptr.Description
                $DescriptionData = $jDesc.Split('|')
                $jobBatchId = $DescriptionData[0]
                $jobGlobalId = $DescriptionData[1]
                $jobDestFile = $DescriptionData[2]
                $jobUrl = $DescriptionData[3]

                if ($state -eq 'Transferred') {
                    Write-Verbose "JOB $JobGuid is Transferred, Measure-JobStats..."
                    $JobStats = Measure-JobStats -JobId $JobGuid
                    #Write-JobTransferStatsLog $JobStats
                    Update-GlobalJobsStats $JobStats
                    $StatsEmpty = $False
                    Add-CompletedDownloads
                    $log = '✔️ Job [{0}] COMPLETED. BATCH {1} TOTAL {2}' -f $SmallGuid,$jobBatchId,$jobGlobalId
                    Write-DownloadLogs "$log" -f Gray
                    $jobptr | Complete-BitsTransfer
                    $completedJobs.Add($JobGuid)
                }
                foreach ($j in $completedJobs) {
                    [void]$runningJobs.Remove($j)
                }
            }
        }

        $CompletedDownloads = Get-CompletedDownloads
        if ($CompletedDownloads -ge $TotalNumberFiles) {
            $CurrentState = 'Done'
            $ShouldExit = $True

        } elseif ($ErrorOccured) {
            $CurrentState = 'Error'
            $ShouldExit = $True
        } else {
            $CurrentState = 'InProgress'
        }
    }

    Write-GlobalTransferStatsLog

}

