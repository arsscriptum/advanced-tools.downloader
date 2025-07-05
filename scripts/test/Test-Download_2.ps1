[string]$RootFolder = (Resolve-Path -Path "$PSScriptRoot\..\..").Path
[string]$ScriptsFolder = (Resolve-Path -Path "$RootFolder\scripts").Path
[string]$SaveDataScripts = (Resolve-Path -Path "$ScriptsFolder\Save-DataFiles.ps1").Path

. "$SaveDataScripts"

function ConfirmDel { if ((Read-Host "Delete Downloaded Data Y/n").ToUpper() -eq 'Y') { return $True } else { return $False } }

Reset-GlobalJobsStats

Write-Verbose "Preparing Save-DataFiles..."
$Destination = (New-TmpDirectory).FullName
Write-Verbose "Destination `"$Destination`""
$Priority = 'Normal'
$BaseURL = 'http://mini:81/advanced-tools'
$url = 'http://mini:81/advanced-tools/data/bmw_installer_package.rar0192.cpp'
$fileName = $url.Replace("$BaseURL/", '').Replace('/', '\')
$dest = Join-Path $Destination $fileName
$desc = '{0}/{1}|{2}/{3}|{4}|{5}' -f 1,1,1,1, $dest, $url
$destDir = Split-Path -Parent $dest
if (-not (Test-Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
}
$arguments = @{
    Source = "$url"
    Destination = "$dest"
    Description = "$desc"
    TransferType = "Download"
    Asynchronous = $True
    DisplayName = "$dest"
    Priority = $Priority
    RetryTimeout = 60
    RetryInterval = 60
}

$res = Start-BitsTransfer @arguments
$JobGuid = $res.JobId.GUID
$jobptr = Get-BitsTransfer -JobId $JobGuid -ErrorAction Stop

$SmallGuid = $res.JobId.Guid.Substring(14,4)
$StartedJobs++
$NewJobGuid = $res.JobId.GUID
$log = " 🡲 Start-BitsTransfer [{0}] - {1} out of {2} in this batch, {3} out of {4} in total." -f $SmallGuid,1,1,1,1
Write-DownloadLogs "$log" -f Blue
$state = $res.JobState
$Done = if ($state -eq 'Transferred') { $True } else { $False }
Write-Host "Current State -> $state"
Write-Host "Waiting..." -n
while (!$Done) {
    Start-Sleep 1
    Write-Host ". " -n
    $jobptr = Get-BitsTransfer -JobId $JobGuid -ErrorAction Stop
    $state = $res.JobState
    $Done = if ($state -eq 'Transferred') { $True } else { $False }
}
Write-Host "`nDONE!!!`n" -f DarkGreen
$jDesc = $jobptr.Description
$DescriptionData = $jDesc.Split('|')
$jobBatchId = $DescriptionData[0]
$jobGlobalId = $DescriptionData[1]
$jobDestFile = $DescriptionData[2]
$jobUrl = $DescriptionData[3]
$log = '✔️ Job [{0}] COMPLETED. BATCH {1} TOTAL {2}' -f $SmallGuid,$jobBatchId,$jobGlobalId
Write-DownloadLogs "$log" -f Gray
$JobStats = Measure-JobStats -JobId $JobGuid
Write-JobTransferStatsLog $JobStats
Update-GlobalJobsStats $JobStats

Write-GlobalTransferStatsLog

#if (ConfirmDel) {
    Write-Host "Delete Data.."
    Remove-Item -Path "$Destination" -Recurse -Force -EA Ignore | Out-Null
#}

