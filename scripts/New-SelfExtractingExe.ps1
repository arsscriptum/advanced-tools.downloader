#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   New-SelfExtractingExe.ps1                                                    ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <codegp@icloud.com>                                         ║
#║   Code licensed under the GNU GPL v3.0. See the LICENSE file for details.      ║
#╚════════════════════════════════════════════════════════════════════════════════╝



# Helper function
function find-program {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Alias('n')]
        [string]$Name,
        [Parameter(Mandatory = $False)]
        [Alias('e')]
        [string]$ErrorsList,
        [Parameter(Mandatory = $False)]
        [ValidateRange(1, 5)]
        [Alias('d')]
        [int]$Depth = 4,
        [Parameter(Mandatory = $False)]
        [Alias('q')]
        [switch]$Quiet,
        [Parameter(Mandatory = $False)]
        [Alias('p')]
        [switch]$PathOnly,
        [Parameter(Mandatory = $False)]
        [Alias('f')]
        [switch]$FirstMatch
    )
    begin {
        [string]$InternalErrorsListVariable = Get-RandomString -AlphaNum -Length 22

        [string[]]$ProgPaths = Get-ChildItem ENV:\ | Where Name -Match "Programs|ProgramFi" | Where Name -NotMatch "Common" | Select -ExpandProperty Value -Unique
        if ($False -eq ([string]::IsNullOrEmpty($ENV:TEXLIVE_BIN))) {
            Write-Verbose "Adding search path `"$ENV:TEXLIVE_BIN`" from environment"
            $ProgPaths += "$ENV:TEXLIVE_BIN"
        } else {
            $ProgPaths += "D:\texlive\2025\bin\windows"
            Write-Verbose "Adding HARDCODED search path `"$ENV:TEXLIVE_BIN`" FIXE THIS"
        }
        if ($False -eq ([string]::IsNullOrEmpty($ENV:CHOCOLATEY_BIN))) {
            Write-Verbose "Adding search path `"$ENV:CHOCOLATEY_BIN`" from environment"
            $ProgPaths += "$ENV:CHOCOLATEY_BIN"
        } else {
            $ChocoBin = Join-Path "$ENV:ProgramData" "chocolatey\bin"
            $ProgPaths += $ChocoBin
            Write-Verbose "Adding HARDCODED search path `"$ChocoBin`" FIXE THIS"
        }
        if ($False -eq ([string]::IsNullOrEmpty($ENV:PYTHON_SCRIPTS_BIN))) {
            Write-Verbose "Adding python install apps path `"$ENV:PYTHON_SCRIPTS_BIN`" from environment"
            $ProgPaths += "$ENV:PYTHON_SCRIPTS_BIN"
        } else {
            Write-Verbose "PYTHON_SCRIPTS_BIN is not set, will no search in PIP installed appts"
        }
        [string]$SearchFor = $Name.ToLower()
        [string]$Base = $SearchFor
        [string]$Extension = $SearchFor.Substring($SearchFor.Length - 4)
        $IsValid = $Extension[0] -eq '.'

        if ($SearchFor.EndsWith('.exe')) {
            $Base = $SearchFor.Replace('.exe', '')
        } elseif ($IsValid) {
            $Base = $SearchFor.Replace($Extension, '')
        }



    }
    process {
        try {
            $Results = @()

            $AppCmd = Get-Command -Name "$SearchFor" -ShowCommandInfo -CommandType Application -ErrorAction Ignore

            if (($AppCmd -ne $Null) -and (Test-Path -Path "$($AppCmd.Definition)" -PathType Leaf)) {
                Write-Verbose "Try [Get-Command CommandType Application] Success! $($AppCmd.Definition)"
                $AppInfo = Get-Item "$($AppCmd.Definition)"
                if ($PathOnly) {
                    return $AppInfo.FullName
                }
                return $AppInfo
            } else {
                Write-Verbose "Try [Get-Command CommandType Application] Failed"
                if ($FirstMatch) {
                    $Results = Get-ChildItem -Path $ProgPaths -Filter "$SearchFor" -Recurse -File -Depth $Depth -ErrorVariable "$InternalErrorsListVariable" -ErrorAction SilentlyContinue | Select-Object -First 1
                } else {
                    $Results = Get-ChildItem -Path $ProgPaths -Filter "$SearchFor" -Recurse -File -Depth $Depth -ErrorVariable "$InternalErrorsListVariable" -ErrorAction SilentlyContinue
                }

                if ((!$Results) -or ($Results.Count -eq 0)) {
                    Write-Verbose "Try [Get-ChildItem 1st attempt] no results found"
                    if ($FirstMatch) {
                        $Results = Get-ChildItem -Path $ProgPaths -Filter "*.exe" -Recurse -File -Depth $Depth -ErrorVariable "$InternalErrorsListVariable" -ErrorAction SilentlyContinue | Where Basename -Match $Base | Select-Object -First 1
                    } else {
                        $Results = Get-ChildItem -Path $ProgPaths -Filter "*.exe" -Recurse -File -Depth $Depth -ErrorVariable "$InternalErrorsListVariable" -ErrorAction SilentlyContinue | Where Basename -Match $Base
                    }

                    if ((!$Results) -or ($Results.Count -eq 0)) {
                        Write-Verbose "Try [Get-ChildItem 2nd attempt] no results found"
                        return $Null
                    }
                }
                Write-Verbose "Try [Get-ChildItem] Success! $($Results.Count) Matches"
            }
            if ($False -eq ([string]::IsNullOrEmpty($ErrorsList))) {
                [System.Collections.ArrayList]$VarErr = Get-Variable -Name "$InternalErrorsListVariable" -ValueOnly -ErrorAction Ignore
                if (($VarErr) -and ($VarErr.Count)) {
                    $VarErrCount = $VarErr.Count

                    Write-Verbose "[Get-ChildItem $VarErrCount Errors Occured]"
                    Write-Verbose "Use 'Get-Variable -Name `"$ErrorsList`" -ValueOnly -ErrorAction Ignore' or Trace-ErrorVariable  `"$ErrorsList`""
                    $matchingName = Get-Variable -Scope 1 | Where-Object { $_.Value -eq $ErrorsList } | Select-Object -First 1
                    try {
                        $tmpv = $MyInvocation.Line.Split('-ErrorsList ')[1].Split(' ')[0].Trim("`"").Trim("'").Trim()
                        if ($tmpv.StartsWith('$')) {
                            Write-Verbose "or 'Get-Variable -Name `"$tmpv`" -ValueOnly -ErrorAction Ignore'"
                        }
                    } catch {
                        Write-Verbose "$_"
                    }

                    if ($PSBoundParameters.ContainsKey('Quiet') -eq $False) {
                        foreach ($e in $VarErr) {
                            $cat = $e.CategoryInfo.Category
                            $dir = $e.CategoryInfo.TargetName
                            Write-Warning "[$cat] $dir"
                        }
                    }
                }
                Set-Variable -Name "$ErrorsList" -Value $VarErr -Description "Errors List Set in the Find-Program Cmdlet, Core Module" -Option AllScope -Force -Visibility Public -Scope Global | Out-Null
            }
            if ($PathOnly) {
                return $Results.FullName
            }
            return $Results
        } catch {
            throw $_
        }
    }


}
function New-SelfExtractingExe {

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $false, HelpMessage = "Path to the folder containing split parts")]
        [string]$Path
    )

    [string]$TempBuildRoot = Join-Path -Path $env:TEMP -ChildPath ([guid]::NewGuid().ToString())

    if (!$Path) {
        # Create temp build folder
        [string]$TempBuildRoot = Join-Path -Path $env:TEMP -ChildPath ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $TempBuildRoot | Out-Null
    } else {
        [string]$TempBuildRoot = $Path
        New-Item -ItemType Directory -Path $TempBuildRoot | Out-Null
    }

    # Root of your repository
    [string]$RootFolder = (Resolve-Path -Path "$PSScriptRoot\..").Path

    # Installer data folder
    [string]$InstallerDataPath = (Resolve-Path -Path "$RootFolder\installer").Path

    # Locate 7z.exe automatically
    [string]$SevenZipExe = Find-Program "7z.exe"


    # Define the archive's source folder
    [string]$SourceFolder = Join-Path -Path $TempBuildRoot -ChildPath "advanced-tools.downloader"
    New-Item -ItemType Directory -Path $SourceFolder | Out-Null

    # Copy project files except .git and installer
    Get-ChildItem -Path $RootFolder -Recurse -Force |
    Where-Object {
        $_.FullName -notmatch "\\\.git($|\\)" -and
        $_.FullName -notmatch "\\installer($|\\)"
    } |
    ForEach-Object {
        $destPath = $_.FullName.Replace($RootFolder, $SourceFolder)
        if ($_.PSIsContainer) {
            if (-not (Test-Path -Path $destPath)) {
                New-Item -Path $destPath -ItemType Directory | Out-Null
            }
        } else {
            Copy-Item -Path $_.FullName -Destination $destPath -Force
        }
    }

    # Copy installer data into source
    Copy-Item -Path (Join-Path $InstallerDataPath "config.txt") -Destination $SourceFolder -Force
    Copy-Item -Path (Join-Path $InstallerDataPath "7zSD.sfx") -Destination $SourceFolder -Force

    # Create the archive
    $ArchivePath = Join-Path $TempBuildRoot "archive.7z"
    & $SevenZipExe a -r $ArchivePath $SourceFolder | Out-Null

    # Read config.txt
    $ConfigPath = Join-Path $SourceFolder "config.txt"
    $SfxModule = Join-Path $SourceFolder "7zSD.sfx"

    # Build the SFX exe
    $OutputExe = Join-Path $RootFolder "advanced-tools.downloader.exe"
    $IsLegacy = ($PSVersionTable.PSVersion.Major -eq 5)
    if ($IsLegacy) {
        $bytesSfx = Get-Content -Path $SfxModule -Encoding Byte
        $bytesCfg = Get-Content -Path $ConfigPath -Encoding Byte
        $bytes7z = Get-Content -Path $ArchivePath -Encoding Byte
    } else {
        $bytesSfx = Get-Content -Path $SfxModule -AsByteStream
        $bytesCfg = Get-Content -Path $ConfigPath -AsByteStream
        $bytes7z = Get-Content -Path $ArchivePath -AsByteStream
    }


    [System.IO.File]::WriteAllBytes($OutputExe, $bytesSfx + $bytesCfg + $bytes7z)

    # Clean up
    Remove-Item -Path $TempBuildRoot -Recurse -Force

    Write-Host "Created SFX archive: $OutputExe"
}


New-SelfExtractingExe
