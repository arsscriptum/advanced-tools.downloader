#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Test-Download.ps1                                                            ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <codegp@icloud.com>                                         ║
#║   Code licensed under the GNU GPL v3.0. See the LICENSE file for details.      ║
#╚════════════════════════════════════════════════════════════════════════════════╝



[string]$RootFolder = (Resolve-Path -Path "$PSScriptRoot\..\..").Path
[string]$ScriptsFolder = (Resolve-Path -Path "$RootFolder\scripts").Path
[string]$SaveDataScripts = (Resolve-Path -Path "$ScriptsFolder\Save-DataFiles.ps1").Path
. "$SaveDataScripts"

function ConfirmDel { if ((Read-Host "Delete Downloaded Data Y/n").ToUpper() -eq 'Y') { return $True } else { return $False } }

function Test-MassDownload {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Write-Verbose "Preparing Save-DataFiles..."
    $TmpPath = (New-TmpDirectory).FullName
    Write-Verbose "Destination `"$TmpPath`""

    [System.Collections.ArrayList]$fList = [System.Collections.ArrayList]::new()
    [void]$fList.Add('scripts.zip')

    $BaseURL = 'http://mini:81/advanced-tools'
    Write-Verbose "Creating File List..."
    1..200 | ForEach-Object {
        $RelativeFilePath = 'data/bmw_installer_package.rar{0:d4}.cpp' -f $_
        [void]$fList.Add($RelativeFilePath)
        Write-Verbose "   + file `"$RelativeFilePath`""
    }
    Write-Verbose "BaseURL `"$BaseURL`""
    Write-Verbose "Files Count $($Files.Count)"
    Write-Verbose "Destination `"$TmpPath`""

    Save-DataFiles -BaseURL $BaseURL -Files $fList -Destination "$TmpPath" -Priority 'Foreground'
    
    if(ConfirmDel){
        Write-Host "Delete Data.."
      Remove-Item -Path "$TmpPath" -Recurse -Force -EA Ignore | Out-Null
    }


}
Test-MassDownload