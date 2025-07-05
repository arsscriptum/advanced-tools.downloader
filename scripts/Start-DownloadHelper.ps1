#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   Start-DownloadHelper.ps1                                                     ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <codegp@icloud.com>                                         ║
#║   Code licensed under the GNU GPL v3.0. See the LICENSE file for details.      ║
#╚════════════════════════════════════════════════════════════════════════════════╝

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = "How top get the files")]
    [ValidateSet('clone', 'http', 'zip')]
    [string]$DownloadMethod = 'http'
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$IsLegacy = ($PSVersionTable.PSVersion.Major -eq 5)
if ($IsLegacy) {
    Add-Type -AssemblyName "mscorlib"
}

$SaveDataScripts = "$PSScriptRoot\Save-DataFiles.ps1"

. "$SaveDataScripts"

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BMW Advanced Tools - Download Helper"
$form.Size = New-Object System.Drawing.Size (500, 400)
$form.StartPosition = "CenterScreen"

# === GroupBox: Paths ===
$groupPaths = New-Object System.Windows.Forms.GroupBox
$groupPaths.Text = "Paths"
$groupPaths.Size = New-Object System.Drawing.Size (460, 150)
$groupPaths.Location = New-Object System.Drawing.Point (10, 10)

# Label: Temp Path
$labelTmp = New-Object System.Windows.Forms.Label
$labelTmp.Text = "Temp Path:"
$labelTmp.Size = New-Object System.Drawing.Size (70, 20)
$labelTmp.Location = New-Object System.Drawing.Point (10, 25)
$groupPaths.Controls.Add($labelTmp)

# TextBox: Temp Path
$textTmpPath = New-Object System.Windows.Forms.TextBox
$textTmpPath.Size = New-Object System.Drawing.Size (300, 20)
$textTmpPath.Location = New-Object System.Drawing.Point (80, 25)
$textTmpPath.Text = "$ENV:TEMP\BMWAdvancedTools"
$textTmpPath.Enabled = $False
$groupPaths.Controls.Add($textTmpPath)

# Button: Browse Temp
$btnBrowseTmp = New-Object System.Windows.Forms.Button
$btnBrowseTmp.Text = "..."
$btnBrowseTmp.Size = New-Object System.Drawing.Size (50, 20)
$btnBrowseTmp.Location = New-Object System.Drawing.Point (390, 25)

$groupPaths.Controls.Add($btnBrowseTmp)

# Label: Destination Path
$labelDest = New-Object System.Windows.Forms.Label
$labelDest.Text = "Path:"
$labelDest.Size = New-Object System.Drawing.Size (70, 20)
$labelDest.Location = New-Object System.Drawing.Point (10, 60)
$groupPaths.Controls.Add($labelDest)

# TextBox: Destination Path
$textDestPath = New-Object System.Windows.Forms.TextBox
$textDestPath.Size = New-Object System.Drawing.Size (300, 20)
$textDestPath.Enabled = $False
$textDestPath.Location = New-Object System.Drawing.Point (80, 60)
$groupPaths.Controls.Add($textDestPath)

# Button: Browse Destination
$btnBrowseDest = New-Object System.Windows.Forms.Button
$btnBrowseDest.Text = "..."
$btnBrowseDest.Size = New-Object System.Drawing.Size (50, 20)
$btnBrowseDest.Location = New-Object System.Drawing.Point (390, 60)
$groupPaths.Controls.Add($btnBrowseDest)

# Button: GO
$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = "GO"
$btnGo.Size = New-Object System.Drawing.Size (440, 30)
$btnGo.Location = New-Object System.Drawing.Point (10, 100)
$groupPaths.Controls.Add($btnGo)

$form.Controls.Add($groupPaths)

# === GroupBox: Status ===
$groupStatus = New-Object System.Windows.Forms.GroupBox
$groupStatus.Text = "Status"
$groupStatus.Size = New-Object System.Drawing.Size (460, 120)
$groupStatus.Location = New-Object System.Drawing.Point (10, 180)

# Label: State
$labelState = New-Object System.Windows.Forms.Label
$labelState.Text = "Ready"
$labelState.Size = New-Object System.Drawing.Size (440, 20)
$labelState.Location = New-Object System.Drawing.Point (10, 30)
$groupStatus.Controls.Add($labelState)

# ProgressBar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size (440, 25)
$progressBar.Location = New-Object System.Drawing.Point (10, 60)
$progressBar.Style = 'Continuous'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$groupStatus.Controls.Add($progressBar)

$form.Controls.Add($groupStatus)

function Show-ExceptionDetails {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$Record,
        [Parameter(Mandatory = $false)]
        [switch]$ShowStack
    )
    $formatstring = "{0}`n{1}"
    $fields = $Record.FullyQualifiedErrorId, $Record.Exception.ToString()
    $ExceptMsg = ($formatstring -f $fields)
    $Stack = $Record.ScriptStackTrace
    Write-Host "`n[ERROR] -> " -NoNewline -ForegroundColor DarkRed;
    Write-Host "$ExceptMsg`n`n" -ForegroundColor DarkYellow
    if ($ShowStack) {
        Write-Host "--stack begin--" -ForegroundColor DarkGreen
        Write-Host "$Stack" -ForegroundColor Gray
        Write-Host "--stack end--`n" -ForegroundColor DarkGreen
    }
    if ((Get-Variable -Name 'ShowExceptionDetailsTextBox' -Scope Global -ErrorAction Ignore -ValueOnly) -eq 1) {
        Show-MessageBoxException $ExceptMsg $Stack
    }


}


function Invoke-AesBinaryEncryption {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string]$InputFile,
        [Parameter(Position = 1, Mandatory = $true)]
        [string]$OutputFile,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$Password,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Encrypt', 'Decrypt')]
        [string]$Mode,
        [Parameter(Mandatory = $false)]
        [switch]$TextMode
    )

    begin {
        $shaManaged = New-Object System.Security.Cryptography.SHA256Managed
        $aesManaged = New-Object System.Security.Cryptography.AesManaged
        $aesManaged.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aesManaged.Padding = [System.Security.Cryptography.PaddingMode]::Zeros
        $aesManaged.BlockSize = 128
        $aesManaged.KeySize = 256
    }

    process {
        try {
            $aesManaged.Key = $shaManaged.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Password))

            switch ($Mode) {
                'Encrypt' {

                    $File = Get-Item -Path $InputFile -ErrorAction SilentlyContinue
                    if (!$File.FullName) {
                        Write-Error -Message "File not found!"
                        break
                    }

                    $plainBytes = [System.IO.File]::ReadAllBytes($File.FullName)

                    $encryptor = $aesManaged.CreateEncryptor()
                    $encryptedBytes = $encryptor.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
                    $encryptedBytes = $aesManaged.IV + $encryptedBytes
                    $aesManaged.Dispose()

                    if ($TextMode) {
                        Write-Host "Writing TEXT data in $OutputFile..."
                        [System.Convert]::ToBase64String($encryptedBytes) | Set-Content -Path $OutputFile -Encoding ascii -Force
                    } else {
                        [System.IO.File]::WriteAllBytes($OutputFile, $encryptedBytes)
                        (Get-Item $OutputFile).LastWriteTime = $File.LastWriteTime
                        Write-Host "File encrypted to $OutputFile"

                    }
                }

                'Decrypt' {


                    $File = Get-Item -Path $InputFile -ErrorAction SilentlyContinue
                    if (!$File.FullName) {
                        Write-Error -Message "File not found!"
                        break
                    }

                    $tmpBytes = [System.IO.File]::ReadAllBytes($File.FullName)

                    if ($TextMode) {
                        $content = Get-Content -Path $File.FullName -Encoding ascii -Force
                        $cipherBytes = [System.Convert]::FromBase64String($content)
                    } else {
                        $cipherBytes = $tmpBytes;
                    }

                    $aesManaged.IV = $cipherBytes[0..15]
                    $decryptor = $aesManaged.CreateDecryptor()
                    $decryptedBytes = $decryptor.TransformFinalBlock($cipherBytes, 16, $cipherBytes.Length - 16)
                    $aesManaged.Dispose()

                    if ($TextMode) {
                        Write-Host "Writing TEXT data in $OutputFile..."
                        [System.Text.Encoding]::ASCII.GetString($decryptedBytes).Trim([char]0) | Set-Content -Path $OutputFile -Encoding ascii -Force
                    } else {
                        [System.IO.File]::WriteAllBytes($OutputFile, $decryptedBytes)
                        (Get-Item $OutputFile).LastWriteTime = $File.LastWriteTime
                        Write-Host "File decrypted to $OutputFile"
                    }
                }
            }
        } catch {
            Write-Error $_
        }
    }

    end {
        $shaManaged.Dispose()
        $aesManaged.Dispose()
    }
}

function Get-PasswordWindow {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    Add-Type -AssemblyName System.Windows.Forms

    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Password"
    $form.Size = New-Object System.Drawing.Size (300, 150)
    $form.StartPosition = "CenterScreen"

    # Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Password:"
    $label.Left = 10
    $label.Top = 20
    $label.AutoSize = $true
    $form.Controls.Add($label)

    # TextBox
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Left = 90
    $textBox.Top = 18
    $textBox.Width = 180
    $textBox.UseSystemPasswordChar = $true
    $form.Controls.Add($textBox)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Left = 90
    $okButton.Top = 60
    $okButton.Width = 80
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Left = 190
    $cancelButton.Top = 60
    $cancelButton.Width = 80
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Show the form
    $dialogResult = $form.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}


function Test-DriveFreeSpace220MB {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Full path to check (e.g. C:\SomeFolder)')]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )


    # Validate path exists
    if (-not (Test-Path -Path $Path)) {
        Write-Error "The specified path '$Path' does not exist."
        return $false
    }

    try {
        # Get drive root from the path
        $fullPath = (Resolve-Path $Path).Path
        $driveRoot = [System.IO.Path]::GetPathRoot($fullPath)

        # Use DriveInfo instead of WMI
        $driveInfo = New-Object System.IO.DriveInfo ($driveRoot)

        if ($driveInfo -eq $null) {
            Write-Error "Unable to retrieve drive information for $driveRoot."
            return $false
        }

        # Free space in MB
        $freeMB = [math]::Round($driveInfo.AvailableFreeSpace / 1MB, 2)


        return ($freeMB -ge 220)
    }
    catch {
        Write-Error "Error while checking free space: $_"
        return $false
    }
}



# === Browse Buttons Logic ===
$btnBrowseTmp.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderDlg.ShowDialog() -eq 'OK') {
            $textTmpPath.Text = $folderDlg.SelectedPath
        }
    })

function Expand-RarFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "RAR file path")]
        [string]$RarFile,

        [Parameter(Mandatory = $true, HelpMessage = "Destination folder")]
        [string]$Destination
    )

    [string]$RootFolder = (Resolve-Path -Path "$PSScriptRoot\..").Path
    [string]$SevenZipPath = (Resolve-Path -Path "$RootFolder\7z").Path
    [string]$SevenZipExe = (Resolve-Path -Path "$SevenZipPath\7z.exe").Path

    # Validate 7z.exe exists

    if (-not (Test-Path $SevenZipExe)) {
        throw "7z.exe not found. Please install 7-Zip or specify the full path via -SevenZipExe."
    }

    # Ensure destination folder exists
    if (-not (Test-Path $Destination)) {
        New-Item -Path $Destination -ItemType Directory | Out-Null
    }

    $arguments = @(
        "x" # extract
        "`"$RarFile`"" # rar file path
        "-o`"$Destination`"" # output folder
        "-y" # auto-confirm
    )

    Write-Verbose "Running: $SevenZipExe $($arguments -join ' ')"

    $process = Start-Process -FilePath "$SevenZipExe" -ArgumentList $arguments -Wait -NoNewWindow -Passthru -WorkingDirectory "$SevenZipPath"

    if ($process.ExitCode -ne 0) {
        throw "Extraction failed. 7z.exe exited with code $($process.ExitCode)."
    }
}


function Save-AndDecrypt {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Destination folder")]
        [string]$Destination,

        [Parameter(Mandatory = $true, HelpMessage = "Password for decryption.")]
        [string]$Password
    )

    begin {
        $sourceUrl = "https://arsscriptum.github.io/files/advanced-tools/advanced-tools.aes"
        $tempPath = [System.IO.Path]::GetTempPath()
        $tempFolder = Join-Path $tempPath ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

        $downloadedFile = Join-Path $tempFolder "advanced-tools.aes"
        $decryptedFile = Join-Path $tempFolder "advanced-tools.zip"
    }

    process {
        try {
            Write-Verbose "Downloading AES file from $sourceUrl to $downloadedFile..."
            Invoke-WebRequest -Uri $sourceUrl -OutFile $downloadedFile -ErrorAction Stop

            Write-Verbose "Decrypting downloaded AES file..."
            $result = Invoke-AesBinaryEncryption -InputFile $downloadedFile -OutputFile $decryptedFile -Password $Password -Mode Decrypt

            if (Test-Path $decryptedFile) {
                Write-Verbose "Decryption succeeded. Extracting zip file to $Destination..."

                if (-not (Test-Path $Destination)) {
                    New-Item -Path $Destination -ItemType Directory -Force | Out-Null
                }

                Expand-Archive -Path $decryptedFile -DestinationPath $Destination -Force

                Write-Host "Extraction completed successfully to $Destination."
                return $true
            } else {
                Write-Error "Decryption failed. The decrypted file was not created."
                return $false
            }
        }
        catch {
            Write-Error "An error occurred: $_"
            return $false
        }
        finally {
            if (Test-Path $tempFolder) {
                Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}


$btnBrowseDest.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($folderDlg.ShowDialog() -eq 'OK') {
            $selPath = $folderDlg.SelectedPath
            $HasEnough = Test-DriveFreeSpace220MB -Path $selPath
            if (-not $HasEnough) {
                [System.Windows.Forms.MessageBox]::Show(
                    "The selected drive does not have at least 220MB of free space.",
                    "Insufficient Space",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return # Exit without setting the path
            }
            $textDestPath.Text = $selPath
        }
    })

# === GO Button Logic ===
$btnGo.Add_Click({
        $labelState.Text = "Processing..."
        $progressBar.Value = 0
        $form.Refresh()

        # Get paths from textboxes
        $tempPath = $textTmpPath.Text
        $destPath = $textDestPath.Text

        if (-not (Test-Path $tempPath)) {
            [System.IO.Directory]::CreateDirectory($tempPath)
            #[System.Windows.Forms.MessageBox]::Show("Temporary path does not exist!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            #$labelState.Text = "Error"
            #return
        }

        try {
            $btnGo.Enabled = $False
            $DestPath = $textDestPath.Text
            if (-not (Test-Path $DestPath)) {
                throw "$DestPath not found."
            }
            $labelState.Text = "Cloning repository..."
            $progressBar.Value = 20
            $form.Refresh()

            # Clone repo into temp path
            $clonePath = Join-Path -Path $tempPath -ChildPath "my.special.tools"
            if (Test-Path $clonePath) {
                Remove-Item -Path $clonePath -Recurse -Force
            }
            if ($DownloadMethod -eq 'clone') {
                $repoUrl = "https://github.com/arsscriptum/advanced-tools.git"
                $gitCmd = "git clone $repoUrl `"$clonePath`""
                $gitResult = & cmd /c $gitCmd
            } elseif ($DownloadMethod -eq 'http') {

                Write-Verbose "Preparing Save-DataFiles..."
                $TmpPath = (New-TmpDirectory).FullName
                Write-Verbose "Destination `"$TmpPath`""

                [System.Collections.ArrayList]$fList = [System.Collections.ArrayList]::new()
                [void]$fList.Add('scripts.zip')

                $BaseURL = 'https://arsscriptum.github.io/files/advanced-tools'
                Write-Verbose "Creating File List..."
                1..200 | ForEach-Object {
                    $RelativeFilePath = 'data/bmw_installer_package.rar{0:d4}.cpp' -f $_
                    [void]$fList.Add($RelativeFilePath)
                    Write-Verbose "   + file `"$RelativeFilePath`""
                }
                Write-Verbose "BaseURL `"$BaseURL`""
                Write-Verbose "Files Count $($Files.Count)"
                Write-Verbose "Destination `"$TmpPath`""

                Save-DataFiles -BaseURL $BaseURL -Files $fList -Destination "$TmpPath" -Priority 'High'

                [string]$RootFolder = (Resolve-Path -Path "$PSScriptRoot\..").Path
                [string]$SevenZipPath = (Resolve-Path -Path "$RootFolder\7z").Path
                [string]$SevenZipExe = (Resolve-Path -Path "$SevenZipPath\7z.exe").Path
                # Validate 7z.exe exists

                if (-not (Test-Path $SevenZipExe)) {
                    throw "7z.exe not found. Please install 7-Zip or specify the full path via -SevenZipExe."
                }
                $Destination = "$TmpPath"

                # Ensure destination folder exists
                if (-not (Test-Path $Destination)) {
                    New-Item -Path $Destination -ItemType Directory | Out-Null
                }
                $arguments = @(
                    "x" # extract
                    "-p`"secret`"" ## THE PASSWORD 'secret' is just the password for the zip package (additional protection). No need to try to use this in the main data files decryption, it wont work. The required password is 16 characters long
                    "`"$TmpPath\scripts.zip`"" # rar file path
                    "-o`"$Destination`"" # output folder
                    "-y" # auto-confirm
                )
                Write-Verbose "Running: $SevenZipExe $($arguments -join ' ')"

                $process = Start-Process -FilePath "$SevenZipExe" -ArgumentList $arguments -Wait -NoNewWindow -Passthru -WorkingDirectory "$SevenZipPath"

                if ($process.ExitCode -ne 0) {
                    throw "Extraction failed. 7z.exe exited with code $($process.ExitCode)."
                }
                $clonePath = $TmpPath
            } elseif ($DownloadMethod -eq 'zip') {
                Write-Error "Not fully tested and done. Duplicate of clone...."
                #Save-AndDecrypt -Destination "$clonePath" -Password "secret" ## THE PASSWORD 'secret' is just the password for the zip package (additional protection). No need to try to use this in the main data files decryption, it wont work. The required password is 16 characters long
            }

            $progressBar.Value = 60
            $labelState.Text = "Clone completed. Running Decode.ps1..."
            $form.Refresh()

            $decodeScript = Join-Path -Path $clonePath -ChildPath "scripts\Decode.ps1"
            if (-not (Test-Path $decodeScript)) {
                throw "Decode.ps1 not found."
            }
            $password = Get-PasswordWindow

            if ($password) {
                # Execute Decode.ps1
                $arguments = @(
                    '-ExecutionPolicy', 'Bypass'
                    '-File', $decodeScript
                    '-Password', "$password"
                )

                & powershell @arguments


                $progressBar.Value = 100

                $labelState.Text = "Done"
            } else {
                throw "Cancelled."
            }
            $packageFile = Join-Path -Path $clonePath -ChildPath "binsrc\bmw_installer_package.rar"
            if (-not (Test-Path $packageFile)) {
                throw "$packageFile not found."
            }

            $bmw_installer_package = Join-Path -Path $DestPath -ChildPath "bmw_installer_package.rar"

            Move-Item -LiteralPath $packageFile -Destination $DestPath -ErrorAction Stop -Force
            $expath = (Get-Command 'explorer.exe').Source
            & "$expath" "$DestPath"
            try {
                Expand-RarFile -RarFile "$bmw_installer_package" -Destination "$DestPath"
            } catch {
                [System.Windows.Forms.MessageBox]::Show("❌❌❌ WRONG PASSWORD! ❌❌❌", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Error "⚡ WRONG PASSWORD!"
            }
            Remove-Item -Path "$clonePath" -Recurse -Force -EA Ignore | Out-Null

            $msi_installer = Join-Path -Path $DestPath -ChildPath "BMW_Advanced_Tools_1.0.0_Install.msi"
            if (-not (Test-Path $msi_installer)) {
                Write-Host "$msi_installer not found."
            }
            & "$msi_installer"

            [void]$form.Close()
            [void]$form.Dispose()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $labelState.Text = "Error $_"
            $btnGo.Enabled = $True

            #Show-ExceptionDetails ($_) -ShowStack

            Remove-Item -Path "$clonePath" -Recurse -Force -EA Ignore | Out-Null

            [void]$form.Close()
            [void]$form.Dispose()

        }
    })

# Run the form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# SIG # Begin signature block
# MIIFvAYJKoZIhvcNAQcCoIIFrTCCBakCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCnkckrhDmFwQoi
# 9OOV0d1lUzRhqpHkYV2j9GRGm9X9YqCCAyQwggMgMIICCKADAgECAhAUwHk0RgY2
# mkinoG1tITQ7MA0GCSqGSIb3DQEBBQUAMCgxJjAkBgNVBAMMHUd1aWxsYXVtZSBQ
# bGFudGUgQ29kZSBTaWduaW5nMB4XDTI1MDYxNTIyMTIwNFoXDTMwMDYxNTIyMjIw
# NFowKDEmMCQGA1UEAwwdR3VpbGxhdW1lIFBsYW50ZSBDb2RlIFNpZ25pbmcwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDTebEkjrQp6qGOt6/DJ/YYJyQr
# FDvr0RQMDA9PNpAQiDloh2o9+tHl2TxFoe0fG+q6UtYTm6dx+o9dcyeABufnZQWc
# gVyDdLUGA/O/o6IuX/8lAUWNTCXEzOYxXxCoiaslywSLFA5NhFC8IlBd/JyKx/T2
# TBAZ11BBeYwZc59adW8HWAI40MHite6ywMSEEUCqrD6x4wvp7z5JknnCP1GuO/7s
# kJTGSMiX5MBVVOcFr6WmMWTQ9nZmjqxOLdl50W1tUYYX346iold0RFSMPPVUfA+A
# +3YJ+vcJlEjNmby7ZL/lGCsWeI1NSgPkGw7ruin1Quxfe3BoPOLoG17Eb6bhAgMB
# AAGjRjBEMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU5ds4r4aitwgyK/SV0pbH8fSqzAEwDQYJKoZIhvcNAQEFBQADggEBAGdx
# FoCVYN0+PNiEGeBu6pPboJLcfLCf+EfaEUVCP5QmY81HEncN6ITSgzu9N34rszqO
# 4i6daYJ2fjr5dt7w5h/hYi7YLm0brcRF9HGErZvkC4r4jggcQTmAtZuXMvYG9T3G
# HuaaJGMYSf5yHQoD0l6u5VVsPTvZwAfaSqMe1sSVDHMPkf72MiFs710u72HOgjxz
# qtTDFz+KJiAx1u4gsJM6BLTB/JYJcw9bFH7ULiRepd4SRZiiZ0++JPlcoRybJwPr
# mr/7ACvySNEJfvH778ax0wePjb/TpKT+CfGyGEkbJt+/y7/TH8llOK1ZlqBPNgfy
# FWirvfY2ou8ZyPtoKnoxggHuMIIB6gIBATA8MCgxJjAkBgNVBAMMHUd1aWxsYXVt
# ZSBQbGFudGUgQ29kZSBTaWduaW5nAhAUwHk0RgY2mkinoG1tITQ7MA0GCWCGSAFl
# AwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJ
# KoZIhvcNAQkEMSIEIDEPXvC1CGJgOIggHCQnf8T4Vsv2XfZzx72fFlOMxyMuMA0G
# CSqGSIb3DQEBAQUABIIBAAzFdBNDnpBiYvCCHbZxuvpIwsshoA8ufsOZ/fo8BUja
# eGNGT0r8c0pk5IFO47qKLVvWwwDeBPgHfLjZn6ud/L9DxcdmVRL5lGzxA1qsrJLO
# B0s27+NrvCXvgqI+7JYSAUJNnN86Nfbu1sPqWm37Ysdnn2W8WiEsOI65wmBbvnGI
# NDgIYkyeWmsSld9/JanRyKpmS/a5MEgSoUzbYMFgfB3oIOmpyvcgwDosDHOBusmN
# 5BoI1UMtECLrQVrozhTP7nz8tR/Spxm1XR+wjCUdrFG77jdMP9s+agPNVV/FNBaN
# ye3kep90D30S8ugSG+2iC7ep95LhcWhxrjfo1neu+DY=
# SIG # End signature block
