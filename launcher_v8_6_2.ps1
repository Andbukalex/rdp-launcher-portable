Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-AppRoot {
    try {
        $mainModule = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $mainName = [System.IO.Path]::GetFileName($mainModule).ToLowerInvariant()

        if ($mainModule -and $mainName -ne "powershell.exe" -and $mainName -ne "pwsh.exe") {
            return [System.IO.Path]::GetDirectoryName($mainModule)
        }
    } catch {}

    if ($PSScriptRoot) {
        return $PSScriptRoot
    }

    if ($MyInvocation.MyCommand.Path) {
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    return (Get-Location).Path
}

$scriptRoot = Get-AppRoot
$configPath = Join-Path $scriptRoot "config.json"
$defaultRdpPath = Join-Path $scriptRoot "rdp"

$rdpItems = @()
$rdpData = @()
$favorites = @()
$statusCache = @{}

function Find-RdpFolder {
    param([string]$configuredFolder)

    # 1) Config-Pfad verwenden, wenn er existiert
    if ($configuredFolder -and (Test-Path $configuredFolder)) {
        return [string]$configuredFolder
    }

    # 2) Portable Standard-Struktur: RDP-Launcher.exe + rdp\
    if (Test-Path $defaultRdpPath) {
        return [string]$defaultRdpPath
    }

    # 3) Falls die .rdp-Dateien direkt neben der EXE/dem Script liegen
    try {
        $localRdpFiles = @(Get-ChildItem -Path $scriptRoot -Filter *.rdp -File -ErrorAction SilentlyContinue)
        if ($localRdpFiles.Count -gt 0) {
            return [string]$scriptRoot
        }
    } catch {}

    # 4) Nahe liegende Unterordner durchsuchen, zuerst Ordner mit "rdp" im Namen
    try {
        $dirs = @(Get-ChildItem -Path $scriptRoot -Directory -ErrorAction SilentlyContinue | Sort-Object @{Expression={ if ($_.Name -match "rdp") { 0 } else { 1 } }}, Name)
        foreach ($dir in $dirs) {
            $files = @(Get-ChildItem -Path $dir.FullName -Filter *.rdp -File -ErrorAction SilentlyContinue)
            if ($files.Count -gt 0) {
                return [string]$dir.FullName
            }
        }
    } catch {}

    # 5) Fallback: lokalen rdp-Ordner anlegen, damit die portable Struktur sofort klar ist
    try {
        if (!(Test-Path $defaultRdpPath)) {
            New-Item -ItemType Directory -Path $defaultRdpPath | Out-Null
        }
    } catch {}

    return [string]$defaultRdpPath
}

function Load-Config {
    $result = @{
        RdpFolder = $defaultRdpPath
        Favorites = @()
        DefaultSshUser = ""
        LastUsedFile = ""
        History = @()
    }

    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json

            if ($config.rdpFolder) {
                $result.RdpFolder = [string]$config.rdpFolder
            }

            if ($config.favorites) {
                $result.Favorites = @($config.favorites)
            }

            if ($config.defaultSshUser) {
                $result.DefaultSshUser = [string]$config.defaultSshUser
            }

            if ($config.lastUsedFile) {
                $result.LastUsedFile = [string]$config.lastUsedFile
            }

            if ($config.history) {
                $result.History = @($config.history)
            }
        } catch {}
    }
    $result.RdpFolder = Find-RdpFolder -configuredFolder $result.RdpFolder
    return $result
}

function Save-Config {
    param(
        [string]$folder,
        [array]$favoriteFiles,
        [string]$defaultSshUser,
        [string]$lastUsedFile,
        [array]$history
    )

    $config = @{
        rdpFolder = $folder
        favorites = @($favoriteFiles)
        defaultSshUser = $defaultSshUser
        lastUsedFile = $lastUsedFile
        history = @($history)
    } | ConvertTo-Json

    $configDir = Split-Path $configPath

    if (!(Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir | Out-Null
    }

    Set-Content -Path $configPath -Value $config -Encoding UTF8
}

function Export-ConfigFile {
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Config exportieren"
    $dialog.Filter = "JSON-Datei (*.json)|*.json|Alle Dateien (*.*)|*.*"
    $dialog.FileName = "rdp-launcher-config.json"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Save-Config -folder $script:basePath -favoriteFiles $script:favorites -defaultSshUser $script:defaultSshUser -lastUsedFile $script:lastUsedFile -history $script:history
        Copy-Item -Path $configPath -Destination $dialog.FileName -Force
        $lastCheckLabel.Text = "Config exportiert"
    }
}

function Import-ConfigFile {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Config importieren"
    $dialog.Filter = "JSON-Datei (*.json)|*.json|Alle Dateien (*.*)|*.*"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $imported = Get-Content $dialog.FileName -Raw | ConvertFrom-Json

            if ($imported.rdpFolder) {
                $script:basePath = [string]$imported.rdpFolder
            }

            if ($imported.favorites) {
                $script:favorites = @($imported.favorites)
            } else {
                $script:favorites = @()
            }

            Save-Config -folder $script:basePath -favoriteFiles $script:favorites -defaultSshUser $script:defaultSshUser -lastUsedFile $script:lastUsedFile -history $script:history
            $searchBox.Text = ""
            Load-RdpCards
            $lastCheckLabel.Text = "Config importiert"
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Config konnte nicht importiert werden.",
                "RDP Launcher",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

function Format-Name {
    param([string]$name)
    $name = $name -replace "-", " "
    return (Get-Culture).TextInfo.ToTitleCase($name)
}

function Get-RdpValue {
    param(
        [string]$rdpFilePath,
        [string]$key
    )

    try {
        $pattern = "$key" + ":s:*"
        $line = Get-Content $rdpFilePath | Where-Object { $_ -like $pattern } | Select-Object -First 1
        if ($line) {
            return ($line -replace ("^" + [regex]::Escape($key) + ":s:"), "").Trim()
        }
    } catch {}

    return ""
}

function Get-RdpAddress {
    param([string]$rdpFilePath)
    return Get-RdpValue -rdpFilePath $rdpFilePath -key "full address"
}

function Get-RdpUsername {
    param([string]$rdpFilePath)
    return Get-RdpValue -rdpFilePath $rdpFilePath -key "username"
}

function Test-IsInvalidRdpAddress {
    param([string]$hostName)

    if ([string]::IsNullOrWhiteSpace($hostName)) {
        return $true
    }

    $value = $hostName.Trim()

    $placeholders = @(
        "unbekannt",
        "DEIN-RECHNERNAME-ODER-IP",
        "RECHNERNAME",
        "IP",
        "HOST",
        "SERVER"
    )

    foreach ($placeholder in $placeholders) {
        if ($value -ieq $placeholder) {
            return $true
        }
    }

    if ($value -match "(?i)DEIN-.*(RECHNER|HOST|SERVER|IP)" -or $value -match "(?i)(RECHNERNAME|ODER-IP)") {
        return $true
    }

    return $false
}

function Test-RdpHost {
    param([string]$hostName)

    if (Test-IsInvalidRdpAddress $hostName) {
        return "INVALID"
    }

    $pingTarget = $hostName.Trim()
    if ($pingTarget -match "^\[(.+)\]:(\d+)$") {
        $pingTarget = $matches[1]
    } elseif ($pingTarget -match "^([^:]+):(\d+)$") {
        $pingTarget = $matches[1]
    }

    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $reply = $ping.Send($pingTarget, 700)

        if ($reply.Status -eq "Success") {
            return "ONLINE"
        } else {
            return "OFFLINE"
        }
    } catch {
        return "OFFLINE"
    }
}

function Set-StatusView {
    param(
        [System.Windows.Forms.Panel]$dot,
        [System.Windows.Forms.Label]$label,
        [string]$status
    )

    if ($status -eq "ONLINE") {
        $dot.BackColor = [System.Drawing.Color]::FromArgb(20,160,70)
        $label.Text = "Status: online"
        $label.ForeColor = [System.Drawing.Color]::FromArgb(20,130,60)
    } elseif ($status -eq "OFFLINE") {
        $dot.BackColor = [System.Drawing.Color]::FromArgb(190,50,40)
        $label.Text = "Status: offline"
        $label.ForeColor = [System.Drawing.Color]::FromArgb(180,50,40)
    } elseif ($status -eq "INVALID") {
        $dot.BackColor = [System.Drawing.Color]::FromArgb(230,145,35)
        $label.Text = "Status: ungültig"
        $label.ForeColor = [System.Drawing.Color]::FromArgb(190,100,20)
    } elseif ($status -eq "CHECKING") {
        $dot.BackColor = [System.Drawing.Color]::FromArgb(235,170,20)
        $label.Text = "Status: Check..."
        $label.ForeColor = [System.Drawing.Color]::FromArgb(150,105,20)
    } else {
        $dot.BackColor = [System.Drawing.Color]::FromArgb(130,130,130)
        $label.Text = "Status: nicht gecheckt"
        $label.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
    }
}

function Set-ClipboardText {
    param(
        [string]$text,
        [string]$message
    )

    if ([string]::IsNullOrWhiteSpace($text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Kein Wert zum Kopieren vorhanden.",
            "RDP Launcher",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    [System.Windows.Forms.Clipboard]::SetText($text)
    $lastCheckLabel.Text = $message
}

function Add-HistoryEntry {
    param([string]$fileName)

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        return
    }

    $script:lastUsedFile = $fileName
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp | $fileName"

    $script:history = @($entry) + @($script:history | Where-Object { $_ -notlike "*| $fileName" })
    $script:history = @($script:history | Select-Object -First 10)

    Save-Config -folder $script:basePath -favoriteFiles $script:favorites -defaultSshUser $script:defaultSshUser -lastUsedFile $script:lastUsedFile -history $script:history
}

function Show-History {
    if ($script:history.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Noch keine History vorhanden.",
            "History",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    [System.Windows.Forms.MessageBox]::Show(
        ($script:history -join "`n"),
        "History",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

function Start-RdpFile {
    param([string]$fileName)

    $path = Join-Path $basePath $fileName

    if (Test-Path $path) {
        Add-HistoryEntry -fileName $fileName
        Start-Process -FilePath "mstsc.exe" -ArgumentList ('"{0}"' -f $path)

        # Nur die Sortierung/Markierung aktualisieren, aber bekannte Statuswerte behalten.
        # Ein kompletter Reload setzte bisher alle Statusanzeigen auf "nicht gecheckt" zurück.
        foreach ($entry in $script:rdpData) {
            $entry.Group = Get-RdpGroup -fileName $entry.FileName -displayName $entry.DisplayName -isFavorite $entry.IsFavorite
        }
        Render-FilteredCards
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Datei nicht gefunden:`n$path",
            "RDP Launcher",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Open-RdpInEditor {
    param([string]$fileName)

    $path = Join-Path $basePath $fileName

    if (Test-Path $path) {
        Start-Process -FilePath "notepad.exe" -ArgumentList ('"{0}"' -f $path)
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Datei nicht gefunden:`n$path",
            "Editor",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Open-RdpInExplorer {
    param([string]$fileName)

    $path = Join-Path $basePath $fileName

    if (Test-Path $path) {
        Start-Process "explorer.exe" -ArgumentList "/select,`"$path`""
    } elseif (Test-Path $basePath) {
        Start-Process "explorer.exe" -ArgumentList "`"$basePath`""
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Ordner nicht gefunden:`n$basePath",
            "Explorer",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Toggle-Favorite {
    param([string]$fileName)

    if ($script:favorites -contains $fileName) {
        $script:favorites = @($script:favorites | Where-Object { $_ -ne $fileName })
    } else {
        $script:favorites += $fileName
    }

    Save-Config -folder $script:basePath -favoriteFiles $script:favorites -defaultSshUser $script:defaultSshUser -lastUsedFile $script:lastUsedFile -history $script:history

    foreach ($entry in $script:rdpData) {
        $entry.IsFavorite = ($script:favorites -contains $entry.FileName)
        $entry.Group = Get-RdpGroup -fileName $entry.FileName -displayName $entry.DisplayName -isFavorite $entry.IsFavorite
    }

    Render-FilteredCards
}

function Get-RdpGroup {
    param(
        [string]$fileName,
        [string]$displayName,
        [bool]$isFavorite
    )

    if ($script:lastUsedFile -eq $fileName) {
        return "Zuletzt genutzt"
    }

    if ($isFavorite) {
        return "Favoriten"
    }

    $text = ($fileName + " " + $displayName).ToLower()

    if ($text -like "*server*" -or $text -like "*dev*") {
        return "Server"
    }

    if ($text -like "*test*" -or $text -like "*android*") {
        return "Testsysteme"
    }

    if ($text -like "*linux*" -or $text -like "*mint*") {
        return "Linux"
    }

    if ($text -like "*mini*" -or $text -like "*pc*" -or $text -like "*rechner*") {
        return "Rechner"
    }

    return "Weitere Verbindungen"
}


function Shorten-PathForDisplay {
    param(
        [string]$Path,
        [int]$MaxLen = 62
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    if ($Path.Length -le $MaxLen) {
        return $Path
    }

    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        $leaf = Split-Path $Path -Leaf
        $parentPath = Split-Path $Path -Parent
        $parent = Split-Path $parentPath -Leaf

        if ([string]::IsNullOrWhiteSpace($parent)) {
            return "$root...\$leaf"
        }

        return "$root...\$parent\$leaf"
    } catch {
        return "..." + $Path.Substring([Math]::Max(0, $Path.Length - $MaxLen + 3))
    }
}

function Update-FolderLabel {
    param([string]$Path)

    $displayPath = Shorten-PathForDisplay -Path $Path -MaxLen 62
    $folderLabel.Text = "RDP-Ordner: $displayPath"

    if ($script:folderToolTip) {
        $script:folderToolTip.SetToolTip($folderLabel, $Path)
    }
}

$config = Load-Config
$basePath = $config.RdpFolder
$favorites = @($config.Favorites)
$defaultSshUser = $config.DefaultSshUser
$lastUsedFile = $config.LastUsedFile
$history = @($config.History)

# Die automatisch gefundene portable Struktur direkt in config.json sichern.
Save-Config -folder $basePath -favoriteFiles $favorites -defaultSshUser $defaultSshUser -lastUsedFile $lastUsedFile -history $history

$form = New-Object System.Windows.Forms.Form
$form.Text = "RDP Launcher"
$form.Size = New-Object System.Drawing.Size(960,760)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(960,760)
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

$title = New-Object System.Windows.Forms.Label
$title.Text = "RDP Launcher"
$title.Font = New-Object System.Drawing.Font("Segoe UI",20,[System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(32,20)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Lokale Remote-Desktop-Verbindungen · Version 8.6.2 Komfort"
$subtitle.Font = New-Object System.Drawing.Font("Segoe UI",10)
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(90,90,90)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(35,62)

$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Text = "by Andreas Husemann"
$authorLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$authorLabel.ForeColor = [System.Drawing.Color]::FromArgb(140,140,140)
$authorLabel.AutoSize = $true
$authorLabel.Location = New-Object System.Drawing.Point(35,83)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "RDP-Ordner: $(Shorten-PathForDisplay -Path $basePath -MaxLen 62)"
$folderLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$folderLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
$folderLabel.AutoSize = $false
$folderLabel.Width = 850
$folderLabel.Height = 20
$folderLabel.Location = New-Object System.Drawing.Point(35,106)

$script:folderToolTip = New-Object System.Windows.Forms.ToolTip
$script:folderToolTip.SetToolTip($folderLabel, $basePath)

$lastCheckLabel = New-Object System.Windows.Forms.Label
$lastCheckLabel.Text = "Letzter Check: -"
$lastCheckLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$lastCheckLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
$lastCheckLabel.AutoSize = $true
$lastCheckLabel.Location = New-Object System.Drawing.Point(35,132)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Suche:"
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI",9)
$searchLabel.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
$searchLabel.AutoSize = $true
$searchLabel.Location = New-Object System.Drawing.Point(35,248)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Size = New-Object System.Drawing.Size(250,24)
$searchBox.Location = New-Object System.Drawing.Point(88,245)
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI",9)

$clearSearchButton = New-Object System.Windows.Forms.Button
$clearSearchButton.Text = "X"
$clearSearchButton.Size = New-Object System.Drawing.Size(32,24)
$clearSearchButton.Location = New-Object System.Drawing.Point(345,244)
$clearSearchButton.Font = New-Object System.Drawing.Font("Segoe UI",8)

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Text = ""
$countLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$countLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
$countLabel.AutoSize = $true
$countLabel.Location = New-Object System.Drawing.Point(390,248)

$changeFolderButton = New-Object System.Windows.Forms.Button
$changeFolderButton.Text = "RDP-Ordner wählen"
$changeFolderButton.Size = New-Object System.Drawing.Size(155,30)
$changeFolderButton.Location = New-Object System.Drawing.Point(405,145)
$changeFolderButton.Font = New-Object System.Drawing.Font("Segoe UI",9)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Aktualisieren"
$refreshButton.Size = New-Object System.Drawing.Size(115,30)
$refreshButton.Location = New-Object System.Drawing.Point(575,145)
$refreshButton.Font = New-Object System.Drawing.Font("Segoe UI",9)

$statusButton = New-Object System.Windows.Forms.Button
$statusButton.Text = "Status checken"
$statusButton.Size = New-Object System.Drawing.Size(130,30)
$statusButton.Location = New-Object System.Drawing.Point(705,145)
$statusButton.Font = New-Object System.Drawing.Font("Segoe UI",9)

$autoCheckBox = New-Object System.Windows.Forms.CheckBox
$autoCheckBox.Text = "Auto-Check aktiv"
$autoCheckBox.Size = New-Object System.Drawing.Size(140,22)
$autoCheckBox.Location = New-Object System.Drawing.Point(430,214)
$autoCheckBox.Font = New-Object System.Drawing.Font("Segoe UI",8)
$autoCheckBox.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

$toolsLabel = New-Object System.Windows.Forms.Label
$toolsLabel.Text = "Tools:"
$toolsLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$toolsLabel.ForeColor = [System.Drawing.Color]::FromArgb(130,130,130)
$toolsLabel.AutoSize = $true
$toolsLabel.Location = New-Object System.Drawing.Point(365,184)

$settingsLabel = New-Object System.Windows.Forms.Label
$settingsLabel.Text = "Settings:"
$settingsLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$settingsLabel.ForeColor = [System.Drawing.Color]::FromArgb(130,130,130)
$settingsLabel.AutoSize = $true
$settingsLabel.Location = New-Object System.Drawing.Point(365,216)

$sshUserLabel = New-Object System.Windows.Forms.Label
$sshUserLabel.Text = "SSH-User:"
$sshUserLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
$sshUserLabel.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
$sshUserLabel.AutoSize = $true
$sshUserLabel.Location = New-Object System.Drawing.Point(585,216)

$sshUserBox = New-Object System.Windows.Forms.TextBox
$sshUserBox.Size = New-Object System.Drawing.Size(115,22)
$sshUserBox.Location = New-Object System.Drawing.Point(655,213)
$sshUserBox.Font = New-Object System.Drawing.Font("Segoe UI",8)
$sshUserBox.Text = $script:defaultSshUser

$saveSshUserButton = New-Object System.Windows.Forms.Button
$saveSshUserButton.Text = "Save"
$saveSshUserButton.Size = New-Object System.Drawing.Size(50,22)
$saveSshUserButton.Location = New-Object System.Drawing.Point(780,213)
$saveSshUserButton.Font = New-Object System.Drawing.Font("Segoe UI",8)

$historyButton = New-Object System.Windows.Forms.Button
$historyButton.Text = "History"
$historyButton.Size = New-Object System.Drawing.Size(90,24)
$historyButton.Location = New-Object System.Drawing.Point(670,178)
$historyButton.Font = New-Object System.Drawing.Font("Segoe UI",8)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Config export"
$exportButton.Size = New-Object System.Drawing.Size(120,24)
$exportButton.Location = New-Object System.Drawing.Point(410,178)
$exportButton.Font = New-Object System.Drawing.Font("Segoe UI",8)

$importButton = New-Object System.Windows.Forms.Button
$importButton.Text = "Config import"
$importButton.Size = New-Object System.Drawing.Size(120,24)
$importButton.Location = New-Object System.Drawing.Point(540,178)
$importButton.Font = New-Object System.Drawing.Font("Segoe UI",8)

$form.Controls.Add($title)
$form.Controls.Add($subtitle)
$form.Controls.Add($authorLabel)
$form.Controls.Add($folderLabel)
$form.Controls.Add($lastCheckLabel)
$form.Controls.Add($searchLabel)
$form.Controls.Add($searchBox)
$form.Controls.Add($clearSearchButton)
$form.Controls.Add($countLabel)
$form.Controls.Add($changeFolderButton)
$form.Controls.Add($refreshButton)
$form.Controls.Add($statusButton)
$form.Controls.Add($autoCheckBox)
$form.Controls.Add($toolsLabel)
$form.Controls.Add($settingsLabel)
$form.Controls.Add($exportButton)
$form.Controls.Add($importButton)
$form.Controls.Add($sshUserLabel)
$form.Controls.Add($sshUserBox)
$form.Controls.Add($saveSshUserButton)
$form.Controls.Add($historyButton)

$panel = New-Object System.Windows.Forms.FlowLayoutPanel
$panel.Location = New-Object System.Drawing.Point(32,286)
$panel.Size = New-Object System.Drawing.Size(880,392)
$panel.AutoScroll = $true
$panel.FlowDirection = "TopDown"
$panel.WrapContents = $false
$panel.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

$form.Controls.Add($panel)

function New-GroupHeader {
    param([string]$text)

    $header = New-Object System.Windows.Forms.Label
    $header.Text = $text
    $header.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    $header.ForeColor = [System.Drawing.Color]::FromArgb(70,70,70)
    $header.AutoSize = $false
    $header.Size = New-Object System.Drawing.Size(840,26)
    $header.Margin = New-Object System.Windows.Forms.Padding(0,10,0,6)
    $header.TextAlign = "MiddleLeft"

    return $header
}

function Update-OneStatus {
    param(
        [string]$fileName,
        [string]$address,
        [System.Windows.Forms.Panel]$dot,
        [System.Windows.Forms.Label]$label
    )

    Set-StatusView -dot $dot -label $label -status "CHECKING"
    [System.Windows.Forms.Application]::DoEvents()

    $status = Test-RdpHost $address
    if (-not [string]::IsNullOrWhiteSpace($fileName)) {
        $script:statusCache[$fileName] = $status
    }
    Set-StatusView -dot $dot -label $label -status $status
    $lastCheckLabel.Text = "Letzter Check: " + (Get-Date).ToString("HH:mm:ss")
}

function Create-ContextMenu {
    param(
        [string]$fileName,
        [string]$address,
        [string]$username,
        [System.Windows.Forms.Panel]$statusDot,
        [System.Windows.Forms.Label]$statusLabel
    )

    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    $openItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openItem.Text = "▶ Start"
    $openItem.Tag = $fileName
    $openItem.Add_Click({
        Start-RdpFile -fileName $this.Tag
    })

    $statusItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $statusItem.Text = "Status checken"
    $statusItem.Tag = @{
        FileName = $fileName
        Address = $address
        Dot = $statusDot
        Label = $statusLabel
    }
    $statusItem.Add_Click({
        Update-OneStatus -fileName $this.Tag.FileName -address $this.Tag.Address -dot $this.Tag.Dot -label $this.Tag.Label
    })

    $editItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $editItem.Text = "RDP-Datei bearbeiten"
    $editItem.Tag = $fileName
    $editItem.Add_Click({
        Open-RdpInEditor -fileName $this.Tag
    })

    $folderItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $folderItem.Text = "Im Ordner anzeigen"
    $folderItem.Tag = $fileName
    $folderItem.Add_Click({
        Open-RdpInExplorer -fileName $this.Tag
    })

    $copyTargetItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyTargetItem.Text = "Ziel kopieren"
    $copyTargetItem.Tag = $address
    $copyTargetItem.Add_Click({
        Set-ClipboardText -text ([string]$this.Tag) -message "Kopiert: Ziel"
    })

    $copyRdpItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyRdpItem.Text = "RDP-Adresse kopieren"
    $copyRdpItem.Tag = $address
    $copyRdpItem.Add_Click({
        $target = [string]$this.Tag
        Set-ClipboardText -text ("mstsc /v:$target") -message "Kopiert: RDP-Adresse"
    })

    $copySshItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copySshItem.Text = "SSH-Befehl kopieren"
    $copySshItem.Tag = @{
        Address = $address
        Username = $username
    }
    $copySshItem.Add_Click({
        $addr = [string]$this.Tag.Address
        $user = [string]$this.Tag.Username

        if ([string]::IsNullOrWhiteSpace($user)) {
            $user = $script:defaultSshUser
        }

        if ([string]::IsNullOrWhiteSpace($user)) {
            Set-ClipboardText -text ("ssh " + $addr) -message "Kopiert: SSH-Befehl"
        } else {
            Set-ClipboardText -text ("ssh " + $user + "@" + $addr) -message "Kopiert: SSH-Befehl"
        }
    })

    $favItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $favItem.Text = "Favorit umschalten"
    $favItem.Tag = $fileName
    $favItem.Add_Click({
        Toggle-Favorite -fileName $this.Tag
    })

    $menu.Items.Add($openItem) | Out-Null
    $menu.Items.Add($statusItem) | Out-Null
    $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
    $menu.Items.Add($editItem) | Out-Null
    $menu.Items.Add($folderItem) | Out-Null
    $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
    $menu.Items.Add($copyTargetItem) | Out-Null
    $menu.Items.Add($copyRdpItem) | Out-Null
    $menu.Items.Add($copySshItem) | Out-Null
    $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
    $menu.Items.Add($favItem) | Out-Null

    return $menu
}

function New-RdpCard {
    param(
        [string]$displayName,
        [string]$fileName,
        [string]$address,
        [string]$username,
        [bool]$isFavorite
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size(840,86)
    $card.Margin = New-Object System.Windows.Forms.Padding(0,0,0,14)
    if ($script:lastUsedFile -eq $fileName) {
        $card.BackColor = [System.Drawing.Color]::FromArgb(245,250,255)
    } elseif ($isFavorite) {
        $card.BackColor = [System.Drawing.Color]::FromArgb(255,252,245)
    } else {
        $card.BackColor = [System.Drawing.Color]::White
    }
    $card.BorderStyle = "FixedSingle"
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Tag = $fileName

    $icon = New-Object System.Windows.Forms.Label
    $icon.Text = "RDP"
    $icon.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    $icon.ForeColor = [System.Drawing.Color]::FromArgb(40,90,160)
    $icon.AutoSize = $true
    $icon.Location = New-Object System.Drawing.Point(24,30)
    $icon.Tag = $fileName
    $icon.Cursor = [System.Windows.Forms.Cursors]::Hand
    $icon.BackColor = $card.BackColor

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = $displayName
    $nameLabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
    $nameLabel.AutoSize = $true
    $nameLabel.Location = New-Object System.Drawing.Point(86,15)
    $nameLabel.Tag = $fileName
    $nameLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $nameLabel.BackColor = $card.BackColor

    $fileLabel = New-Object System.Windows.Forms.Label
    $fileLabel.Text = $fileName
    $fileLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
    $fileLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
    $fileLabel.AutoSize = $true
    $fileLabel.Location = New-Object System.Drawing.Point(88,40)
    $fileLabel.Tag = $fileName
    $fileLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $fileLabel.BackColor = $card.BackColor

    $addressLabel = New-Object System.Windows.Forms.Label
    $addressLabel.Text = "Ziel: $address"
    $addressLabel.Font = New-Object System.Drawing.Font("Segoe UI",8)
    $addressLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
    $addressLabel.AutoSize = $true
    $addressLabel.Location = New-Object System.Drawing.Point(88,58)
    $addressLabel.Tag = $fileName
    $addressLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $addressLabel.BackColor = $card.BackColor

    $statusDot = New-Object System.Windows.Forms.Panel
    $statusDot.Size = New-Object System.Drawing.Size(11,11)
    $statusDot.Location = New-Object System.Drawing.Point(555,24)
    $statusDot.BackColor = [System.Drawing.Color]::FromArgb(130,130,130)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $statusLabel.AutoSize = $true
    $statusLabel.Location = New-Object System.Drawing.Point(575,20)
    $statusLabel.BackColor = $card.BackColor
    if ($script:statusCache.ContainsKey($fileName)) {
        Set-StatusView -dot $statusDot -label $statusLabel -status $script:statusCache[$fileName]
    } else {
        Set-StatusView -dot $statusDot -label $statusLabel -status "NOT_CHECKED"
    }

    $launchLabel = New-Object System.Windows.Forms.Label
    $launchLabel.Text = "▶ Start"
    $launchLabel.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $launchLabel.ForeColor = [System.Drawing.Color]::FromArgb(40,90,160)
    $launchLabel.AutoSize = $true
    $launchLabel.Location = New-Object System.Drawing.Point(575,48)
    $launchLabel.Tag = $fileName
    $launchLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $launchLabel.BackColor = $card.BackColor

    $editLabel = New-Object System.Windows.Forms.Label
    $editLabel.Text = "Edit"
    $editLabel.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
    $editLabel.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
    $editLabel.AutoSize = $true
    $editLabel.Location = New-Object System.Drawing.Point(645,49)
    $editLabel.Tag = $fileName
    $editLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $editLabel.BackColor = $card.BackColor

    $copyLabel = New-Object System.Windows.Forms.Label
    $copyLabel.Text = "Copy"
    $copyLabel.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
    $copyLabel.ForeColor = [System.Drawing.Color]::FromArgb(80,80,80)
    $copyLabel.AutoSize = $true
    $copyLabel.Location = New-Object System.Drawing.Point(690,49)
    $copyLabel.Tag = @{
        Address = $address
        Username = $username
    }
    $copyLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $copyLabel.BackColor = $card.BackColor

    $favoriteLabel = New-Object System.Windows.Forms.Label
    if ($isFavorite) {
        $favoriteLabel.Text = "Favorit"
        $favoriteLabel.ForeColor = [System.Drawing.Color]::FromArgb(170,95,0)
    } else {
        $favoriteLabel.Text = "Markieren"
        $favoriteLabel.ForeColor = [System.Drawing.Color]::FromArgb(145,145,145)
    }
    $favoriteLabel.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
    $favoriteLabel.AutoSize = $true
    $favoriteLabel.Location = New-Object System.Drawing.Point(760,58)
    $favoriteLabel.Tag = $fileName
    $favoriteLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $favoriteLabel.BackColor = $card.BackColor

    $contextMenu = Create-ContextMenu -fileName $fileName -address $address -username $username -statusDot $statusDot -statusLabel $statusLabel

    $card.ContextMenuStrip = $contextMenu
    $icon.ContextMenuStrip = $contextMenu
    $nameLabel.ContextMenuStrip = $contextMenu
    $fileLabel.ContextMenuStrip = $contextMenu
    $addressLabel.ContextMenuStrip = $contextMenu
    $statusDot.ContextMenuStrip = $contextMenu
    $statusLabel.ContextMenuStrip = $contextMenu
    $launchLabel.ContextMenuStrip = $contextMenu
    $favoriteLabel.ContextMenuStrip = $contextMenu
    $editLabel.ContextMenuStrip = $contextMenu
    $copyLabel.ContextMenuStrip = $contextMenu

    $openAction = {
        $file = $this.Tag
        Start-RdpFile -fileName $file
    }

    $favAction = {
        Toggle-Favorite -fileName $this.Tag
    }

    $editAction = {
        Open-RdpInEditor -fileName $this.Tag
    }

    $copyAction = {
        $addr = [string]$this.Tag.Address
        $user = [string]$this.Tag.Username
        if ([string]::IsNullOrWhiteSpace($user)) {
            $user = $script:defaultSshUser
        }

        if ([string]::IsNullOrWhiteSpace($user)) {
            Set-ClipboardText -text ("ssh " + $addr) -message "Kopiert: SSH-Befehl"
        } else {
            Set-ClipboardText -text ("ssh " + $user + "@" + $addr) -message "Kopiert: SSH-Befehl"
        }
    }

    $card.Add_Click($openAction)
    $icon.Add_Click($openAction)
    $nameLabel.Add_Click($openAction)
    $fileLabel.Add_Click($openAction)
    $addressLabel.Add_Click($openAction)
    $launchLabel.Add_Click($openAction)
    $favoriteLabel.Add_Click($favAction)
    $editLabel.Add_Click($editAction)
    $copyLabel.Add_Click($copyAction)

    $card.Controls.Add($icon)
    $card.Controls.Add($nameLabel)
    $card.Controls.Add($fileLabel)
    $card.Controls.Add($addressLabel)
    $card.Controls.Add($statusDot)
    $card.Controls.Add($statusLabel)
    $card.Controls.Add($launchLabel)
    $card.Controls.Add($favoriteLabel)
    $card.Controls.Add($editLabel)
    $card.Controls.Add($copyLabel)

    return @{
        Card = $card
        FileName = $fileName
        DisplayName = $displayName
        Address = $address
        Username = $username
        StatusDot = $statusDot
        StatusLabel = $statusLabel
    }
}

function Render-FilteredCards {
    $panel.Controls.Clear()
    $script:rdpItems = @()

    $filter = $searchBox.Text.Trim().ToLower()
    $filtered = @()

    foreach ($entry in $script:rdpData) {
        $haystack = ($entry.DisplayName + " " + $entry.FileName + " " + $entry.Address + " " + $entry.Username + " " + $entry.Group).ToLower()
        if ([string]::IsNullOrWhiteSpace($filter) -or $haystack.Contains($filter)) {
            $filtered += $entry
        }
    }

    $groupOrder = @("Zuletzt genutzt", "Favoriten", "Rechner", "Linux", "Server", "Testsysteme", "Weitere Verbindungen")
    $ordered = @()

    foreach ($group in $groupOrder) {
        $items = @($filtered | Where-Object { $_.Group -eq $group } | Sort-Object DisplayName)
        if ($items.Count -gt 0) {
            $ordered += @{
                Group = $group
                Items = $items
            }
        }
    }

    if ($filtered.Count -eq 0) {
        $emptyLabel = New-Object System.Windows.Forms.Label
        $emptyLabel.Text = "Keine passenden RDP-Dateien gefunden."
        $emptyLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
        $emptyLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
        $emptyLabel.AutoSize = $true
        $emptyLabel.Margin = New-Object System.Windows.Forms.Padding(10,20,0,0)
        $panel.Controls.Add($emptyLabel)
        $countLabel.Text = "0 Treffer"
        return
    }

    foreach ($groupBlock in $ordered) {
        $panel.Controls.Add((New-GroupHeader -text $groupBlock.Group))
        foreach ($entry in $groupBlock.Items) {
            $item = New-RdpCard -displayName $entry.DisplayName -fileName $entry.FileName -address $entry.Address -username $entry.Username -isFavorite $entry.IsFavorite
            $script:rdpItems += $item
            $panel.Controls.Add($item.Card)
        }
    }

    $favCount = @($filtered | Where-Object { $_.IsFavorite }).Count

    if ($filtered.Count -eq 1) {
        if ($favCount -gt 0) {
            $countLabel.Text = "1 Treffer, 1 Favorit"
        } else {
            $countLabel.Text = "1 Treffer"
        }
    } else {
        if ($favCount -gt 0) {
            $countLabel.Text = "$($filtered.Count) Treffer, $favCount Favorit(en)"
        } else {
            $countLabel.Text = "$($filtered.Count) Treffer"
        }
    }
}

function Load-RdpCards {
    $script:rdpData = @()
    $script:rdpItems = @()
    $panel.Controls.Clear()
    Update-FolderLabel -Path $basePath
    $lastCheckLabel.Text = "Letzter Check: -"

    if (!(Test-Path $basePath)) {
        $script:basePath = Find-RdpFolder -configuredFolder $script:basePath
        $basePath = $script:basePath
        Update-FolderLabel -Path $basePath
    }

    if (!(Test-Path $basePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "RDP-Ordner nicht gefunden:`n$basePath",
            "RDP Launcher",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $rdpFiles = Get-ChildItem -Path $basePath -Filter *.rdp | Sort-Object BaseName

    if ($rdpFiles.Count -eq 0) {
        $emptyLabel = New-Object System.Windows.Forms.Label
        $emptyLabel.Text = "Keine .rdp-Dateien im ausgewaehlten Ordner gefunden."
        $emptyLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
        $emptyLabel.ForeColor = [System.Drawing.Color]::FromArgb(120,120,120)
        $emptyLabel.AutoSize = $true
        $emptyLabel.Margin = New-Object System.Windows.Forms.Padding(10,20,0,0)
        $panel.Controls.Add($emptyLabel)
        $countLabel.Text = "0 Verbindungen"
        return
    }

    foreach ($rdp in $rdpFiles) {
        $displayName = Format-Name $rdp.BaseName
        $address = Get-RdpAddress $rdp.FullName
        $username = Get-RdpUsername $rdp.FullName

        if ([string]::IsNullOrWhiteSpace($address)) {
            $address = "unbekannt"
        }

        $isFav = ($script:favorites -contains $rdp.Name)
        $group = Get-RdpGroup -fileName $rdp.Name -displayName $displayName -isFavorite $isFav

        $script:rdpData += @{
            DisplayName = $displayName
            FileName = $rdp.Name
            Address = $address
            Username = $username
            IsFavorite = $isFav
            Group = $group
        }
    }

    Render-FilteredCards
}

function Update-OnlyStatus {
    if ($script:rdpItems.Count -eq 0) {
        return
    }

    foreach ($item in $script:rdpItems) {
        Set-StatusView -dot $item.StatusDot -label $item.StatusLabel -status "CHECKING"
    }

    [System.Windows.Forms.Application]::DoEvents()

    foreach ($item in $script:rdpItems) {
        $status = Test-RdpHost $item.Address
        $script:statusCache[$item.FileName] = $status
        Set-StatusView -dot $item.StatusDot -label $item.StatusLabel -status $status
    }

    $lastCheckLabel.Text = "Letzter Check: " + (Get-Date).ToString("HH:mm:ss")
}

$changeFolderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Ordner mit .rdp-Dateien auswaehlen"
    $dialog.SelectedPath = $script:basePath

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:basePath = $dialog.SelectedPath

        Save-Config `
            -folder $script:basePath `
            -favoriteFiles $script:favorites `
            -defaultSshUser $script:defaultSshUser `
            -lastUsedFile $script:lastUsedFile `
            -history $script:history

        Update-FolderLabel -Path $script:basePath
        $searchBox.Text = ""
        Load-RdpCards
        $lastCheckLabel.Text = "Ordner gewechselt"
    }
})

$refreshButton.Add_Click({
    Load-RdpCards
})

$statusButton.Add_Click({
    Update-OnlyStatus
})

$searchBox.Add_TextChanged({
    Render-FilteredCards
})

$clearSearchButton.Add_Click({
    $searchBox.Text = ""
    Render-FilteredCards
})

$exportButton.Add_Click({
    Export-ConfigFile
})

$importButton.Add_Click({
    Import-ConfigFile
})

$saveSshUserButton.Add_Click({
    $script:defaultSshUser = $sshUserBox.Text.Trim()

    Save-Config `
        -folder $script:basePath `
        -favoriteFiles $script:favorites `
        -defaultSshUser $script:defaultSshUser `
        -lastUsedFile $script:lastUsedFile `
        -history $script:history

    $lastCheckLabel.Text = "SSH-User gespeichert"
})

$historyButton.Add_Click({
    Show-History
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000

$timer.Add_Tick({
    if ($autoCheckBox.Checked) {
        Update-OnlyStatus
    }
})

$autoCheckBox.Add_CheckedChanged({
    $timer.Enabled = $autoCheckBox.Checked
    if ($autoCheckBox.Checked) {
        Update-OnlyStatus
    }
})

Load-RdpCards

[void]$form.ShowDialog()
