if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Restarting as administrator..." -ForegroundColor Red
    $arg = if ([string]::IsNullOrEmpty($PSCommandPath)) {
        "-NoProfile -ExecutionPolicy Bypass -Command `"irm https://go.bibica.net/edge_multi | iex`""
    } else {
        "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    }
    Start-Process powershell.exe $arg -Verb RunAs
    exit
}

Clear-Host
Write-Host " Microsoft Edge Browser Multi Installer " -BackgroundColor DarkGreen

# Kill processes
@("msedge", "MicrosoftEdgeUpdate", "edgeupdate", "edgeupdatem", "MicrosoftEdgeSetup") | ForEach-Object {
    Get-Process -Name $_ -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# Create main temp directory
$tempDir = "$env:USERPROFILE\Downloads\microsoft-edge-multi-installer-debloater"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Define download URLs and channels
$edgeChannels = @{
    "1" = "msedge-stable-win-x64"
    "2" = "msedge-beta-win-x64"
    "3" = "msedge-dev-win-x64"
    "4" = "msedge-canary-win-x64"
}
$channelNames = @{
    "1" = "Stable"
    "2" = "Beta"
    "3" = "Developer"
    "4" = "Canary"
}
$channelPaths = @{
    "1" = "C:\Program Files (x86)\Microsoft\Edge"
    "2" = "C:\Program Files (x86)\Microsoft\Edge Beta"
    "3" = "C:\Program Files (x86)\Microsoft\Edge Dev"
    "4" = "C:\Program Files (x86)\Microsoft\Edge Canary"
}
$shortcutNames = @{
    "1" = "Microsoft Edge"
    "2" = "Microsoft Edge Beta"
    "3" = "Microsoft Edge Dev"
    "4" = "Microsoft Edge Canary"
}
$userDataPaths = @{
    "1" = "Microsoft\Edge"
    "2" = "Microsoft\Edge Beta"
    "3" = "Microsoft\Edge Dev"
    "4" = "Microsoft\Edge Canary"
}
$iconUrls = @{
    "1" = "https://github.com/bibicadotnet/microsoft-edge-multi-installer-debloater/raw/refs/heads/main/icon/Edge.ico"
    "2" = "https://github.com/bibicadotnet/microsoft-edge-multi-installer-debloater/raw/refs/heads/main/icon/Edge%20Beta.ico"
    "3" = "https://github.com/bibicadotnet/microsoft-edge-multi-installer-debloater/raw/refs/heads/main/icon/Edge%20Dev.ico"
    "4" = "https://github.com/bibicadotnet/microsoft-edge-multi-installer-debloater/raw/refs/heads/main/icon/Edge%20Canary.ico"
}
$checkVersionUrl = "https://msedge.api.cdp.microsoft.com/api/v2/contents/Browser/namespaces/Default/names/{0}/versions/latest?action=select"
$getDownloadLinkUrl = "https://msedge.api.cdp.microsoft.com/api/v1.1/internal/contents/Browser/namespaces/Default/names/{0}/versions/{1}/files?action=GenerateDownloadInfo"
$headers = @{
    "User-Agent" = "Microsoft Edge Update/1.3.183.29;winhttp"
}

# Prompt for channel selection
#Write-Host "Fetching latest versions for all channels..."

# Fetch all Edge release data once
$allReleases = irm "https://edgeupdates.microsoft.com/api/products"

# Lấy phiên bản cho từng channel
$channelVersions = @{
    "1" = ($allReleases | Where-Object { $_.Product -eq "Stable"    }).Releases |
           Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
           Sort-Object PublishedTime -Descending |
           Select-Object -First 1 -ExpandProperty ProductVersion

    "2" = ($allReleases | Where-Object { $_.Product -eq "Beta"      }).Releases |
           Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
           Sort-Object PublishedTime -Descending |
           Select-Object -First 1 -ExpandProperty ProductVersion

    "3" = ($allReleases | Where-Object { $_.Product -eq "Dev"       }).Releases |
           Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
           Sort-Object PublishedTime -Descending |
           Select-Object -First 1 -ExpandProperty ProductVersion

    "4" = ($allReleases | Where-Object { $_.Product -eq "Canary"    }).Releases |
           Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } |
           Sort-Object PublishedTime -Descending |
           Select-Object -First 1 -ExpandProperty ProductVersion
}

Write-Host "`nSelect Edge channel to install:"
Write-Host "1. Stable ($($channelVersions['1']))"
Write-Host "2. Beta ($($channelVersions['2']))"
Write-Host "3. Developer ($($channelVersions['3']))"
Write-Host "4. Canary ($($channelVersions['4']))"
$selection = Read-Host "`nEnter number (1-4) or press Enter for Stable"

if (![string]::IsNullOrWhiteSpace($selection) -and $edgeChannels.ContainsKey($selection)) {
    $appId = $edgeChannels[$selection]
    $channelName = $channelNames[$selection]
    $channelPath = $channelPaths[$selection]
    $shortcutName = $shortcutNames[$selection]
    $userDataPath = $userDataPaths[$selection]
    $iconUrl = $iconUrls[$selection]
} else {
    $appId = $edgeChannels["1"]
    $channelName = $channelNames["1"]
    $channelPath = $channelPaths["1"]
    $shortcutName = $shortcutNames["1"]
    $userDataPath = $userDataPaths["1"]
    $iconUrl = $iconUrls["1"]
}

# Fetch latest version
$versionBody = @{
    "targetingAttributes" = @{
        "IsInternalUser" = $true
        "Updater" = "MicrosoftEdgeUpdate"
        "UpdaterVersion" = "1.3.183.29"
    }
} | ConvertTo-Json -Depth 3

$versionUrl = [string]::Format($checkVersionUrl, $appId)
$versionResponse = Invoke-RestMethod -Method Post -Uri $versionUrl -Headers $headers -Body $versionBody -ContentType "application/json"
$version = $versionResponse.ContentId.Version
Write-Host "[INFO] Latest version for ${channelName}: $version"

# Get download links
$downloadUrl = [string]::Format($getDownloadLinkUrl, $appId, $version)
$downloadResponse = Invoke-RestMethod -Method Post -Uri $downloadUrl -Headers $headers

# Find the .exe installer
$installer = $downloadResponse | Where-Object {
    $_.FileId -match "\.exe$"
} | Sort-Object SizeInBytes -Descending | Select-Object -First 1

if ($null -eq $installer) {
    Write-Host "[ERROR] No .exe installer found."
    exit 1
}

$originalFileName = $installer.FileId
$downloadUrl = $installer.Url
$size = [Math]::Round($installer.SizeInBytes / 1MB, 2)
$sha256 = $installer.Hashes.Sha256

Write-Host "[INFO] Installer found:"
Write-Host "  File: $originalFileName"
Write-Host "  Size: ${size} MB"
Write-Host "  SHA256: $sha256"

# Download with proper filename to temp directory
$outputPath = Join-Path $tempDir $originalFileName
Write-Host "[INFO] Downloading $originalFileName to temp directory..."
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($downloadUrl, $outputPath)
Write-Host "[INFO] Downloaded to: $outputPath"

# Download 7zr.exe to temp directory
$sevenZPath = Join-Path $tempDir "7zr.exe"
if (-not (Test-Path $sevenZPath)) {
    Write-Host "[INFO] Downloading 7zr.exe to temp directory..."
    $wc.DownloadFile("https://www.7-zip.org/a/7zr.exe", $sevenZPath)
}

# Create subdirectories in temp for organization
$tempExtractFolder = Join-Path $tempDir "temp_extract"
$finalExtractFolder = Join-Path $tempDir "MSEDGE_EXTRACTED"
$tempAppFolder = Join-Path $tempDir "temp_application"

# Clean up existing subdirectories
@($tempExtractFolder, $finalExtractFolder, $tempAppFolder) | ForEach-Object {
    if (Test-Path $_) {
        Remove-Item $_ -Recurse -Force
    }
}

New-Item -ItemType Directory -Path $tempExtractFolder | Out-Null

Write-Host "[INFO] Extracting $originalFileName..."
& $sevenZPath x $outputPath "-o$tempExtractFolder" -y | Out-Null

# Find and extract MSEDGE.7z
$msedge7zPath = Get-ChildItem -Path $tempExtractFolder -Name "MSEDGE.7z" -Recurse | Select-Object -First 1
if ($msedge7zPath) {
    $fullMsedge7zPath = Join-Path $tempExtractFolder $msedge7zPath
    Write-Host "[INFO] Found MSEDGE.7z, extracting..."
    
    New-Item -ItemType Directory -Path $finalExtractFolder | Out-Null
    
    & $sevenZPath x $fullMsedge7zPath "-o$finalExtractFolder" -y | Out-Null
    
    # Find the Chrome-bin\version folder
    $chromeBinPath = Get-ChildItem -Path $finalExtractFolder -Directory -Recurse | Where-Object { 
        $_.Parent.Name -eq "Chrome-bin" -and $_.Name -match "^\d+\.\d+\.\d+\.\d+$"
    } | Select-Object -First 1
    
    if ($chromeBinPath) {
        # Create temporary Application folder structure
        New-Item -ItemType Directory -Path $tempAppFolder -Force | Out-Null
        
        # Copy version folder to temp Application
        $versionInApp = Join-Path $tempAppFolder $chromeBinPath.Name
        Copy-Item $chromeBinPath.FullName $versionInApp -Recurse -Force
        
        # Create/copy main Application files
        $mainFiles = @(
            "msedge.exe",
            "msedge.VisualElementsManifest.xml", 
            "msedge_proxy.exe",
            "pwahelper.exe",
            "initial_preferences",
            "delegatedWebFeatures.sccd"
        )
        
        foreach ($file in $mainFiles) {
            $sourceFile = Get-ChildItem -Path $finalExtractFolder -Name $file -Recurse | Select-Object -First 1
            if ($sourceFile) {
                $sourcePath = Join-Path $finalExtractFolder $sourceFile
                $destPath = Join-Path $tempAppFolder $file
                Copy-Item $sourcePath $destPath -Force -ErrorAction SilentlyContinue
            }
        }
        
        # Create SetupMetrics folder
        $setupMetricsFolder = Join-Path $tempAppFolder "SetupMetrics"
        New-Item -ItemType Directory -Path $setupMetricsFolder -Force | Out-Null
        
        # Copy main msedge.exe to Application root if not already there
        $mainMsedgeExe = Join-Path $tempAppFolder "msedge.exe"
        $versionMsedgeExe = Join-Path $versionInApp "msedge.exe"
        if ((Test-Path $versionMsedgeExe) -and (-not (Test-Path $mainMsedgeExe))) {
            Copy-Item $versionMsedgeExe $mainMsedgeExe -Force
        }
        
        # Download and copy icon file to version folder
        Write-Host "[INFO] Downloading icon for $channelName..."
        $iconFileName = "Edge_Icon.ico"
        $iconPath = Join-Path $versionInApp $iconFileName
        
        try {
            $wc.DownloadFile($iconUrl, $iconPath)
            Write-Host "[INFO] Icon downloaded to: $iconPath"
        } catch {
            Write-Host "[WARNING] Failed to download icon: $($_.Exception.Message)"
            $iconPath = $versionMsedgeExe  # Fallback to msedge.exe for icon
        }
        
        Write-Host "[INFO] Temporary Application folder ready: $tempAppFolder"
        
        # Create user data directory for this channel
        $userDataFullPath = "$env:LOCALAPPDATA\$userDataPath\User Data"
        Write-Host "[INFO] Creating user data directory: $userDataFullPath"
        if (-not (Test-Path $userDataFullPath)) {
            New-Item -ItemType Directory -Path $userDataFullPath -Force | Out-Null
        }
        
        # Now copy to system location
        $systemPath = $channelPath
        $systemAppPath = Join-Path $systemPath "Application"
        
        Write-Host "[INFO] Installing to: $systemPath"
        
        try {
            # Create system directory if not exists
            if (-not (Test-Path $systemPath)) {
                New-Item -ItemType Directory -Path $systemPath -Force | Out-Null
            }
            
            # Remove existing Application folder if exists
            if (Test-Path $systemAppPath) {
                Remove-Item $systemAppPath -Recurse -Force
            }
            
            # Copy Application folder to system location
            Copy-Item $tempAppFolder $systemAppPath -Recurse -Force
            
            Write-Host "[SUCCESS] Edge $channelName installed successfully!"
            Write-Host "[INFO] Installation path: $systemPath"
            Write-Host "[INFO] Executable: $(Join-Path $systemAppPath 'msedge.exe')"
            Write-Host "[INFO] Version: $version ($(Join-Path $systemAppPath $chromeBinPath.Name))"
            Write-Host "[INFO] Profile path: $userDataFullPath\Default"
            
            # Create shortcuts with improved method
            $edgeExePath = Join-Path $systemAppPath "msedge.exe"
            $iconPath = Join-Path (Join-Path $systemAppPath $chromeBinPath.Name) "Edge_Icon.ico"
            
            # Fallback to msedge.exe if icon not found
            if (-not (Test-Path $iconPath)) {
                $iconPath = $edgeExePath
            }
            
            if (Test-Path $edgeExePath) {
                Write-Host "[INFO] Creating shortcuts for $shortcutName..."
                
                # Remove old shortcuts from ALL locations first
                @(
                    [Environment]::GetFolderPath("Desktop"),
                    [Environment]::GetFolderPath("CommonDesktopDirectory"), 
                    [Environment]::GetFolderPath("Programs"),
                    [Environment]::GetFolderPath("CommonPrograms")
                ) | ForEach-Object {
                    if (Test-Path $_) {
                        Get-ChildItem "$_\$shortcutName.lnk" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                        # Also remove potential variations
                        Get-ChildItem "$_\*Edge*.lnk" -ErrorAction SilentlyContinue | Where-Object { 
                            $_.Name -like "*$channelName*" -or $_.Name -eq "$shortcutName.lnk"
                        } | Remove-Item -Force -ErrorAction SilentlyContinue
                    }
                }
                
                # Create shortcuts arguments
                $edgeArgs = "--user-data-dir=`"$userDataFullPath`" --disable-features=msCopilotMode2"
                
                # Create Start Menu folder path
                $startMenuPath = Join-Path ([Environment]::GetFolderPath("CommonPrograms")) "Microsoft"
                if (-not (Test-Path $startMenuPath)) {
                    New-Item -ItemType Directory -Path $startMenuPath -Force | Out-Null
                }
                
				# Create new shortcuts using improved method
                $shortcutPaths = @([Environment]::GetFolderPath("Desktop"), $startMenuPath)
                foreach ($shortcutPath in $shortcutPaths) {
                    $WshShell = New-Object -ComObject WScript.Shell
                    $temp = "$shortcutPath\temp_edge_shortcut.lnk"
                    $final = "$shortcutPath\$shortcutName.lnk"
                    
                    try {
                        $shortcut = $WshShell.CreateShortcut($temp)
                        $shortcut.TargetPath = $edgeExePath
                        $shortcut.Arguments = $edgeArgs
                        $shortcut.WorkingDirectory = $systemAppPath
                        $shortcut.Description = $shortcutName
                        $shortcut.IconLocation = "$iconPath,0"
                        $shortcut.Save()
                        
                        # Rename temp to final name
                        if (Test-Path $temp) {
                            if (Test-Path $final) {
                                Remove-Item $final -Force -ErrorAction SilentlyContinue
                            }
                            Rename-Item $temp $final -ErrorAction SilentlyContinue
                            
                            if (Test-Path $final) {
                                $location = if ($shortcutPath -like "*Desktop*") { "Desktop" } else { "Start Menu" }
                                Write-Host "[INFO] $location shortcut created: $final"
                            }
                        }
                        
                        # Remove temp file if still exists
                        Remove-Item $temp -ErrorAction SilentlyContinue
                        
                    } catch {
                        Write-Host "[WARNING] Failed to create shortcut at $shortcutPath`: $($_.Exception.Message)"
                        Remove-Item $temp -ErrorAction SilentlyContinue
                    } finally {
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
                    }
                }
                
                Write-Host "[INFO] Shortcut arguments: $edgeArgs"
                
            } else {
                Write-Host "[WARNING] msedge.exe not found, shortcuts not created"
            }
            
        } catch {
            Write-Host "[ERROR] Failed to copy to system location: $($_.Exception.Message)"
            Write-Host "[INFO] You may need to run PowerShell as Administrator"
            Write-Host "[INFO] Temporary files available at: $tempAppFolder"
            Write-Host "[INFO] Manually copy to: $systemPath"
        }
    } else {
        Write-Host "[WARNING] Could not find Chrome-bin version folder."
        Write-Host "[INFO] Files extracted to: $finalExtractFolder"
    }
} else {
    Write-Host "[ERROR] MSEDGE.7z not found in extracted files."
}

# Remove scheduled tasks
Get-ScheduledTask -TaskName "MicrosoftEdgeUpdate*" -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue

# Remove EdgeUpdate
@("msedge", "MicrosoftEdgeUpdate", "edgeupdate", "edgeupdatem", "MicrosoftEdgeSetup") | ForEach-Object {
    Get-Process -Name $_ -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
Remove-Item "${env:ProgramFiles(x86)}\Microsoft\EdgeUpdate" -Recurse -Force -ErrorAction SilentlyContinue

# Apply registry tweaks
$regUrl="https://raw.githubusercontent.com/bibicadotnet/microsoft-edge-debloater/refs/heads/main/vi.edge.reg"
$regFile="$tempDir\debloat.reg"
$wc=New-Object Net.WebClient
try{$wc.DownloadFile($regUrl,$regFile)}finally{$wc.Dispose()}
Start-Process regedit "/s `"$regFile`"" -Wait -NoNewWindow

# Clean up
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

# === Installation Summary ===
Write-Host ""
Write-Host "=== Installation Summary ==="
Write-Host "Channel: $channelName"
Write-Host "Version: $version"
Write-Host "Installation Path: $systemPath"
Write-Host "Profile Path: $userDataFullPath\Default"
Write-Host "Shortcut Arguments: --user-data-dir=`"$userDataFullPath`" --disable-features=msCopilotMode2"

$wc.Dispose()

Write-Host "`nNOTICE: To update Microsoft Edge when needed, please:" -ForegroundColor Cyan -BackgroundColor DarkGreen
Write-Host "1. Open PowerShell with Administrator privileges" -ForegroundColor White
Write-Host "2. Run the following command: irm https://go.bibica.net/edge_multi | iex" -ForegroundColor Yellow
Write-Host "3. Wait for the installation process to complete" -ForegroundColor White
Write-Host
