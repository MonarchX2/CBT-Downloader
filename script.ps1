# ==============================
# Configuration (Courses)
# ==============================

$Courses = @{
    "1" = @{
        Name = "CD0078 - Pumps and Pumping Operations"
        Type = "file"
        URL  = "https://drive.google.com/uc?export=download&id=1xPWJnqwpVftOmzcTfrowxBJ8w5PI1qFZ"
        FolderName = "CD0078 - Pumps and Pumping Operations"
    }
    "2" = @{
        Name = "CD0017 - Steering Gear"
        Type = "folder"
        URL  = "https://drive.google.com/uc?export=download&id=13_way_Eg7CrUC1ZA--luTYqAygVH7zJt"
        FolderName = "CD0017 - Steering Gear"
    }
    "3" = @{
        Name = "CD0118 - Steering Gear (RAM Type)"
        Type = "folder"
        URL  = "https://drive.google.com/uc?export=download&id=1_QxJ9Dpex75CSBZcD9DZYwzQT91kRr_v"
        FolderName = "CD0118 - Steering Gear (RAM Type)"
    }
}

$DownloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

# ==============================
# Menu
# ==============================

Write-Host "Select a course:"
foreach ($key in $Courses.Keys) {
    Write-Host "$key. $($Courses[$key].Name)"
}

$choice = Read-Host "Enter number"

if (-not $Courses.ContainsKey($choice)) {
    Write-Host "Invalid selection."
    exit
}

$Selected = $Courses[$choice]
$ExtractPath = Join-Path $DownloadsPath $Selected.FolderName
$ZipPath     = Join-Path $DownloadsPath ($Selected.FolderName + ".zip")

# ==============================
# Handle Existing Folder or Folder Download
# ==============================

if (Test-Path $ExtractPath) {
    Write-Host "Existing folder found. Skipping download..."
} else {

    if ($Selected.Type -eq "file") {
        # Direct download ZIP
        Write-Host "Downloading file..."
        Invoke-WebRequest -Uri $Selected.URL -OutFile $ZipPath -UseBasicParsing

        Write-Host "Extracting..."
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    }
    else {
        # Folder case: Google Drive folder
        Write-Host "This is a manual-download file. You need to download it manually."
        Write-Host "Please save it to your Downloads folder."
        Write-Host "The browser will open in 5 seconds..."
        Start-Sleep -Seconds 5

        Start-Process $Selected.URL

        # Bring PowerShell window to front
        $psWindow = Get-Process -Id $PID
        Add-Type '[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);' -Name WinAPI -Namespace User32
        [User32.WinAPI]::SetForegroundWindow($psWindow.MainWindowHandle) | Out-Null

        # Wait for the ZIP to appear
        Write-Host "Waiting for the downloaded ZIP file..."
        while (-not (Test-Path $ZipPath)) {
            Start-Sleep -Seconds 2
        }

        Write-Host "ZIP file detected. Extracting..."
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    }
}

# ==============================
# Check folder vs ZIP size
# ==============================
function Get-FolderSize($folderPath) {
    $size = (Get-ChildItem -Path $folderPath -Recurse -ErrorAction SilentlyContinue | 
             Measure-Object -Property Length -Sum).Sum
    return $size
}

$FolderSize = Get-FolderSize $ExtractPath
$ZipSize    = (Get-Item $ZipPath -ErrorAction SilentlyContinue).Length

if ($ZipSize -and $FolderSize) {
    # Allow 5% difference due to compression overhead
    $difference = [math]::Abs($FolderSize - $ZipSize) / $ZipSize
    if ($difference -lt 0.05) {
        Write-Host "Folder size is roughly equal to ZIP size. Deleting ZIP..."
        Remove-Item $ZipPath -Force
    } else {
        Write-Host "Folder size differs significantly from ZIP. Keeping ZIP."
    }
}

# ==============================
# Launch viewer.exe
# ==============================

Write-Host "Searching for viewer.exe..."

$ExePath = Get-ChildItem -Path $ExtractPath -Recurse -Filter "viewer.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

if ($ExePath) {
    Write-Host "Launching viewer.exe..."
    Start-Process -FilePath $ExePath.FullName
} else {
    Write-Host "viewer.exe not found. Opening folder instead..."
    Start-Process $ExtractPath
}

Write-Host "Done."
