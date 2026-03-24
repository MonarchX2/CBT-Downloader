# ==============================
# Configuration (Courses)
# ==============================

$Courses = @{
    "1" = @{
        Name = "CD0078 - Pumps and Pumping Operations"
        URL  = "https://drive.google.com/uc?export=download&id=1xPWJnqwpVftOmzcTfrowxBJ8w5PI1qFZ"
        FolderName = "CD0078 - Pumps and Pumping Operations"
    }
    "2" = @{
        Name = "CD0017 - Steering Gear"
        URL  = "https://drive.google.com/uc?export=download&id=13_way_Eg7CrUC1ZA--luTYqAygVH7zJt"
        FolderName = "CD0017 - Steering Gear"
    }
    "3" = @{
        Name = "CD0118 - Steering Gear (RAM Type)"
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
# Check Existing Folder
# ==============================

if (Test-Path $ExtractPath) {
    Write-Host "Existing folder found. Skipping download..."
} else {
    Write-Host "Downloading..."
    Invoke-WebRequest -Uri $Selected.URL -OutFile $ZipPath -UseBasicParsing

    Write-Host "Extracting..."
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
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
    Write-Host "viewer.exe not found. Opening folder..."
    Start-Process $ExtractPath
}

Write-Host "Done."
