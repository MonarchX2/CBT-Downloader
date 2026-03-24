# ==============================
# Configuration
# ==============================

# Google Drive File ID
$FileID = "1xPWJnqwpVftOmzcTfrowxBJ8w5PI1qFZ"

# Universal direct download link
$DownloadURL = "https://drive.google.com/uc?export=download&id=$FileID"

# Get current user's default Downloads folder
$DownloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"

# File paths
$ZipPath     = Join-Path $DownloadsPath "viewer_package.zip"
$ExtractPath = Join-Path $DownloadsPath "viewer_extracted"

# ==============================
# Download ZIP
# ==============================

Write-Host "Downloading file to Downloads folder..."
Invoke-WebRequest -Uri $DownloadURL -OutFile $ZipPath -UseBasicParsing

# ==============================
# Extract ZIP
# ==============================

Write-Host "Extracting ZIP..."

If (Test-Path $ExtractPath) {
    Remove-Item $ExtractPath -Recurse -Force
}

Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force

# ==============================
# Locate and Launch viewer.exe
# ==============================

Write-Host "Searching for viewer.exe..."

$ExePath = Get-ChildItem -Path $ExtractPath -Recurse -Filter "viewer.exe" | Select-Object -First 1

if ($ExePath) {
    Write-Host "Launching viewer.exe..."
    Start-Process -FilePath $ExePath.FullName
} else {
    Write-Host "viewer.exe not found."
}

Write-Host "Process complete."
