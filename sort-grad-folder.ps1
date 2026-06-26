# ==============================================================================
# Graduation Folder Sorter - Direct Source Version
# Extracts the last word of a folder name and groups them accordingly.
# ==============================================================================

# 1. Define your specific directories here
$sourceDir = "C:\Users\MIS\Downloads\Grad Pic\May\MAY 20-20260604T033919Z-3-001\MAY 20"
$destDir   = "C:\Users\MIS\Downloads\Filtered"

# Create the destination root directory if it doesn't exist
if (-not (Test-Path -Path $destDir)) {
    New-Item -ItemType Directory -Path $destDir | Out-Null
    Write-Host "Created destination root: $destDir" -ForegroundColor Cyan
}

# 2. Grab all student folders directly inside the source directory
$studentFolders = Get-ChildItem -Path $sourceDir -Directory

foreach ($folder in $studentFolders) {
    # 3. Extract the last word
    # Splits the folder name by underscores, spaces, or hyphens and grabs the last piece
    $lastWord = ($folder.Name -split '[-_\s]')[-1].ToUpper()
    
    # Ensure the extracted word isn't blank
    if (-not [string]::IsNullOrWhiteSpace($lastWord)) {
        $targetGroupFolder = Join-Path -Path $destDir -ChildPath $lastWord
        
        # 4. Create the Course/Group folder (e.g., \BSAIS) if it doesn't exist yet
        if (-not (Test-Path -Path $targetGroupFolder)) {
            New-Item -ItemType Directory -Path $targetGroupFolder | Out-Null
            Write-Host "Created new group folder: $lastWord" -ForegroundColor Cyan
        }
        
        # 5. Move the folder to its new home
        $targetPath = Join-Path -Path $targetGroupFolder -ChildPath $folder.Name
        
        if (-not (Test-Path -Path $targetPath)) {
            Move-Item -Path $folder.FullName -Destination $targetGroupFolder
            Write-Host "Moved: $($folder.Name) -> $lastWord" -ForegroundColor Green
        } else {
            Write-Host "Skipped: $($folder.Name) already exists in $lastWord" -ForegroundColor Yellow
        }
    }
}

Write-Host "Sorting complete!" -ForegroundColor Magenta