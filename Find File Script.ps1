# Concise Microsoft Graph File Search Script
# Searches for files across entire tenant

param(
    [Parameter(Mandatory=$false)]
    [int]$ResultSize = 100,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId = "your tenet id",
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId = "your client id",
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret = "your client secret id"
)

# Get access token
function Get-GraphAccessToken {
    param($TenantId, $ClientId, $ClientSecret)
    
    $body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }
    
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
    return @{Authorization = "Bearer $($tokenResponse.access_token)"}
}

# Get all drives via automatic discovery
function Get-AllDrives {
    param($Headers)
    
    $allDrives = @()
    
    # Method 1: Get standard drives (OneDrive, etc.)
    try {
        $drives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/drives" -Headers $Headers
        $allDrives += $drives.value
    } catch {
        Write-Host "Could not get standard drives" -ForegroundColor Yellow
    }
    
    # Method 2: Discover all SharePoint sites automatically
    try {
        Write-Host "Discovering SharePoint sites..." -ForegroundColor Gray
        $allSites = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites" -Headers $Headers
        
        foreach ($site in $allSites.value) {
            try {
                $siteDrives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" -Headers $Headers -ErrorAction Stop
                
                foreach ($drive in $siteDrives.value) {
                    if (-not ($allDrives | Where-Object { $_.id -eq $drive.id })) {
                        $drive | Add-Member -NotePropertyName "siteName" -NotePropertyValue $site.displayName -Force
                        $drive | Add-Member -NotePropertyName "siteId" -NotePropertyValue $site.id -Force
                        $allDrives += $drive
                    }
                }
            } catch {
                # Skip inaccessible site drives
            }
        }
    } catch {
        Write-Host "Could not discover sites" -ForegroundColor Yellow
    }
    
    return $allDrives
}

# Enhanced search matching
function Test-FileMatch {
    param([string]$FileName, [string]$SearchTerm)
    
    if ([string]::IsNullOrWhiteSpace($SearchTerm) -or $SearchTerm -eq "*") {
        return $true
    }
    
    # Split search term and check if all words are present
    $searchWords = $SearchTerm.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    
    foreach ($word in $searchWords) {
        if ($FileName -notlike "*$word*") {
            return $false
        }
    }
    
    return $true
}

# Search files recursively
function Search-DriveFiles {
    param($DriveId, $FolderId, $Headers, $SearchTerm, $SiteId = $null, $MaxDepth = 3, $CurrentDepth = 0)
    
    if ($CurrentDepth -gt $MaxDepth) {
        return @()
    }
    
    $results = @()
    
    try {
        $url = if ($SiteId) {
            "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$FolderId/children"
        } else {
            "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderId/children"
        }
        
        $items = Invoke-RestMethod -Uri $url -Headers $Headers -ErrorAction Stop
        
        foreach ($item in $items.value) {
            if ($item.folder) {
                # Recurse into subfolder
                $subResults = Search-DriveFiles -DriveId $DriveId -FolderId $item.id -Headers $Headers -SearchTerm $SearchTerm -SiteId $SiteId -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
                if ($subResults) {
                    $results += $subResults
                }
            } else {
                # Check if file matches
                if (Test-FileMatch -FileName $item.name -SearchTerm $SearchTerm) {
                    $results += [PSCustomObject]@{
                        Name = $item.name
                        WebUrl = $item.webUrl
                        Size = if($item.size) { "$([math]::Round($item.size/1KB, 2)) KB" } else { "N/A" }
                        Modified = if($item.lastModifiedDateTime) { $item.lastModifiedDateTime } else { "N/A" }
                        Type = if($item.name.Contains('.')) { $item.name.Split('.')[-1].ToUpper() } else { "Unknown" }
                    }
                    Write-Host "  ✅ Found: $($item.name)" -ForegroundColor Green
                }
            }
        }
    } catch {
        # Skip inaccessible folders
    }
    
    return $results
}

# Main execution
try {
    Clear-Host
    Write-Host "Microsoft Graph File Search" -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    
    # Connect
    Write-Host "Connecting..." -ForegroundColor Yellow
    $headers = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    
    # Get drives
    Write-Host "Getting drives..." -ForegroundColor Yellow
    $drives = Get-AllDrives -Headers $headers
    Write-Host "Found $($drives.Count) drives" -ForegroundColor Green
    
    # Get search term
    $SearchQuery = Read-Host "`nEnter search term"
    if ([string]::IsNullOrWhiteSpace($SearchQuery)) {
        Write-Host "No search term entered." -ForegroundColor Red
        exit
    }
    
    Write-Host "`nSearching for: '$SearchQuery'" -ForegroundColor Green
    
    # Search all drives
    $allResults = @()
    foreach ($drive in $drives) {
        Write-Host "Scanning: $($drive.name)..." -ForegroundColor Gray
        
        $driveResults = Search-DriveFiles -DriveId $drive.id -FolderId "root" -Headers $headers -SearchTerm $SearchQuery -SiteId $drive.siteId
        
        if ($driveResults) {
            foreach ($result in $driveResults) {
                $location = if ($drive.siteName) { 
                    "$($drive.siteName) (Site)" 
                } else { 
                    $drive.name 
                }
                $result | Add-Member -NotePropertyName "DriveLocation" -NotePropertyValue $location -Force
            }
            $allResults += $driveResults
            Write-Host "  Found $($driveResults.Count) files" -ForegroundColor Green
        }
        
        if ($allResults.Count -ge $ResultSize) { break }
    }
    
    # Display results
    if ($allResults.Count -gt 0) {
        Write-Host "`n=== RESULTS ===" -ForegroundColor Green
        Write-Host "Found $($allResults.Count) files`n" -ForegroundColor Yellow
        
        $allResults | Sort-Object Modified -Descending | ForEach-Object {
            Write-Host "📄 $($_.Name)" -ForegroundColor White
            Write-Host "   🔗 $($_.WebUrl)" -ForegroundColor Blue
            Write-Host "   📊 $($_.Size) | $($_.Type) | $($_.Modified)" -ForegroundColor Gray
            Write-Host "   📍 $($_.DriveLocation)" -ForegroundColor DarkGray
            Write-Host ""
        }
    } else {
        Write-Host "`n❌ No files found for '$SearchQuery'" -ForegroundColor Red
        Write-Host "Try: 'Azure', 'pdf', 'docx'" -ForegroundColor Gray
    }
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
} finally {
    Read-Host "`nPress Enter to exit"
}