# ========================= 
# CONFIGURATION 
# ========================= 
$ClientID       = ""   # Replace with actual client ID 
$ClientSecret   = ""      # Replace with actual client secret 
$TenantID       = ""   # Replace with actual tenant ID 

# Create date-based folder and filenames 
$today  = Get-Date -Format "yyyy-MM-dd" 
$year   = Get-Date -Format "yyyy" 
$month  = Get-Date -Format "MM" 
$reportFolder = "D:\IntuneSCCMHealthReports\$year\$month" 
New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null 

$ExportCsv  = Join-Path $reportFolder "Windows_Problematic_ConfigMgr_Devices_$today.csv" 
$ExportHtml = Join-Path $reportFolder "Windows_Summary_Report_$today.html" 

# ========================= 
# STEP 1: Authenticate 
# ========================= 
Write-Host "🔐 Authenticating to Microsoft Graph..." -ForegroundColor Cyan 

$body = @{ 
    grant_type    = "client_credentials" 
    scope         = " https://graph.microsoft.com/.default" 
    client_id     = $ClientId 
    client_secret = $ClientSecret 
} 

$tokenResponse = Invoke-RestMethod -Method Post -Uri " https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded' 
$AccessToken = $tokenResponse.access_token 

$Headers = @{ 
    Authorization    = "Bearer $AccessToken" 
    ConsistencyLevel = "eventual" 
} 

# ========================= 
# STEP 2: Retrieve All Devices with Pagination 
# ========================= 
Write-Host "📡 Retrieving all managed devices..." -ForegroundColor Cyan 
$allDevices = @() 
$uri = " https://graph.microsoft.com/beta/deviceManagement/managedDevices?$top=1000" 
$page = 1 

do { 
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $Headers 
    if ($response.value) { 
        $allDevices += $response.value 
        $percent = [Math]::Min((($allDevices.Count / 10000) * 100), 100) 
        Write-Progress -Activity "Retrieving Devices" -Status "Page $page" -PercentComplete $percent 
        $page++ 
    } 
    $uri = $response.'@odata.nextLink' 
} while ($uri) 

$totalDevices = $allDevices.Count 
Write-Host "📋 Total devices retrieved: $totalDevices" -ForegroundColor Green 

# ========================= 
# STEP 3: Filter Windows Devices 
# ========================= 
$windowsDevices = $allDevices | Where-Object { 
    $_.operatingSystem -like "Windows*" 
} 

$totalWindows = $windowsDevices.Count 
Write-Host "🖥️  Windows devices found: $totalWindows" -ForegroundColor Cyan 

# ========================= 
# STEP 4: Filter Problematic ConfigMgr Devices 
# ========================= 
$problemDevices = $windowsDevices | Where-Object { 
    $_.configurationManagerClientHealthState.state -in @("communicationError", "unknown") 
} 

$problemCount = $problemDevices.Count 
Write-Host "⚠️  Problematic Windows devices: $problemCount" -ForegroundColor Yellow 

# ========================= 
# STEP 5: Export Problem Devices to CSV 
# ========================= 
$problemDevices | ForEach-Object { 
    [PSCustomObject]@{ 
        deviceName                         = $_.deviceName 
        userPrincipalName                  = $_.userPrincipalName 
        managementAgent                    = $_.managementAgent 
        operatingSystem                    = $_.operatingSystem 
        complianceState                    = $_.complianceState 
        enrolledDateTime                   = $_.enrolledDateTime 
        lastSyncDateTime                   = $_.lastSyncDateTime 
        configMgrClientEnabled             = $_.configurationManagerClientEnabled 
        configMgrHealthState               = $_.configurationManagerClientHealthState.state 
        configMgrHealthErrorCode           = $_.configurationManagerClientHealthState.errorCode 
        configMgrHealthLastSyncDateTime    = $_.configurationManagerClientHealthState.lastSyncDateTime 
    } 
} | Export-Csv -Path $ExportCsv -NoTypeInformation 

Write-Host "✅ Exported to CSV: $ExportCsv" -ForegroundColor Green 

# ========================= 
# STEP 6: Create HTML Summary 
# ========================= 
$stateSummary = $windowsDevices | 
    Where-Object { $_.configurationManagerClientHealthState.state } | 
    Group-Object { $_.configurationManagerClientHealthState.state } | 
    Sort-Object Name 

$htmlBody = @" 
<!DOCTYPE html> 
<html> 
<head> 
    <style> 
        body { font-family: Segoe UI, sans-serif; } 
        h1 { color: #2e6c80; } 
        table { border-collapse: collapse; width: 50%; } 
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; } 
        th { background-color: #f2f2f2; } 
        .summary { margin-bottom: 20px; } 
    </style> 
</head> 
<body> 
    <h1>Intune ConfigMgr Health Summary (Windows Devices)</h1> 
    <div class="summary"> 
        <p><strong>Date:</strong> $today</p> 
        <p><strong>Total Devices (All OS):</strong> $totalDevices</p> 
        <p><strong>Windows Devices:</strong> $totalWindows</p> 
        <p><strong>Problematic Windows Devices Exported:</strong> $problemCount</p> 
    </div> 
    <h2>Configuration Manager Health Breakdown</h2> 
    <table> 
        <tr><th>Health State</th><th>Device Count</th></tr> 
"@ 

foreach ($item in $stateSummary) { 
    $htmlBody += "<tr><td>$($item.Name)</td><td>$($item.Count)</td></tr>" 
} 

$htmlBody += @" 
    </table> 
</body> 
</html> 
"@ 

# Save HTML to disk (optional record) 
$htmlBody | Out-File -FilePath $ExportHtml -Encoding UTF8 
Write-Host "HTML summary saved to: $ExportHtml" -ForegroundColor Green 
