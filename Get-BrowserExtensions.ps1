#
# Find all non-default browser extensions for the active user.
#

# --- Configuration ---
# Set to $true to save the CSV locally. Set to $false to upload to Azure.
$testMode = $true

# Path for the local CSV report (only used in test mode)
$savePath = "C:\ProgramData\ExtensionReports"

# Azure details (only used when $testMode is $false)
$blobContainerUrl = "YOUR_AZURE_CONTAINER_URL_HERE"


# --- Functions ---

function Get-BrowserExtensions {
    param(
        $BrowserName,
        $ExtensionsPath,
        $IgnoredIds
    )

    if (-not (Test-Path $ExtensionsPath)) {
        Write-Host "-> Path not found, skipping."
        return
    }
    
    $found = @()
    # Get manifests, sort by newest version first to ensure we keep the latest one
    $manifests = Get-ChildItem -Path $ExtensionsPath -Filter 'manifest.json' -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending
    
    if (-not $manifests) {
        Write-Host "-> No extensions found."
        return
    }

    foreach ($manifestFile in $manifests) {
        try {
            $extensionId = $manifestFile.Directory.Parent.Name
            if ($IgnoredIds -contains $extensionId) {
                continue
            }

            $manifestContent = Get-Content -Path $manifestFile.FullName -Raw | ConvertFrom-Json
            $extName = $manifestContent.name

            # If the name is a language key like __MSG_extName__, find the real name
            if ($extName -like "__MSG_*__") {
                $msgKey = $extName -replace "__MSG_(.*)__", '$1'
                $localeFile = Join-Path $manifestFile.DirectoryName "_locales\en\messages.json"
                if (Test-Path $localeFile) {
                    $localeContent = Get-Content -Path $localeFile -Raw | ConvertFrom-Json
                    $resolvedName = $localeContent.PSObject.Properties[$msgKey].Value.message
                    if ($resolvedName) {
                        $extName = $resolvedName
                    }
                }
            }

            if ($extName) {
                $found += [PSCustomObject]@{
                    ExtensionID = $extensionId
                    Name        = $extName
                    Browser     = $BrowserName
                }
            }
        }
        catch {
            # Silently continue if one manifest is broken
        }
    }
    return $found
}

function Get-ActiveUserAppData {
    try {
        $session = Get-CimInstance -ClassName Win32_LogonSession -Filter "LogonType = 2" | Select-Object -First 1
        $user = $session | Get-CimAssociatedInstance -ResultClassName Win32_UserAccount
        $profilePath = Get-ItemPropertyValue -Path "Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($user.SID)" -Name "ProfileImagePath"
        $localAppData = Join-Path -Path $profilePath -ChildPath 'AppData\Local'
        if (Test-Path $localAppData) {
            Write-Host "Found active user profile: $($user.Name)"
            return $localAppData
        }
    }
    catch {}
    Write-Error "Could not find active user's AppData path."
    return $null
}


# --- Main ---

Write-Host "Starting extension scan on $($env:COMPUTERNAME)..."
$localAppData = Get-ActiveUserAppData
if (-not $localAppData) {
    exit 1
}

# List of browsers and their extension paths
$browsersToCheck = @{
    'Edge'   = Join-Path $localAppData 'Microsoft\Edge\User Data\Default\Extensions';
    'Chrome' = Join-Path $localAppData 'Google\Chrome\User Data\Default\Extensions';
}

# Default extensions to ignore
$ignoredExts = @(
    'nmmhkkegccagdldgiimedpiccmgmieda', # Chrome Web Store
    'ghbmnnjooekpmoecnnnilnnbdlolhkhi', # Google Docs Offline
    'jmjflgjpcpepeafmmgdpfkogkghcpiha', # Office
    'dppgmdbiimibapkepcbdbmkaabgiofem'  # Microsoft Bing Search
)

$allFoundExtensions = @()
# This loop ensures both browsers are checked sequentially
foreach ($browser in $browsersToCheck.GetEnumerator()) {
    Write-Host "--- Checking $($browser.Name) extensions ---"
    $allFoundExtensions += Get-BrowserExtensions -BrowserName $browser.Name -ExtensionsPath $browser.Value -IgnoredIds $ignoredExts
}

if ($allFoundExtensions.Count -eq 0) {
    Write-Host "Scan complete. No non-default extensions found."
    exit 0
}

# Filter so it only shows one of each extension
$uniqueExtensions = $allFoundExtensions | Sort-Object -Property ExtensionID -Unique

Write-Host "Found $($uniqueExtensions.Count) unique extension(s)."

if ($testMode) {
    # Save the report locally
    if (-not (Test-Path $savePath)) {
        New-Item -Path $savePath -ItemType Directory -Force | Out-Null
    }
    $csvFile = Join-Path $savePath "Extensions_$($env:COMPUTERNAME).csv"
    $uniqueExtensions | Export-Csv -Path $csvFile -NoTypeInformation -Force
    Write-Host "SUCCESS: Report saved to $csvFile"
}
else {
    # Upload the report to Azure
    if (-not $blobContainerUrl -or $blobContainerUrl -eq "YOUR_AZURE_CONTAINER_URL_HERE") {
        Write-Error "Azure upload is enabled, but blob container URL is not set."
        exit 1
    }
    $csvContent = $uniqueExtensions | ConvertTo-Csv -NoTypeInformation
    $uploadUrl = "$($blobContainerUrl.TrimEnd('/'))/Extensions_$($env:COMPUTERNAME).csv"
    $headers = @{ 'x-ms-blob-type' = 'BlockBlob' }
    Invoke-RestMethod -Method Put -Uri $uploadUrl -Body $csvContent -Headers $headers -ContentType 'text/csv; charset=utf-8'
    Write-Host "SUCCESS: Report uploaded to Azure."
}
