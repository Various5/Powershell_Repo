# Revert Cloud Save Removal from Office 365
# Run as Administrator
# Restores all cloud functionality

Write-Host "Reverting Office 365 Cloud Save Changes..." -ForegroundColor Yellow
Write-Host "This will restore all cloud functionality including OneDrive" -ForegroundColor Cyan

# Registry paths
$OfficeRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0"
$CloudRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0"

# Restore OneDrive and cloud locations in Save/Open dialogs
Write-Host "`nRestoring OneDrive in Save/Open dialogs..."
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "PreferCloudSaveLocations" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "SkipOpenAndSaveAsPlace" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "PreferOneDriveOverEnterpriseDocLibs" -Force -ErrorAction SilentlyContinue

# Restore cloud locations in file dialogs
Remove-ItemProperty -Path "$OfficeRegPath\Common\Internet" -Name "UseOnlineContent" -Force -ErrorAction SilentlyContinue

# Restore OneDrive places (including corporate)
Write-Host "Restoring corporate OneDrive and SharePoint locations..."
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "ShowSkyDriveSaveAsPlace" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "ShowSkyDriveOpenPlace" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\Place Bar" -Name "NoPlaceBar" -Force -ErrorAction SilentlyContinue

# Re-enable SharePoint and OneDrive for Business integration
Remove-ItemProperty -Path "$OfficeRegPath\Common\Toolbars" -Name "NoSharePointPlaces" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\Open Find\Places\StandardPlaces" -Name "DisableSharePointPlaces" -Force -ErrorAction SilentlyContinue

# Re-enable AutoSave to cloud
Write-Host "Re-enabling AutoSave to cloud..."
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "AutoSaveDefaultSetting" -Force -ErrorAction SilentlyContinue

# Re-enable cloud-based file sharing
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableFileSharing" -Force -ErrorAction SilentlyContinue

# Restore "Save to Cloud" button
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "HideSaveToCloud" -Force -ErrorAction SilentlyContinue

# Re-enable "Sites" section in Save dialog
Write-Host "Re-enabling Sites section in Save dialog..."
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableDiscoverLocations" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "HideSitesTab" -Force -ErrorAction SilentlyContinue

# Re-enable Connected Experiences
Write-Host "Re-enabling cloud-connected features..."
Remove-ItemProperty -Path "$CloudRegPath\Common\Privacy" -Name "DownloadContentDisabled" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "$CloudRegPath\Common\Privacy" -Name "UserContentDisabled" -Force -ErrorAction SilentlyContinue

# Restore per-application cloud options
$apps = @("Word", "Excel", "PowerPoint", "OneNote")
foreach ($app in $apps) {
    $appPath = "$OfficeRegPath\$app"
    
    # Re-enable cloud document features
    Remove-ItemProperty -Path "$appPath\Options" -Name "DisableRoamingFileStorage" -Force -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "$appPath\Options" -Name "LocalOnly" -Force -ErrorAction SilentlyContinue
    
    # Restore Places in Save/Open dialogs
    Remove-ItemProperty -Path "$appPath\Save As\Settings" -Name "ItemsToHide" -Force -ErrorAction SilentlyContinue
}

# Restore OneDrive places for current user
Write-Host "Restoring OneDrive places for current user..."
$userRegPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
Remove-ItemProperty -Path $userRegPath -Name "PreferCloudSaveLocations" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $userRegPath -Name "ShowSkyDriveSaveAsPlace" -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $userRegPath -Name "ShowSkyDriveOpenPlace" -Force -ErrorAction SilentlyContinue

# Restore OneDrive in Windows Explorer sidebar
Write-Host "Restoring OneDrive in Explorer sidebar..."
Set-ItemProperty -Path "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Name "System.IsPinnedToNameSpaceTree" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Name "System.IsPinnedToNameSpaceTree" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue

# Re-enable Office Start screen
Write-Host "Re-enabling Office Start screen..."
Remove-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableBootToOfficeStart" -Force -ErrorAction SilentlyContinue

# Remove custom default save locations (restore to Office defaults)
$apps = @("Word", "Excel", "PowerPoint")
foreach ($app in $apps) {
    $appPath = "$OfficeRegPath\$app\Options"
    Remove-ItemProperty -Path $appPath -Name "DefaultPath" -Force -ErrorAction SilentlyContinue
}

# Clean up empty registry keys
Write-Host "`nCleaning up empty registry keys..."
$cleanupPaths = @(
    "$OfficeRegPath\Common\Place Bar",
    "$OfficeRegPath\Common\Toolbars",
    "$OfficeRegPath\Common\Open Find\Places\StandardPlaces",
    "$OfficeRegPath\Common\Open Find\Places",
    "$OfficeRegPath\Common\Open Find",
    "$CloudRegPath\Common\Privacy",
    "$CloudRegPath\Common",
    "$CloudRegPath"
)

foreach ($path in $cleanupPaths) {
    if (Test-Path $path) {
        $properties = Get-Item $path | Select-Object -ExpandProperty Property
        if ($properties.Count -eq 0) {
            Remove-Item $path -Force -ErrorAction SilentlyContinue
        }
    }
}

# Force update group policy
Write-Host "`nUpdating group policy..."
gpupdate /force

Write-Host "`n✓ All changes have been reverted!" -ForegroundColor Green
Write-Host "`nWhat has been restored:" -ForegroundColor Cyan
Write-Host "- OneDrive appears in Save/Open dialogs"
Write-Host "- AutoSave to cloud is available"
Write-Host "- Cloud storage options are visible"
Write-Host "- Office Start screen shows recent files"
Write-Host "- Sites section in Save dialog is available"
Write-Host "- Connected experiences are enabled"

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Restart all Office applications"
Write-Host "2. Sign in to Office if needed: File > Account > Sign In"
Write-Host "3. OneDrive may need to be restarted manually"

# Optional: Restart OneDrive
$restart = Read-Host "`nDo you want to restart OneDrive now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    $onedrivePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
    if (Test-Path $onedrivePath) {
        Write-Host "Starting OneDrive..." -ForegroundColor Green
        Start-Process $onedrivePath
    } else {
        Write-Host "OneDrive not found at default location. Please start it manually." -ForegroundColor Yellow
    }
}
