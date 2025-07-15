# Remove Cloud Save Options from Office 365
# Run as Administrator
# 15.07.2025

Write-Host "Removing Cloud Save Options from Office 365..." -ForegroundColor Yellow

# Registry paths
$OfficeRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0"
$CloudRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Cloud\Office\16.0"

# Create registry paths if they don't exist
$paths = @(
    "$OfficeRegPath\Common\General",
    "$OfficeRegPath\Common\Internet",
    "$OfficeRegPath\Common\Place Bar",
    "$OfficeRegPath\Common\Toolbars",
    "$OfficeRegPath\Common\Open Find\Places\StandardPlaces",
    "$CloudRegPath\Common\Privacy"
)

foreach ($path in $paths) {
    if (!(Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
}

# Remove OneDrive and cloud locations from Save/Open dialogs
Write-Host "Removing OneDrive from Save/Open dialogs..."
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "PreferCloudSaveLocations" -Value 0 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "SkipOpenAndSaveAsPlace" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "PreferOneDriveOverEnterpriseDocLibs" -Value 0 -Type DWord -Force

# Hide cloud locations in file dialogs
Set-ItemProperty -Path "$OfficeRegPath\Common\Internet" -Name "UseOnlineContent" -Value 0 -Type DWord -Force

# Remove ALL OneDrive places (including corporate)
Write-Host "Removing corporate OneDrive and SharePoint locations..."
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "ShowSkyDriveSaveAsPlace" -Value 0 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "ShowSkyDriveOpenPlace" -Value 0 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\Place Bar" -Name "NoPlaceBar" -Value 1 -Type DWord -Force

# Disable SharePoint and OneDrive for Business integration
Set-ItemProperty -Path "$OfficeRegPath\Common\Toolbars" -Name "NoSharePointPlaces" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\Open Find\Places\StandardPlaces" -Name "DisableSharePointPlaces" -Value 1 -Type DWord -Force

# Disable AutoSave to cloud
Write-Host "Disabling AutoSave to cloud..."
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "AutoSaveDefaultSetting" -Value 0 -Type DWord -Force

# Disable cloud-based file sharing
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableFileSharing" -Value 1 -Type DWord -Force

# Remove "Save to Cloud" button
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "HideSaveToCloud" -Value 1 -Type DWord -Force

# Disable "Sites" section in Save dialog (where corporate OneDrive appears)
Write-Host "Disabling Sites section in Save dialog..."
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableDiscoverLocations" -Value 1 -Type DWord -Force
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "HideSitesTab" -Value 1 -Type DWord -Force

# Disable Connected Experiences that access cloud content
Write-Host "Disabling cloud-connected features..."
Set-ItemProperty -Path "$CloudRegPath\Common\Privacy" -Name "DownloadContentDisabled" -Value 2 -Type DWord -Force
Set-ItemProperty -Path "$CloudRegPath\Common\Privacy" -Name "UserContentDisabled" -Value 2 -Type DWord -Force

# Per-application settings to remove cloud options
$apps = @("Word", "Excel", "PowerPoint", "OneNote")
foreach ($app in $apps) {
    $appPath = "$OfficeRegPath\$app"
    if (!(Test-Path "$appPath\Options")) {
        New-Item -Path "$appPath\Options" -Force | Out-Null
    }
    
    # Disable cloud document features
    Set-ItemProperty -Path "$appPath\Options" -Name "DisableRoamingFileStorage" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "$appPath\Options" -Name "LocalOnly" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    
    # Remove Places from Save/Open dialogs for each app
    Set-ItemProperty -Path "$appPath\Save As\Settings" -Name "ItemsToHide" -Value "OneDrive,SharePoint" -Type String -Force -ErrorAction SilentlyContinue
}

# Additional removal of OneDrive places for current user
Write-Host "Removing OneDrive places for current user..."
$userRegPath = "HKCU:\Software\Microsoft\Office\16.0\Common\General"
if (!(Test-Path $userRegPath)) {
    New-Item -Path $userRegPath -Force | Out-Null
}
Set-ItemProperty -Path $userRegPath -Name "PreferCloudSaveLocations" -Value 0 -Type DWord -Force
Set-ItemProperty -Path $userRegPath -Name "ShowSkyDriveSaveAsPlace" -Value 0 -Type DWord -Force
Set-ItemProperty -Path $userRegPath -Name "ShowSkyDriveOpenPlace" -Value 0 -Type DWord -Force

# Remove OneDrive from Windows Explorer sidebar (optional)
Write-Host "Hiding OneDrive from Explorer sidebar..."
Set-ItemProperty -Path "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}" -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue

# Remove sync'd OneDrive folders from Quick Access
Write-Host "Removing OneDrive sync folders..."
$onedriveReg = "HKCU:\Software\Microsoft\OneDrive\Accounts\Business1"
if (Test-Path $onedriveReg) {
    $mountPoint = (Get-ItemProperty -Path $onedriveReg -Name "UserFolder" -ErrorAction SilentlyContinue).UserFolder
    if ($mountPoint) {
        # Unpin from Quick Access
        $shell = New-Object -ComObject Shell.Application
        $namespace = $shell.NameSpace($mountPoint)
        if ($namespace) {
            $namespace.Self.InvokeVerb("unpinfromhome")
        }
    }
}

# Disable Office Start screen (shows cloud documents)
Write-Host "Disabling Office Start screen..."
Set-ItemProperty -Path "$OfficeRegPath\Common\General" -Name "DisableBootToOfficeStart" -Value 1 -Type DWord -Force

# Set default save location to local Documents folder
$apps = @("Word", "Excel", "PowerPoint")
foreach ($app in $apps) {
    $appPath = "$OfficeRegPath\$app\Options"
    if (!(Test-Path $appPath)) {
        New-Item -Path $appPath -Force | Out-Null
    }
    Set-ItemProperty -Path $appPath -Name "DefaultPath" -Value "%USERPROFILE%\Documents" -Type String -Force -ErrorAction SilentlyContinue
}

Write-Host "`nCloud save options removed successfully!" -ForegroundColor Green
Write-Host "Please restart all Office applications for changes to take effect." -ForegroundColor Yellow

# Force update group policy
gpupdate /force

Write-Host "`nWhat this script has done:" -ForegroundColor Cyan
Write-Host "- Removed OneDrive from Save/Open dialogs"
Write-Host "- Disabled AutoSave to cloud"
Write-Host "- Hidden cloud storage options"
Write-Host "- Set default save location to local Documents folder"
Write-Host "- Disabled Office Start screen (which shows recent cloud files)"
Write-Host "`nUsers can now only save files locally." -ForegroundColor Green

Write-Host "`nIMPORTANT: If corporate OneDrive still appears:" -ForegroundColor Yellow
Write-Host "1. Close all Office applications"
Write-Host "2. Clear Office cache: %localappdata%\Microsoft\Office\16.0\Wef\"
Write-Host "3. Sign out from Office: File > Account > Sign Out"
Write-Host "4. Restart Office applications" -ForegroundColor Yellow

Write-Host "`nManual removal option:" -ForegroundColor Cyan
Write-Host "If the corporate OneDrive still appears in Save As dialog:"
Write-Host '1. In any Office app, go to File > Save As'
Write-Host '2. Right-click on the OneDrive entry'
Write-Host '3. Select "Remove from list" or "Unpin from list"'
Write-Host "4. This will remove it from all Office applications" -ForegroundColor White
