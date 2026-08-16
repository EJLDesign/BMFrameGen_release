#Requires -Version 5.1
<#
.SYNOPSIS
    TaylorFab Studio Installer - Installs the TaylorFab Studio frame generator plugin for AutoCAD.
.DESCRIPTION
    Detects installed AutoCAD versions, downloads the latest release,
    places the DLL and model library, and registers for autoload via registry.
.NOTES
    Run with: iex (irm https://raw.githubusercontent.com/EJLDesign/TaylorFab_Studio_release/main/install.ps1)
    Or: .\install.ps1
#>

$ErrorActionPreference = 'Stop'

trap {
    Write-Host ""
    Write-Host "  ERROR: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# -- Config -------------------------------------------------------------------
$RepoOwner   = "EJLDesign"
$RepoName    = "TaylorFab_Studio_release"
$PluginName  = "BMFrameGenCAD"
$BundleDir   = Join-Path $env:APPDATA "Autodesk\ApplicationPlugins\BMFrameGenCAD.bundle"
$InstallDir  = Join-Path $BundleDir "Contents"

# -- Functions ----------------------------------------------------------------

function Write-Banner {
    Write-Host ""
    Write-Host "  +--------------------------------------------+" -ForegroundColor Cyan
    Write-Host "  |     TaylorFab Studio Plugin Installer       |" -ForegroundColor Cyan
    Write-Host "  |     Frame Generator for AutoCAD             |" -ForegroundColor Cyan
    Write-Host "  +--------------------------------------------+" -ForegroundColor Cyan
    Write-Host ""
}

function Get-InstalledAutoCADVersions {
    $versions = @()
    $acadKey = "HKLM:\SOFTWARE\Autodesk\AutoCAD"

    if (-not (Test-Path $acadKey)) {
        return $versions
    }

    Get-ChildItem $acadKey | ForEach-Object {
        $versionKey = $_
        $versionId = $versionKey.PSChildName  # e.g. R25.0
        Get-ChildItem $versionKey.PSPath | ForEach-Object {
            $productKey = $_
            $productId = $productKey.PSChildName  # e.g. ACAD-8101:409
            $props = Get-ItemProperty $productKey.PSPath -ErrorAction SilentlyContinue
            $productName = $props.ProductName
            $installPath = $props.AcadLocation

            if ($productName -and $installPath -and (Test-Path $installPath)) {
                $versions += [PSCustomObject]@{
                    Name        = $productName
                    VersionId   = $versionId
                    ProductId   = $productId
                    InstallPath = $installPath
                }
            }
        }
    }
    return $versions
}

function Get-LatestRelease {
    Write-Host "  Checking for latest release..." -ForegroundColor Gray
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest" -UseBasicParsing
        $asset = $release.assets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1

        if (-not $asset) {
            throw "No zip asset found in latest release."
        }

        return [PSCustomObject]@{
            TagName     = $release.tag_name
            DownloadUrl = $asset.browser_download_url
            AssetName   = $asset.name
        }
    }
    catch {
        throw "Could not fetch latest release. Make sure a release exists at https://github.com/$RepoOwner/$RepoName/releases -- $_"
    }
}

function Install-Plugin {
    param(
        [string]$DownloadUrl,
        [string]$TagName
    )

    # Clean previous install (wait for the delete to actually finish --
    # Windows deletes are async and antivirus handles can delay them)
    if (Test-Path $BundleDir) {
        Write-Host "  Removing previous installation..." -ForegroundColor Yellow
        Remove-Item $BundleDir -Recurse -Force
        for ($i = 0; $i -lt 25 -and (Test-Path $BundleDir); $i++) { Start-Sleep -Milliseconds 200 }
        if (Test-Path $BundleDir) {
            throw "Could not remove the previous installation at '$BundleDir'. Close AutoCAD and any Explorer windows in that folder, then run the installer again."
        }
    }

    New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null

    # Write PackageContents.xml for AutoCAD bundle autoload
    $packageXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ApplicationPackage SchemaVersion="1.0"
  Name="BMFrameGenCAD"
  Description="TaylorFab Studio — Frame Generator for AutoCAD"
  AppVersion="1.0.0"
  ProductCode="{B4F2E8A1-7C3D-4E5F-9A1B-2D3E4F5A6B7C}">
  <CompanyDetails Name="EJL Design"/>
  <RuntimeRequirements OS="Win64" Platform="AutoCAD"/>
  <Components>
    <ComponentEntry AppName="BMFrameGenCAD"
      ModuleName="./Contents/BMFrameGenCAD.dll"
      AppType="Managed"
      LoadOnAutoCADStartup="True"/>
  </Components>
</ApplicationPackage>
"@
    Set-Content (Join-Path $BundleDir "PackageContents.xml") $packageXml -Encoding UTF8

    # Best-effort cleanup of debris from older runs (never reuse these paths --
    # a failed delete or an antivirus handle on them must not break this run)
    Get-ChildItem $env:TEMP -Filter "$PluginName-extract*" -Directory -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:TEMP -Filter "$PluginName-v*.zip" -File -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue

    # Download and extract, using paths unique to this run so leftovers from a
    # previous attempt (possibly still locked by antivirus) can't collide
    $runId = Get-Random
    $tempZip = Join-Path $env:TEMP "$PluginName-$TagName-$runId.zip"
    Write-Host "  Downloading $TagName..." -ForegroundColor Gray
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempZip -UseBasicParsing

    # Extract with retries: antivirus scanners often hold just-written files
    # briefly, which makes Expand-Archive fail spuriously
    $tempExtract = Join-Path $env:TEMP "$PluginName-extract-$runId"
    for ($attempt = 1; ; $attempt++) {
        try {
            Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
            break
        }
        catch {
            if ($attempt -ge 3) { throw }
            Write-Host "  Extraction attempt $attempt failed (antivirus may be scanning); retrying..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
            $tempExtract = Join-Path $env:TEMP "$PluginName-extract-$runId-$attempt"
        }
    }

    # Copy DLL
    $dll = Get-ChildItem $tempExtract -Filter "$PluginName.dll" -Recurse | Select-Object -First 1
    if (-not $dll) {
        throw "Could not find $PluginName.dll in the release archive."
    }
    Copy-Item $dll.FullName -Destination $InstallDir
    Write-Host "  Installed $PluginName.dll" -ForegroundColor Gray

    # Copy the site license beside the DLL. The plugin auto-imports a license.lic
    # sitting next to the loaded DLL into %APPDATA%\BMFrameGen on startup, so
    # plugin updates renew licenses silently. Best-effort: an archive without a
    # license (older releases) must not break the install.
    $lic = Get-ChildItem $tempExtract -Filter "license.lic" -Recurse | Select-Object -First 1
    if ($lic) {
        Copy-Item $lic.FullName -Destination $InstallDir
        Write-Host "  Installed license.lic (auto-imports on first load)." -ForegroundColor Gray
    }
    else {
        Write-Host "  NOTE: No license.lic in release -- existing license (if any) stays in effect." -ForegroundColor Yellow
    }

    # Copy the End User License Agreement beside the DLL (human-readable copy;
    # the plugin also carries the text embedded in the DLL). Best-effort: older
    # archives without it must not break the install.
    $eula = Get-ChildItem $tempExtract -Filter "EULA.txt" -Recurse | Select-Object -First 1
    if ($eula) {
        Copy-Item $eula.FullName -Destination $InstallDir
        Write-Host "  Installed EULA.txt" -ForegroundColor Gray
    }

    # Copy assets (sheet-label symbol DWGs). SheetSymbolService loads these from
    # <InstallDir>\assets\symbols\ at sheet-generation time; without them, sheets
    # generate with viewports but no labels.
    $assetsSource = Get-ChildItem $tempExtract -Directory -Filter "assets" -Recurse | Select-Object -First 1
    if ($assetsSource) {
        $assetsDest = Join-Path $InstallDir "assets"
        Copy-Item $assetsSource.FullName -Destination $assetsDest -Recurse
        $symbolCount = (Get-ChildItem (Join-Path $assetsDest "symbols") -Filter "*.dwg" -ErrorAction SilentlyContinue).Count
        Write-Host "  Installed $symbolCount sheet-label symbol file(s)." -ForegroundColor Gray
    }
    else {
        Write-Host "  WARNING: No assets folder found in release -- sheet labels will not appear." -ForegroundColor Yellow
    }

    # Copy Models
    $modelsSource = Get-ChildItem $tempExtract -Directory -Filter "Models" -Recurse | Select-Object -First 1
    if ($modelsSource) {
        $modelsDest = Join-Path $InstallDir "Models"
        Copy-Item $modelsSource.FullName -Destination $modelsDest -Recurse
        $modelCount = (Get-ChildItem $modelsDest -Filter "*.dwg").Count
        Write-Host "  Installed $modelCount model files." -ForegroundColor Gray
    }
    else {
        Write-Host "  WARNING: No Models folder found in release." -ForegroundColor Yellow
    }

    # Write model library path to settings
    $settingsFile = Join-Path $env:APPDATA "bmframegen_settings.txt"
    $modelsPath = Join-Path $InstallDir "Models"
    if (Test-Path $settingsFile) {
        $content = Get-Content $settingsFile -Raw
        if ($content -match "(?m)^LibraryPath=") {
            $content = $content -replace "(?m)^LibraryPath=.*$", "LibraryPath=$modelsPath"
        }
        else {
            $content = "LibraryPath=$modelsPath`n$content"
        }
        Set-Content $settingsFile $content -NoNewline
    }
    else {
        Set-Content $settingsFile "LibraryPath=$modelsPath"
    }

    # Clean up temp files
    Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
    Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
}

function Register-AutoLoad {
    param(
        [array]$AutoCADVersions
    )

    $dllPath = Join-Path $InstallDir "$PluginName.dll"
    $registered = 0

    foreach ($acad in $AutoCADVersions) {
        # Write to HKCU (no admin needed) - same structure as HKLM autoload
        $regPath = "HKCU:\SOFTWARE\Autodesk\AutoCAD\$($acad.VersionId)\$($acad.ProductId)\Applications\$PluginName"

        New-Item -Path $regPath -Force | Out-Null
        Set-ItemProperty -Path $regPath -Name "DESCRIPTION" -Value "TaylorFab Studio for AutoCAD"
        Set-ItemProperty -Path $regPath -Name "LOADCTRLS" -Value 2 -Type DWord
        Set-ItemProperty -Path $regPath -Name "LOADER" -Value $dllPath
        Set-ItemProperty -Path $regPath -Name "MANAGED" -Value 1 -Type DWord

        Write-Host "  Registered autoload for $($acad.Name)" -ForegroundColor Green
        $registered++
    }

    return $registered
}

function Test-Installation {
    param([array]$AutoCADVersions)

    $ok = $true

    # Check DLL
    $dllPath = Join-Path $InstallDir "$PluginName.dll"
    if (Test-Path $dllPath) {
        Write-Host "  [OK] Plugin DLL installed" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Plugin DLL missing" -ForegroundColor Red
        $ok = $false
    }

    # Check Models
    $modelsPath = Join-Path $InstallDir "Models"
    if ((Test-Path $modelsPath) -and (Get-ChildItem $modelsPath -Filter "*.dwg").Count -gt 0) {
        Write-Host "  [OK] Model library installed" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Model library missing or empty" -ForegroundColor Yellow
    }

    # Check sheet-label symbols
    $symbolsPath = Join-Path $InstallDir "assets\symbols"
    if ((Test-Path $symbolsPath) -and (Get-ChildItem $symbolsPath -Filter "*.dwg").Count -gt 0) {
        Write-Host "  [OK] Sheet-label symbols installed" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Sheet-label symbols missing -- generated sheets will have no labels" -ForegroundColor Yellow
    }

    # Check registry entries
    foreach ($acad in $AutoCADVersions) {
        $regPath = "HKCU:\SOFTWARE\Autodesk\AutoCAD\$($acad.VersionId)\$($acad.ProductId)\Applications\$PluginName"
        if (Test-Path $regPath) {
            Write-Host "  [OK] Registry entry for $($acad.Name)" -ForegroundColor Green
        } else {
            Write-Host "  [FAIL] Registry entry missing for $($acad.Name)" -ForegroundColor Red
            $ok = $false
        }
    }

    # Check settings
    $settingsFile = Join-Path $env:APPDATA "bmframegen_settings.txt"
    if (Test-Path $settingsFile) {
        Write-Host "  [OK] Settings configured" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Settings file not found" -ForegroundColor Yellow
    }

    return $ok
}

# -- Main ---------------------------------------------------------------------

Write-Banner

Write-Host "  Detecting installed AutoCAD versions..." -ForegroundColor Gray
$acadVersions = Get-InstalledAutoCADVersions

if ($acadVersions.Count -eq 0) {
    Write-Host ""
    Write-Host "  No AutoCAD installations detected." -ForegroundColor Red
    Write-Host "  AutoCAD must be installed before running this installer." -ForegroundColor Red
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "  Found AutoCAD installations:" -ForegroundColor Green
foreach ($v in $acadVersions) {
    Write-Host "    - $($v.Name)" -ForegroundColor White
}
Write-Host ""

# Get latest release
$release = Get-LatestRelease
Write-Host "  Latest version: $($release.TagName)" -ForegroundColor Cyan
Write-Host ""

# Confirm
Write-Host "  Install TaylorFab Studio $($release.TagName)? (Y/n) " -ForegroundColor White -NoNewline
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host $key.Character
if ($key.Character -eq 'n' -or $key.Character -eq 'N') {
    Write-Host "  Installation cancelled." -ForegroundColor Yellow
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

Write-Host ""

# Install files
Install-Plugin -DownloadUrl $release.DownloadUrl -TagName $release.TagName

# Register autoload in registry for each detected AutoCAD version
Write-Host ""
Write-Host "  Registering plugin with AutoCAD..." -ForegroundColor Gray
$regCount = Register-AutoLoad -AutoCADVersions $acadVersions

# Verify
Write-Host ""
Write-Host "  Verifying installation..." -ForegroundColor Gray
$passed = Test-Installation -AutoCADVersions $acadVersions

Write-Host ""
if ($passed) {
    Write-Host "  Installation complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Plugin location: $BundleDir" -ForegroundColor Gray
    Write-Host "  Registered for $regCount AutoCAD version(s)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Next steps:" -ForegroundColor Cyan
    Write-Host "    1. Launch (or restart) AutoCAD" -ForegroundColor White
    Write-Host '    2. Type "BMFrameGen" in the command line to start' -ForegroundColor White
    Write-Host ""
}
else {
    Write-Host "  Installation completed with errors. Check the messages above." -ForegroundColor Red
}

Write-Host "  Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
