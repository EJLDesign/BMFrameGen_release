#Requires -Version 5.1
<#
.SYNOPSIS
    BMFrameGenCAD Installer - Installs the beMatrix Frame Generator plugin for AutoCAD.
.DESCRIPTION
    Detects installed AutoCAD versions, lets you choose which to install for,
    and sets up the plugin with autoloading and model library.
.NOTES
    Run with: irm https://raw.githubusercontent.com/EJLDesign/BMFrameGen_release/main/install.ps1 -OutFile "$env:TEMP\bmfg-install.ps1"; powershell -ExecutionPolicy Bypass -File "$env:TEMP\bmfg-install.ps1"
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
$RepoName    = "BMFrameGen_release"
$PluginName  = "BMFrameGenCAD"
$BundleName  = "$PluginName.bundle"

# -- Functions ----------------------------------------------------------------

function Write-Banner {
    Write-Host ""
    Write-Host "  +--------------------------------------------+" -ForegroundColor Cyan
    Write-Host "  |     BMFrameGenCAD Plugin Installer          |" -ForegroundColor Cyan
    Write-Host "  |     beMatrix Frame Generator for AutoCAD    |" -ForegroundColor Cyan
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
        $versionKey = $_.PSPath
        Get-ChildItem $versionKey | ForEach-Object {
            $productKey = $_.PSPath
            $productName = (Get-ItemProperty $productKey -ErrorAction SilentlyContinue).ProductName
            $installPath = (Get-ItemProperty $productKey -ErrorAction SilentlyContinue).AcadLocation

            if ($productName -and $installPath -and (Test-Path $installPath)) {
                $versions += [PSCustomObject]@{
                    Name        = $productName
                    Version     = (Split-Path (Split-Path $versionKey -Leaf) -Leaf)
                    InstallPath = $installPath
                    RegistryKey = $productKey
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

    $appPlugins = Join-Path $env:APPDATA "Autodesk\ApplicationPlugins"
    $bundlePath = Join-Path $appPlugins $BundleName

    if (Test-Path $bundlePath) {
        Write-Host "  Removing previous installation..." -ForegroundColor Yellow
        Remove-Item $bundlePath -Recurse -Force
    }

    New-Item -Path $bundlePath -ItemType Directory -Force | Out-Null
    $contentsPath = Join-Path $bundlePath "Contents"
    New-Item -Path $contentsPath -ItemType Directory -Force | Out-Null

    $tempZip = Join-Path $env:TEMP "$PluginName-$TagName.zip"
    Write-Host "  Downloading $TagName..." -ForegroundColor Gray
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $tempZip -UseBasicParsing

    $tempExtract = Join-Path $env:TEMP "$PluginName-extract"
    if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
    Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force

    $dll = Get-ChildItem $tempExtract -Filter "$PluginName.dll" -Recurse | Select-Object -First 1
    if (-not $dll) {
        throw "Could not find $PluginName.dll in the release archive."
    }
    Copy-Item $dll.FullName -Destination $contentsPath

    $modelsSource = Get-ChildItem $tempExtract -Directory -Filter "Models" -Recurse | Select-Object -First 1
    if ($modelsSource) {
        $modelsDest = Join-Path $contentsPath "Models"
        Copy-Item $modelsSource.FullName -Destination $modelsDest -Recurse
        $modelCount = (Get-ChildItem $modelsDest -Filter "*.dwg").Count
        Write-Host "  Installed $modelCount model files." -ForegroundColor Gray
    }
    else {
        Write-Host "  WARNING: No Models folder found in release." -ForegroundColor Yellow
    }

    $settingsFile = Join-Path $env:APPDATA "bmframegen_settings.txt"
    $modelsPath = Join-Path $contentsPath "Models"
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

    Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
    Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue

    return $bundlePath
}

function Write-PackageContents {
    param(
        [string]$BundlePath,
        [string]$TagName
    )

    $xml = '<?xml version="1.0" encoding="utf-8"?>' + "`r`n"
    $xml += '<ApplicationPackage' + "`r`n"
    $xml += '    SchemaVersion="1.0"' + "`r`n"
    $xml += '    AppVersion="' + $TagName + '"' + "`r`n"
    $xml += '    ProductCode="{F7E8D3A1-5B2C-4D6E-9F0A-1B3C5D7E9F0A}"' + "`r`n"
    $xml += '    Name="BMFrameGenCAD"' + "`r`n"
    $xml += '    Description="beMatrix Frame Generator for AutoCAD"' + "`r`n"
    $xml += '    Author="EJL Design">' + "`r`n"
    $xml += '    <CompanyDetails Name="EJL Design" />' + "`r`n"
    $xml += '    <Components Description="BMFrameGenCAD">' + "`r`n"
    $xml += '        <RuntimeRequirements OS="Win64" Platform="AutoCAD" SeriesMin="R24.0" />' + "`r`n"
    $xml += '        <ComponentEntry AppName="BMFrameGenCAD"' + "`r`n"
    $xml += '                        Version="' + $TagName + '"' + "`r`n"
    $xml += '                        ModuleName="./Contents/' + $PluginName + '.dll"' + "`r`n"
    $xml += '                        AppType="Net"' + "`r`n"
    $xml += '                        LoadOnAppStartup="True" />' + "`r`n"
    $xml += '    </Components>' + "`r`n"
    $xml += '</ApplicationPackage>'

    $xmlPath = Join-Path $BundlePath 'PackageContents.xml'
    Set-Content $xmlPath $xml -Encoding UTF8
}

function Test-Installation {
    param([string]$BundlePath)

    $ok = $true
    $contentsPath = Join-Path $BundlePath "Contents"

    $dllPath = Join-Path $contentsPath "$PluginName.dll"
    if (Test-Path $dllPath) {
        Write-Host "  [OK] Plugin DLL installed" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Plugin DLL missing" -ForegroundColor Red
        $ok = $false
    }

    $xmlPath = Join-Path $BundlePath "PackageContents.xml"
    if (Test-Path $xmlPath) {
        Write-Host "  [OK] PackageContents.xml created" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] PackageContents.xml missing" -ForegroundColor Red
        $ok = $false
    }

    $modelsPath = Join-Path $contentsPath "Models"
    if ((Test-Path $modelsPath) -and (Get-ChildItem $modelsPath -Filter "*.dwg").Count -gt 0) {
        Write-Host "  [OK] Model library installed" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Model library missing or empty" -ForegroundColor Yellow
    }

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
    Write-Host "  No AutoCAD installations detected." -ForegroundColor Yellow
    Write-Host "  The plugin will be installed to the ApplicationPlugins folder." -ForegroundColor Yellow
    Write-Host "  It will activate when AutoCAD is installed and launched." -ForegroundColor Yellow
    Write-Host ""
}
else {
    Write-Host "  Found AutoCAD installations:" -ForegroundColor Green
    foreach ($v in $acadVersions) {
        Write-Host "    - $($v.Name)" -ForegroundColor White
    }
    Write-Host ""
}

$release = Get-LatestRelease
Write-Host "  Latest version: $($release.TagName)" -ForegroundColor Cyan
Write-Host ""

Write-Host "  Install BMFrameGenCAD $($release.TagName)? (Y/n) " -ForegroundColor White -NoNewline
$key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host $key.Character
if ($key.Character -eq 'n' -or $key.Character -eq 'N') {
    Write-Host "  Installation cancelled." -ForegroundColor Yellow
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

Write-Host ""

$bundlePath = Install-Plugin -DownloadUrl $release.DownloadUrl -TagName $release.TagName
Write-PackageContents -BundlePath $bundlePath -TagName $release.TagName

Write-Host ""
Write-Host "  Verifying installation..." -ForegroundColor Gray
$passed = Test-Installation -BundlePath $bundlePath

Write-Host ""
if ($passed) {
    Write-Host "  Installation complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Plugin location: $bundlePath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Next steps:" -ForegroundColor Cyan
    Write-Host "    1. Launch AutoCAD" -ForegroundColor White
    Write-Host "    2. The plugin loads automatically on startup" -ForegroundColor White
    Write-Host '    3. Type "BMFrameGen" in the command line to start' -ForegroundColor White
    Write-Host ""
}
else {
    Write-Host "  Installation completed with errors. Check the messages above." -ForegroundColor Red
}

Write-Host "  Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
