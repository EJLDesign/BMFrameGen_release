# TaylorFab Studio

TaylorFab Studio frame generator plugin for AutoCAD 2024+ (formerly BMFrameGenCAD). Generates full multi-wall modular framing towers, elevations, plan views, and composed sheets.

## Quick Install

Open **PowerShell** and paste:

```powershell
iex (irm https://raw.githubusercontent.com/EJLDesign/TaylorFab_Studio_release/main/install.ps1)
```

This runs the installer directly in memory — no temp script file, so antivirus scan locks on a downloaded file can't interrupt it.

The installer will:
- Detect your installed AutoCAD version(s)
- Download the latest release
- Install the plugin with automatic loading
- Install the current license file (imported silently on first load)
- Set up the model library

No admin rights required.

## Manual Install

1. Download the latest `.zip` from [Releases](https://github.com/EJLDesign/TaylorFab_Studio_release/releases/latest)
2. Create folder `%APPDATA%\Autodesk\ApplicationPlugins\TaylorFabStudio.bundle\Contents\`
3. Extract the DLL, license, Models, and assets folders into `Contents\`
4. Launch AutoCAD — the plugin loads automatically via the bundle

## Usage

1. Launch AutoCAD
2. Type `TFAB` in the command line
3. Build towers from the palette and edit them with the Tower and Wall editors

## Uninstall

Delete the folder:
```
%APPDATA%\Autodesk\ApplicationPlugins\TaylorFabStudio.bundle
```

## Requirements

- AutoCAD 2024 or later
- Windows 10/11
- .NET Framework 4.8 (included with Windows 10+)
- A valid license file (contact evanl@taylorinc.com)
