# BMFrameGenCAD

beMatrix Frame Generator plugin for AutoCAD 2024+. Generates modular framing systems on 3D solid faces and volumes.

## Quick Install

Open **PowerShell** and paste:

```powershell
irm https://raw.githubusercontent.com/EJLDesign/BMFrameGen_release/main/install.ps1 -OutFile "$env:TEMP\bmfg-install.ps1"; powershell -ExecutionPolicy Bypass -File "$env:TEMP\bmfg-install.ps1"
```

The installer will:
- Detect your installed AutoCAD version(s)
- Download the latest release
- Install the plugin with automatic loading
- Set up the model library

No admin rights required.

## Manual Install

1. Download the latest `.zip` from [Releases](https://github.com/EJLDesign/BMFrameGen_release/releases/latest)
2. Create folder `%APPDATA%\Autodesk\ApplicationPlugins\BMFrameGenCAD.bundle\Contents\`
3. Extract the DLL and Models folder into `Contents\`
4. Launch AutoCAD — the plugin loads automatically via the bundle

## Usage

1. Launch AutoCAD
2. Type `BMFrameGen` in the command line
3. Configure frame options in the palette
4. Select faces or volumes to frame

## Uninstall

Delete the folder:
```
%APPDATA%\Autodesk\ApplicationPlugins\BMFrameGenCAD.bundle
```

## Requirements

- AutoCAD 2024 or later
- Windows 10/11
- .NET Framework 4.8 (included with Windows 10+)
