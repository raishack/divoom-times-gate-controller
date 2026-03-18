# Divom Times Gate Controller

Windows tray controller for Divoom Times Gate screens.

## What it does

- Keep 5 persistent screen slots (one media per screen).
- Re-send the current profile automatically every X minutes (default: 60).
- Re-send on app startup (useful when the Divoom reboots itself).
- Manual **Send now** and per-screen send.
- LAN discovery of Divoom devices + device picker.
- Live device health indicator (ONLINE/OFFLINE).
- GIF animated previews in UI.
- Hot theme switching (dark/light) without restart.
- In-app Auto-start toggle for Windows startup.

## Why this exists

Times Gate can lose current screen state after device restarts or unstable sessions. This app keeps a "desired state" profile and reapplies it automatically.

## Technical base

The image upload path is based on the working approach used by:

- https://github.com/adiastra/divoom-gaming-gate

Instead of relying on fragile official API flows, it uses the proven `Draw/SendHttpGif` payload strategy (JPEG frames in base64 with screen targeting and frame offsets).

## Run from source (Windows)

```powershell
python -m pip install -r requirements.txt
python app.py
```

## Build Windows executable

```powershell
./build_windows.ps1
```

Output:

- `dist/DivoomKeeper.exe`

## Files

- `app.py` - main app (tray + UI + scheduler + sender)
- `requirements.txt` - dependencies
- `build_windows.ps1` - build helper for PyInstaller
- `setup_startup.ps1` - optional startup shortcut helper
