# Sound Manager

Sound Manager is a Windows desktop app for controlling audio input and output devices from one clean interface. It lets you keep separate priority lists for outputs and inputs, hide devices you do not want to see, disable devices through Windows when needed, adjust device volume, and save different setups as profiles.

The app is written in Python with PyQt5 and talks to Windows Core Audio through `pycaw`.

## What It Does

- Shows real Windows playback and recording devices.
- Keeps separate priority lists for outputs and inputs.
- Moves devices up or down and uses the first active device as the Windows default.
- Changes device volume from inside the app.
- Hides sources only inside Sound Manager.
- Disables or enables sources through Windows so they stop appearing in Windows and other apps.
- Saves profiles with sort order, default devices, volumes, hidden sources, and disabled sources.
- Protects the `Default` profile from rename/delete.
- Opens centered on the screen and adapts to smaller window sizes.
- Shows a source visibility panel for hidden, disabled, and unavailable devices.

## Important Warning

Use **Hide** when you only want to clean up the Sound Manager list.

Use **Disable/Off** carefully. Disabling a source changes the real Windows device state through PnP commands. A disabled microphone, speaker, virtual cable, or headset can disappear from Windows Settings and from apps that choose input/output devices.

If Windows asks for administrator approval, that is expected. Approving it lets the app disable or enable the selected audio endpoint. If you disable the wrong device, open Sound Manager, turn on `Show all`, then use **On** to enable it again.

## Requirements

- Windows 10 or Windows 11
- Python 3.10 or newer
- PowerShell for the build script
- The packages in `requirements.txt`

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Main dependencies:

1. `PyQt5` - desktop GUI
2. `pycaw` - Windows Core Audio access
3. `comtypes` - COM support used by `pycaw`
4. `pyinstaller` - EXE builder

## Run From Source

From the project folder:

```powershell
python main.py
```

You can also run:

```powershell
.\run_sound_manager.bat
```

## Build The EXE

Install dependencies first, then run:

```powershell
.\tools\build_exe.ps1
```

The generated app will be:

```text
dist\SoundManager.exe
```

To remove generated build output:

```powershell
.\tools\build_exe.ps1 -Mode Clean
```

## How To Use

1. Choose `Outputs` or `Inputs` from the left side.
2. Move devices with the up/down buttons to set priority.
3. The first active visible device becomes the Windows default for that type.
4. Use the volume slider on each device to change its Windows endpoint volume.
5. Use **Hide** to remove a source from this app only.
6. Use **Off** to disable a source in Windows.
7. Turn on `Show all` to see hidden, disabled, unplugged, or unavailable sources.
8. Use **On** or **Show** to bring devices back.
9. Save the current setup as a profile when you want to reuse it later.

Profile switching updates the app view immediately, then applies the real Windows changes right after that. During that moment the status text shows that the profile is being applied to Windows.

## Profiles

Profiles save:

- output order
- input order
- current default output
- current default input
- per-device volume
- hidden devices
- disabled devices

The built-in profiles live in `config.json`. Your local profile edits are saved in `sound_manager_state.json`, which is intentionally ignored by Git because it contains machine-specific device IDs.

The `Default` profile is protected. You can use it normally, but the app does not allow renaming or deleting it.

## Project Layout

```text
assets/
  sound_manager.ico
  sound_manager.png

tools/
  build_exe.ps1
  generate_icon.py

config.json
main.py
windows_audio.py
requirements.txt
run_sound_manager.bat
SoundManager.spec
README.md
```

## Git Notes

Generated and local files should not be committed:

- `build/`
- `dist/`
- `__pycache__/`
- `.venv/`
- `sound_manager_state.json`
- logs, coverage files, and editor folders

If any ignored generated file was already committed before `.gitignore` was fixed, remove it from Git tracking without deleting your local copy:

```powershell
git rm -r --cached build dist __pycache__ sound_manager_state.json
git commit -m "Stop tracking generated files"
```

If Git says one of those paths does not exist, remove that path from the command and run it again.

## Windows Notes

Windows does not provide a normal supported setting to reorder the device list shown in Settings. Sound Manager keeps its own priority order and applies the first active device as the real Windows default.

Some Windows apps cache device lists. If an app does not immediately notice a device was enabled or disabled, restart that app or reopen its audio settings.
