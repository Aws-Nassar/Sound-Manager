# Sound Manager

A Windows desktop UI prototype for managing audio input and output devices.

## Run

```powershell
python main.py
```

## Build EXE

```powershell
.\tools\build_exe.ps1
```

The built app is written to:

```text
dist\SoundManager.exe
```

To remove generated build output:

```powershell
.\tools\build_exe.ps1 -Mode Clean
```

## Project Layout

```text
assets\sound_manager.ico
assets\sound_manager.png
tools\build_exe.ps1
tools\generate_icon.py
SoundManager.spec
config.json
main.py
windows_audio.py
```

## Current Features

- Separate priority lists for real Windows outputs and inputs.
- Move devices up or down to control app priority.
- The top visible device is applied as the real Windows default endpoint.
- Adjust input and output endpoint volume from inside the app.
- Apply profile presets from `config.json`.
- Use the protected `Default` profile to return to the normal active-device view.
- Save your current order, defaults, volumes, hidden sources, and disabled sources as a new profile.
- Rename or delete saved profiles. The `Default` profile is protected.
- Your current input/output sort order is saved in `sound_manager_state.json` and restored on the next run.
- Hide or unhide sources inside Sound Manager without changing Windows.
- Disable or enable sources through Windows PnP commands so they stop appearing as app input/output choices.
- The old automation card has been replaced with a source visibility panel so the app only shows controls that currently work.
- Search devices and optionally show disabled, unplugged, and not-present endpoints.
- Smooth scroll behavior while Windows devices refresh in the background.
- The window opens centered on the current screen and switches to a compact stacked layout on smaller window sizes.

## Windows Notes

Windows does not expose a supported setting to reorder the device list shown in Settings. This app keeps its own priority list and applies the first active device as the real Windows default.

Disabling or restoring an endpoint may show a Windows admin approval prompt. After approval, the app refreshes automatically.
