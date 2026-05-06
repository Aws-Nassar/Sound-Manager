# Sound Manager

A Windows desktop UI prototype for managing audio input and output devices.

## Run

```powershell
python main.py
```

## Current Prototype

- Separate priority lists for real Windows outputs and inputs.
- Move devices up or down to control app priority.
- The top visible device is applied as the real Windows default endpoint.
- Disable or restore devices through Windows PnP commands.
- Search devices and optionally show disabled, unplugged, and not-present endpoints.
- Automation panel mockup for future Windows audio rules.

## Windows Notes

Windows does not expose a supported setting to reorder the device list shown in Settings. This app keeps its own priority list and applies the first active device as the real Windows default.

Disabling or restoring an endpoint may show a Windows admin approval prompt. After approval, the app refreshes automatically.
