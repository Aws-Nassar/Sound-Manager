# Sound Manager

A Windows desktop UI prototype for managing audio input and output devices.

## Run

```powershell
python main.py
```

## Current Prototype

- Separate priority lists for outputs and inputs.
- Move devices up or down to control priority.
- Hide devices so they no longer clutter the visible picker.
- Search devices and optionally show hidden endpoints.
- Automation panel mockup for future Windows audio rules.

## Next Backend Step

The UI currently uses sample devices. To control real Windows endpoints, connect the device list and actions to a Windows audio backend such as Core Audio through `pycaw`, a small C# helper using NAudio, or Windows PowerShell/admin commands for device enable/disable behavior.
