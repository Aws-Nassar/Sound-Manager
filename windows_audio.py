from __future__ import annotations

import ctypes
import subprocess
from dataclasses import dataclass


@dataclass
class AudioEndpoint:
    device_id: str
    kind: str
    name: str
    subtitle: str
    status: str
    level: int
    hidden: bool
    can_disable: bool = True


class WindowsAudioBackend:
    def __init__(self) -> None:
        self.available = False
        self.error = ""

        try:
            from pycaw.pycaw import AudioUtilities, DEVICE_STATE, EDataFlow, ERole
            from comtypes import COMError
        except Exception as exc:
            self.error = f"Windows audio backend is unavailable: {exc}"
            return

        self.AudioUtilities = AudioUtilities
        self.DEVICE_STATE = DEVICE_STATE
        self.EDataFlow = EDataFlow
        self.ERole = ERole
        self.COMError = COMError
        self.available = True

    def list_devices(self, kind: str) -> list[AudioEndpoint]:
        if not self.available:
            return []

        flow = self.EDataFlow.eRender if kind == "output" else self.EDataFlow.eCapture
        default_id = self.default_device_id(kind)
        endpoints: list[AudioEndpoint] = []

        for device in self.AudioUtilities.GetAllDevices(flow.value, self.DEVICE_STATE.MASK_ALL.value):
            state_name = self._state_name(device)
            is_active = state_name == "Active"
            is_default = device.id == default_id
            status = "Default" if is_default and is_active else state_name
            hidden = not is_active
            name = device.FriendlyName or "Unknown audio endpoint"
            subtitle = self._subtitle(device)
            level = self._volume_percent(device) if is_active else 0

            endpoints.append(
                AudioEndpoint(
                    device_id=device.id,
                    kind=kind,
                    name=name,
                    subtitle=subtitle,
                    status=status,
                    level=level,
                    hidden=hidden,
                    can_disable=bool(device.id),
                )
            )

        endpoints.sort(key=lambda item: (item.status != "Default", item.hidden, item.name.lower()))
        return endpoints

    def default_device_id(self, kind: str) -> str:
        try:
            if kind == "output":
                return self.AudioUtilities.GetSpeakers().id

            enumerator = self.AudioUtilities.GetDeviceEnumerator()
            device = enumerator.GetDefaultAudioEndpoint(
                self.EDataFlow.eCapture.value,
                self.ERole.eMultimedia.value,
            )
            return device.GetId()
        except Exception:
            return ""

    def set_default(self, device_id: str) -> None:
        if not self.available or not device_id:
            return

        self.AudioUtilities.SetDefaultDevice(
            device_id,
            [self.ERole.eConsole, self.ERole.eMultimedia, self.ERole.eCommunications],
        )

    def set_volume(self, device_id: str, percent: int) -> None:
        if not self.available or not device_id:
            return

        percent = max(0, min(100, int(percent)))
        device = self.AudioUtilities.GetDeviceEnumerator().GetDevice(device_id)
        endpoint = self.AudioUtilities.CreateDevice(device)
        endpoint.EndpointVolume.SetMasterVolumeLevelScalar(percent / 100, None)

    def set_enabled(self, device_id: str, enabled: bool) -> str:
        if not device_id:
            return "Device ID is missing."

        return self.set_many_enabled([device_id], enabled)

    def set_many_enabled(self, device_ids: list[str], enabled: bool) -> str:
        device_ids = [device_id for device_id in device_ids if device_id]
        if not device_ids:
            return "No device changes needed."

        cmdlet = "Enable-PnpDevice" if enabled else "Disable-PnpDevice"
        instance_ids = [self._pnp_instance_id(device_id).replace("'", "''") for device_id in device_ids]
        quoted_ids = ", ".join(f"'{instance_id}'" for instance_id in instance_ids)
        script = f"$ids = @({quoted_ids}); foreach ($id in $ids) {{ {cmdlet} -InstanceId $id -Confirm:$false }}"

        if self._is_admin():
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                capture_output=True,
                text=True,
                check=False,
            )
            if completed.returncode == 0:
                return "Device change applied."
            return (completed.stderr or completed.stdout or "Windows rejected the device change.").strip()

        params = f'-NoProfile -ExecutionPolicy Bypass -Command "{script}"'
        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            "powershell.exe",
            params,
            None,
            0,
        )
        if result <= 32:
            return "Admin approval was cancelled or Windows blocked the request."
        return "Admin approval requested. Refresh after approving the Windows prompt."

    def _state_name(self, device) -> str:
        raw = str(getattr(device, "state", "Unknown")).split(".")[-1]
        return {
            "Active": "Active",
            "Disabled": "Disabled",
            "NotPresent": "Not present",
            "Unplugged": "Unplugged",
        }.get(raw, raw)

    def _subtitle(self, device) -> str:
        properties = getattr(device, "properties", {}) or {}
        interface_name = properties.get("{026E516E-B814-414B-83CD-856D6FEF4822} 2")
        bus = properties.get("{A45C254E-DF1C-4EFD-8020-67D146A850E0} 24")
        if interface_name and bus:
            return f"{interface_name} · {bus}"
        if interface_name:
            return str(interface_name)
        return "Windows audio endpoint"

    def _volume_percent(self, device) -> int:
        try:
            return max(0, min(100, round(device.EndpointVolume.GetMasterVolumeLevelScalar() * 100)))
        except Exception:
            return 0

    def _pnp_instance_id(self, device_id: str) -> str:
        return f"SWD\\MMDEVAPI\\{device_id}"

    def _is_admin(self) -> bool:
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False
