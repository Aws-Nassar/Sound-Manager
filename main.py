import json
import sys
from dataclasses import dataclass
from pathlib import Path

from PyQt5.QtCore import Qt, QSize, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QInputDialog,
    QPushButton,
    QSizePolicy,
    QSlider,
    QVBoxLayout,
    QWidget,
)

from windows_audio import WindowsAudioBackend


@dataclass
class AudioDevice:
    device_id: str
    kind: str
    name: str
    subtitle: str
    status: str
    level: int
    hidden: bool = False
    can_disable: bool = True


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
CONFIG_PATH = APP_DIR / "config.json"
BUNDLED_CONFIG_PATH = RESOURCE_DIR / "config.json"
STATE_PATH = APP_DIR / "sound_manager_state.json"
APP_ICON_PATH = RESOURCE_DIR / "assets" / "sound_manager.ico"

DEFAULT_PROFILES = {
    "Work Mode": {
        "description": "Headset first for meetings, speakers as backup.",
        "output_keywords": ["headset", "arctis", "speakers", "realtek"],
        "input_keywords": ["microphone", "headset", "arctis", "webcam"],
        "output_volume": 72,
        "input_volume": 68,
    },
    "Gaming": {
        "description": "Prioritize headset and VR/game endpoints.",
        "output_keywords": ["headset", "arctis", "quest", "nvidia", "hdmi"],
        "input_keywords": ["headset", "microphone", "arctis"],
        "output_volume": 82,
        "input_volume": 74,
    },
    "Recording": {
        "description": "Studio interfaces first with calmer monitoring.",
        "output_keywords": ["focusrite", "scarlett", "studio", "speakers"],
        "input_keywords": ["scarlett", "focusrite", "microphone", "mic"],
        "output_volume": 60,
        "input_volume": 80,
    },
    "Clean": {
        "description": "Prefer built-in stable devices and keep virtual endpoints lower.",
        "output_keywords": ["speakers", "realtek", "headphones"],
        "input_keywords": ["microphone", "array", "webcam"],
        "output_volume": 65,
        "input_volume": 60,
    },
}


OUTPUT_DEVICES = [
    AudioDevice("out-headset", "output", "SteelSeries Arctis 7", "USB wireless headset", "Default", 84),
    AudioDevice("out-speakers", "output", "Realtek Speakers", "Built-in laptop speakers", "Ready", 52),
    AudioDevice("out-studio", "output", "Focusrite Scarlett", "Studio monitor output", "Ready", 73),
    AudioDevice("out-hdmi", "output", "NVIDIA HDMI Audio", "Monitor audio endpoint", "Hidden", 0, True),
    AudioDevice("out-vr", "output", "Quest Link Audio", "Virtual headset output", "Hidden", 0, True),
]

INPUT_DEVICES = [
    AudioDevice("in-headset", "input", "Arctis 7 Microphone", "Noise cancelling headset mic", "Default", 67),
    AudioDevice("in-studio", "input", "Scarlett Solo Input", "XLR microphone interface", "Ready", 41),
    AudioDevice("in-webcam", "input", "Logitech Webcam Mic", "Camera microphone", "Ready", 20),
    AudioDevice("in-array", "input", "Realtek Mic Array", "Laptop microphone array", "Hidden", 0, True),
]


class DeviceCard(QFrame):
    hide_requested = pyqtSignal(str)
    move_requested = pyqtSignal(str, int)
    volume_requested = pyqtSignal(str, int)

    def __init__(self, device: AudioDevice, priority: int):
        super().__init__()
        self.device = device
        self.setObjectName("deviceCard")
        self.setProperty("hiddenDevice", device.hidden)

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 12, 12, 12)
        root.setSpacing(12)

        handle = QLabel("::")
        handle.setObjectName("dragHandle")
        handle.setAlignment(Qt.AlignCenter)
        handle.setFixedWidth(18)

        badge = QLabel(f"{priority:02}")
        badge.setObjectName("priorityBadge")
        badge.setAlignment(Qt.AlignCenter)
        badge.setFixedSize(42, 42)

        copy = QVBoxLayout()
        copy.setSpacing(7)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        title = QLabel(device.name)
        title.setObjectName("deviceTitle")
        title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        status = QLabel("Disabled" if device.hidden else device.status)
        status.setObjectName("statusPill")
        status.setProperty("tone", "muted" if device.hidden else ("default" if device.status == "Default" else "ready"))

        title_row.addWidget(title)
        title_row.addWidget(status)

        subtitle = QLabel(device.subtitle)
        subtitle.setObjectName("deviceSubtitle")

        meter = QFrame()
        meter.setObjectName("meter")
        meter.setFixedHeight(7)
        meter_fill = QFrame(meter)
        meter_fill.setObjectName("meterFill")
        meter_fill.setGeometry(0, 0, max(4, int(device.level * 1.8)), 7)
        if device.hidden:
            meter_fill.setFixedWidth(0)

        copy.addLayout(title_row)
        copy.addWidget(subtitle)
        copy.addWidget(meter)

        volume_row = QHBoxLayout()
        volume_row.setSpacing(10)
        volume_label = QLabel(f"{device.level}%")
        volume_label.setObjectName("volumeLabel")
        volume_label.setFixedWidth(42)
        volume_slider = QSlider(Qt.Horizontal)
        volume_slider.setObjectName("volumeSlider")
        volume_slider.setRange(0, 100)
        volume_slider.setValue(device.level)
        volume_slider.setEnabled(not device.hidden)
        volume_slider.setCursor(Qt.PointingHandCursor)

        def update_volume(value: int) -> None:
            volume_label.setText(f"{value}%")
            self.volume_requested.emit(device.device_id, value)

        volume_slider.valueChanged.connect(update_volume)
        volume_row.addWidget(volume_slider, 1)
        volume_row.addWidget(volume_label)
        copy.addLayout(volume_row)

        actions = QHBoxLayout()
        actions.setSpacing(6)

        up = self._icon_button("↑", "Move higher priority")
        down = self._icon_button("↓", "Move lower priority")
        hide = self._action_button("Enable" if device.hidden else "Disable", "Enable this source" if device.hidden else "Disable this source")
        hide.setObjectName("restoreActionButton" if device.hidden else "dangerActionButton")

        up.clicked.connect(lambda: self.move_requested.emit(device.device_id, -1))
        down.clicked.connect(lambda: self.move_requested.emit(device.device_id, 1))
        hide.clicked.connect(lambda: self.hide_requested.emit(device.device_id))

        actions.addWidget(up)
        actions.addWidget(down)
        actions.addWidget(hide)

        root.addWidget(handle)
        root.addWidget(badge)
        root.addLayout(copy, 1)
        root.addLayout(actions)

    def _icon_button(self, text: str, tooltip: str) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("iconButton")
        button.setToolTip(tooltip)
        button.setFixedSize(34, 34)
        button.setCursor(Qt.PointingHandCursor)
        return button

    def _action_button(self, text: str, tooltip: str) -> QPushButton:
        button = QPushButton(text)
        button.setToolTip(tooltip)
        button.setFixedSize(76, 34)
        button.setCursor(Qt.PointingHandCursor)
        return button


class StatCard(QFrame):
    def __init__(self, label: str, value: str, tone: str):
        super().__init__()
        self.setObjectName("statCard")
        self.setProperty("tone", tone)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)

        self.value = QLabel(value)
        self.value.setObjectName("statValue")
        text = QLabel(label)
        text.setObjectName("statLabel")

        layout.addWidget(self.value)
        layout.addWidget(text)


class SoundManagerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.backend = WindowsAudioBackend()
        self.outputs = []
        self.inputs = []
        self.state = self._load_state()
        self.priority_order = self.state.get("priority_order", {"output": [], "input": []})
        self.priority_order.setdefault("output", [])
        self.priority_order.setdefault("input", [])
        self.disabled_device_ids = set(self.state.get("disabled_devices", []))
        self.profiles = self._load_profiles()
        self.profile_buttons = {}
        self.profile_layout = None
        self.disable_refresh_attempts = 0
        self.volume_update_timer = QTimer(self)
        self.volume_update_timer.setSingleShot(True)
        self.pending_volume_change = None
        self.show_hidden = False
        self.current_kind = "output"

        self.setWindowTitle("Sound Manager")
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))
        self.setMinimumSize(1120, 760)
        self.resize(1220, 820)

        self.setCentralWidget(self._build_ui())
        self._apply_style()
        self.reload_devices()
        self.refresh()

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(2500)
        self.poll_timer.timeout.connect(lambda: self.reload_devices(refresh=True))
        self.poll_timer.start()
        self.volume_update_timer.timeout.connect(self._flush_volume_change)

    def _build_ui(self) -> QWidget:
        shell = QWidget()
        shell.setObjectName("shell")
        root = QHBoxLayout(shell)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())
        root.addWidget(self._build_workspace(), 1)
        return shell

    def _build_sidebar(self) -> QWidget:
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(258)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(22, 24, 18, 22)
        layout.setSpacing(18)

        logo_row = QHBoxLayout()
        logo = QLabel("SM")
        logo.setObjectName("logo")
        logo.setAlignment(Qt.AlignCenter)
        logo.setFixedSize(48, 48)

        brand = QVBoxLayout()
        title = QLabel("Sound Manager")
        title.setObjectName("brandTitle")
        subtitle = QLabel("Windows audio control")
        subtitle.setObjectName("brandSubtitle")
        brand.addWidget(title)
        brand.addWidget(subtitle)

        logo_row.addWidget(logo)
        logo_row.addLayout(brand)
        layout.addLayout(logo_row)

        self.output_nav = self._nav_button("Outputs", "output")
        self.input_nav = self._nav_button("Inputs", "input")
        for button in (self.output_nav, self.input_nav):
            layout.addWidget(button)

        layout.addSpacing(8)
        divider = QFrame()
        divider.setObjectName("divider")
        divider.setFixedHeight(1)
        layout.addWidget(divider)

        rules_title = QLabel("Profiles")
        rules_title.setObjectName("sectionLabel")
        layout.addWidget(rules_title)

        self.profile_layout = QVBoxLayout()
        self.profile_layout.setSpacing(8)
        layout.addLayout(self.profile_layout)
        self._refresh_profile_buttons()

        save_profile = QPushButton("+ Save current profile")
        save_profile.setObjectName("primarySideButton")
        save_profile.setCursor(Qt.PointingHandCursor)
        save_profile.clicked.connect(self.create_profile_from_current)
        layout.addWidget(save_profile)

        layout.addStretch()

        hint = QFrame()
        hint.setObjectName("sidebarHint")
        hint_layout = QVBoxLayout(hint)
        hint_layout.setContentsMargins(14, 14, 14, 14)
        hint_layout.setSpacing(6)
        hint_title = QLabel("Quiet picker")
        hint_title.setObjectName("hintTitle")
        hint_body = QLabel("Disabled sources stay out of Windows and app input/output menus until you enable them again.")
        hint_body.setObjectName("hintBody")
        hint_body.setWordWrap(True)
        hint_layout.addWidget(hint_title)
        hint_layout.addWidget(hint_body)
        layout.addWidget(hint)

        return sidebar

    def _nav_button(self, label: str, kind: str) -> QPushButton:
        button = QPushButton(label)
        button.setObjectName("navButton")
        button.setCheckable(True)
        button.setCursor(Qt.PointingHandCursor)
        button.clicked.connect(lambda: self.set_kind(kind))
        return button

    def _build_workspace(self) -> QWidget:
        workspace = QWidget()
        workspace.setObjectName("workspace")
        layout = QVBoxLayout(workspace)
        layout.setContentsMargins(34, 30, 34, 30)
        layout.setSpacing(22)

        header = QHBoxLayout()
        header.setSpacing(18)

        headline = QVBoxLayout()
        eyebrow = QLabel("Device priority")
        eyebrow.setObjectName("eyebrow")
        title = QLabel("Control what Windows shows first")
        title.setObjectName("pageTitle")
        headline.addWidget(eyebrow)
        headline.addWidget(title)

        self.search = QLineEdit()
        self.search.setObjectName("search")
        self.search.setPlaceholderText("Search devices")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.refresh)

        self.hidden_toggle = QCheckBox("Show disabled")
        self.hidden_toggle.setObjectName("showHidden")
        self.hidden_toggle.stateChanged.connect(self._toggle_hidden)

        header.addLayout(headline, 1)
        header.addWidget(self.search)
        header.addWidget(self.hidden_toggle)
        layout.addLayout(header)

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        stats = QGridLayout()
        stats.setHorizontalSpacing(14)
        self.visible_stat = StatCard("Visible devices", "0", "green")
        self.hidden_stat = StatCard("Disabled sources", "0", "amber")
        self.default_stat = StatCard("Current default", "None", "coral")
        stats.addWidget(self.visible_stat, 0, 0)
        stats.addWidget(self.hidden_stat, 0, 1)
        stats.addWidget(self.default_stat, 0, 2)
        layout.addLayout(stats)

        body = QHBoxLayout()
        body.setSpacing(18)

        list_panel = QFrame()
        list_panel.setObjectName("panel")
        list_layout = QVBoxLayout(list_panel)
        list_layout.setContentsMargins(18, 18, 18, 18)
        list_layout.setSpacing(14)

        panel_top = QHBoxLayout()
        self.list_title = QLabel("Outputs")
        self.list_title.setObjectName("panelTitle")
        sort_button = QPushButton("Sort active first")
        sort_button.setObjectName("primaryButton")
        sort_button.setCursor(Qt.PointingHandCursor)
        sort_button.clicked.connect(self.sort_current)
        panel_top.addWidget(self.list_title, 1)
        panel_top.addWidget(sort_button)

        self.list_widget = QListWidget()
        self.list_widget.setObjectName("deviceList")
        self.list_widget.setSpacing(10)
        self.list_widget.setDragDropMode(QListWidget.InternalMove)
        self.list_widget.setDefaultDropAction(Qt.MoveAction)
        self.list_widget.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.list_widget.model().rowsMoved.connect(self._sync_from_list_order)

        list_layout.addLayout(panel_top)
        list_layout.addWidget(self.list_widget, 1)

        body.addWidget(list_panel, 1)
        body.addWidget(self._build_rules_panel())
        layout.addLayout(body, 1)

        return workspace

    def _build_rules_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("rulesPanel")
        panel.setFixedWidth(330)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(15)

        title = QLabel("Automation")
        title.setObjectName("panelTitle")
        layout.addWidget(title)

        for text, detail, checked in (
            ("Hide unplugged devices", "Remove disconnected endpoints from menus.", True),
            ("Prefer headset for calls", "Promote headset input and output when meetings start.", True),
            ("Lock studio output", "Keep music apps on your interface.", False),
        ):
            layout.addWidget(self._rule_toggle(text, detail, checked))

        layout.addSpacing(8)
        activity_label = QLabel("Recent activity")
        activity_label.setObjectName("sectionLabel")
        layout.addWidget(activity_label)

        for time, text in (
            ("09:24", "Headset promoted for Work Mode."),
            ("09:10", "NVIDIA HDMI Audio hidden."),
            ("08:57", "Realtek Speakers set as fallback."),
        ):
            row = QFrame()
            row.setObjectName("activityRow")
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 4, 0, 4)
            row_layout.setSpacing(12)
            stamp = QLabel(time)
            stamp.setObjectName("activityTime")
            message = QLabel(text)
            message.setObjectName("activityText")
            message.setWordWrap(True)
            row_layout.addWidget(stamp)
            row_layout.addWidget(message, 1)
            layout.addWidget(row)

        layout.addStretch()
        return panel

    def _rule_toggle(self, title: str, detail: str, checked: bool) -> QWidget:
        row = QFrame()
        row.setObjectName("ruleRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(13, 12, 13, 12)
        layout.setSpacing(10)

        copy = QVBoxLayout()
        copy.setSpacing(3)
        name = QLabel(title)
        name.setObjectName("ruleTitle")
        desc = QLabel(detail)
        desc.setObjectName("ruleDetail")
        desc.setWordWrap(True)
        copy.addWidget(name)
        copy.addWidget(desc)

        toggle = QCheckBox()
        toggle.setChecked(checked)
        toggle.setCursor(Qt.PointingHandCursor)

        layout.addLayout(copy, 1)
        layout.addWidget(toggle)
        return row

    def _refresh_profile_buttons(self) -> None:
        if self.profile_layout is None:
            return

        while self.profile_layout.count():
            item = self.profile_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.profile_buttons = {}
        for text in self.profiles:
            chip = QPushButton(text)
            chip.setObjectName("profileButton")
            chip.setCursor(Qt.PointingHandCursor)
            chip.setToolTip(self.profiles[text].get("description", "Apply this audio profile."))
            chip.clicked.connect(lambda checked=False, name=text: self.apply_profile(name))
            self.profile_buttons[text] = chip
            self.profile_layout.addWidget(chip)

    def _load_state(self) -> dict:
        if not STATE_PATH.exists():
            return {}

        try:
            data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _save_state(self) -> None:
        self.state["priority_order"] = self.priority_order
        self.state["disabled_devices"] = sorted(self.disabled_device_ids)
        self.state.setdefault("profiles", {})

        try:
            STATE_PATH.write_text(json.dumps(self.state, indent=2), encoding="utf-8")
        except OSError as exc:
            if hasattr(self, "status_label"):
                self.status_label.setText(f"Could not save app state: {exc}")

    def _load_profiles(self) -> dict:
        path = CONFIG_PATH if CONFIG_PATH.exists() else BUNDLED_CONFIG_PATH
        profiles = dict(DEFAULT_PROFILES)
        if not path.exists():
            profiles.update(self.state.get("profiles", {}))
            return profiles

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            profiles.update(data.get("profiles", {}))
            profiles.update(self.state.get("profiles", {}))
            return profiles
        except (OSError, json.JSONDecodeError):
            profiles.update(self.state.get("profiles", {}))
            return profiles

    def set_kind(self, kind: str) -> None:
        self.current_kind = kind
        self.output_nav.setChecked(kind == "output")
        self.input_nav.setChecked(kind == "input")
        self.refresh()

    def _toggle_hidden(self) -> None:
        self.show_hidden = self.hidden_toggle.isChecked()
        self.refresh()

    def current_devices(self) -> list[AudioDevice]:
        return self.outputs if self.current_kind == "output" else self.inputs

    def set_current_devices(self, devices: list[AudioDevice]) -> None:
        if self.current_kind == "output":
            self.outputs = devices
        else:
            self.inputs = devices
        self.priority_order[self.current_kind] = [device.device_id for device in devices]
        self._save_state()

    def reload_devices(self, refresh: bool = False) -> None:
        if refresh and QApplication.mouseButtons() != Qt.NoButton:
            return

        if self.backend.available:
            self.outputs = self._apply_disabled_overrides(self._apply_saved_order("output", self.backend.list_devices("output")))
            self.inputs = self._apply_disabled_overrides(self._apply_saved_order("input", self.backend.list_devices("input")))
            if hasattr(self, "status_label") and not self.status_label.text():
                self.status_label.setText("Connected to Windows Core Audio. Top priority changes update the real Windows default device.")
        else:
            self.outputs = self._apply_disabled_overrides(self._apply_saved_order("output", list(OUTPUT_DEVICES)))
            self.inputs = self._apply_disabled_overrides(self._apply_saved_order("input", list(INPUT_DEVICES)))
            if hasattr(self, "status_label"):
                self.status_label.setText(self.backend.error or "Using sample devices because the Windows audio backend is unavailable.")

        if refresh:
            self.refresh()

    def _apply_saved_order(self, kind: str, devices: list[AudioDevice]) -> list[AudioDevice]:
        saved = self.priority_order[kind]
        if not saved:
            self.priority_order[kind] = [device.device_id for device in devices]
            return devices

        lookup = {device.device_id: device for device in devices}
        ordered = [lookup[device_id] for device_id in saved if device_id in lookup]
        ordered_ids = {device.device_id for device in ordered}
        ordered.extend(device for device in devices if device.device_id not in ordered_ids)
        self.priority_order[kind] = [device.device_id for device in ordered]
        return ordered

    def _apply_disabled_overrides(self, devices: list[AudioDevice]) -> list[AudioDevice]:
        for device in devices:
            if device.device_id in self.disabled_device_ids:
                device.hidden = True
                device.status = "Disabled"
                device.level = 0
        return devices

    def refresh(self) -> None:
        if not hasattr(self, "list_widget"):
            return

        self.output_nav.setChecked(self.current_kind == "output")
        self.input_nav.setChecked(self.current_kind == "input")
        self.list_title.setText("Outputs" if self.current_kind == "output" else "Inputs")

        query = self.search.text().strip().lower()
        devices = self.current_devices()
        visible = [
            device
            for device in devices
            if (self.show_hidden or not device.hidden)
            and (not query or query in device.name.lower() or query in device.subtitle.lower())
        ]

        scroll_value = self.list_widget.verticalScrollBar().value()
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for priority, device in enumerate(visible, start=1):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, device.device_id)
            card = DeviceCard(device, priority)
            card.hide_requested.connect(self.toggle_device_hidden)
            card.move_requested.connect(self.move_device)
            card.volume_requested.connect(self.set_device_volume)
            item.setSizeHint(QSize(100, 126))
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, card)
        self.list_widget.blockSignals(False)
        self.list_widget.verticalScrollBar().setValue(scroll_value)

        self._update_stats()

    def _update_stats(self) -> None:
        all_current = self.current_devices()
        visible = [device for device in all_current if not device.hidden]
        hidden = [device for device in all_current if device.hidden]
        default = next((device.name for device in all_current if device.status == "Default" and not device.hidden), "None")

        self.visible_stat.value.setText(str(len(visible)))
        self.hidden_stat.value.setText(str(len(hidden)))
        self.default_stat.value.setText(default)

    def toggle_device_hidden(self, device_id: str) -> None:
        device = next((item for item in self.current_devices() if item.device_id == device_id), None)
        if not device:
            return

        enable_device = device.hidden

        if self.backend.available:
            message = self.backend.set_enabled(device_id, enable_device)
            self.status_label.setText(message)
            if "cancelled" not in message.lower() and "blocked" not in message.lower():
                self._set_device_hidden_locally(device_id, not enable_device)
                self.disable_refresh_attempts = 8
                QTimer.singleShot(1000, self._refresh_after_disable_action)
            return

        self._set_device_hidden_locally(device_id, not enable_device)

    def _set_device_hidden_locally(self, device_id: str, hidden: bool) -> None:
        if hidden:
            self.disabled_device_ids.add(device_id)
        else:
            self.disabled_device_ids.discard(device_id)

        for collection in (self.outputs, self.inputs):
            for device in collection:
                if device.device_id == device_id:
                    device.hidden = hidden
                    device.status = "Disabled" if hidden else "Active"
                    break
        self._save_state()
        self.refresh()

    def _refresh_after_disable_action(self) -> None:
        self.reload_devices(refresh=True)
        self.disable_refresh_attempts -= 1
        if self.disable_refresh_attempts > 0:
            QTimer.singleShot(1200, self._refresh_after_disable_action)

    def set_device_volume(self, device_id: str, value: int) -> None:
        for device in self.current_devices():
            if device.device_id == device_id:
                device.level = value
                break

        if not self.backend.available:
            return

        self.pending_volume_change = (device_id, value)
        self.volume_update_timer.start(120)

    def _flush_volume_change(self) -> None:
        if not self.pending_volume_change:
            return

        device_id, value = self.pending_volume_change
        self.pending_volume_change = None
        try:
            self.backend.set_volume(device_id, value)
            self.status_label.setText(f"Volume set to {value}%.")
        except Exception as exc:
            self.status_label.setText(f"Windows rejected the volume change: {exc}")

    def apply_profile(self, profile_name: str) -> None:
        profile = self.profiles.get(profile_name)
        if not profile:
            return

        self.outputs = self._order_for_profile("output", self.outputs, profile)
        self.inputs = self._order_for_profile("input", self.inputs, profile)
        self.priority_order["output"] = [device.device_id for device in self.outputs]
        self.priority_order["input"] = [device.device_id for device in self.inputs]

        self._apply_profile_disabled_sources(profile)
        self._apply_profile_default("output", profile)
        self._apply_profile_default("input", profile)
        self._apply_profile_volumes(profile)
        self._save_state()
        self.status_label.setText(f"{profile_name} profile applied. {profile.get('description', '')}".strip())
        self.refresh()

    def create_profile_from_current(self) -> None:
        name, accepted = QInputDialog.getText(self, "Save profile", "Profile name:")
        name = name.strip()
        if not accepted or not name:
            return

        profile = self._snapshot_current_profile(name)
        self.profiles[name] = profile
        self.state.setdefault("profiles", {})[name] = profile
        self._save_state()
        self._refresh_profile_buttons()
        self.status_label.setText(f"{name} profile saved with current order, volumes, defaults, and disabled sources.")

    def _snapshot_current_profile(self, name: str) -> dict:
        disabled = set(self.disabled_device_ids)
        disabled.update(device.device_id for device in self.outputs + self.inputs if device.hidden)
        volumes = {device.device_id: device.level for device in self.outputs + self.inputs}
        output_default = next((device.device_id for device in self.outputs if not device.hidden), "")
        input_default = next((device.device_id for device in self.inputs if not device.hidden), "")

        return {
            "description": f"Saved snapshot for {name}.",
            "output_order": [device.device_id for device in self.outputs],
            "input_order": [device.device_id for device in self.inputs],
            "output_default": output_default,
            "input_default": input_default,
            "volumes": volumes,
            "disabled_devices": sorted(disabled),
        }

    def _order_for_profile(self, kind: str, devices: list[AudioDevice], profile: dict) -> list[AudioDevice]:
        order = profile.get(f"{kind}_order")
        if order:
            lookup = {device.device_id: device for device in devices}
            ordered = [lookup[device_id] for device_id in order if device_id in lookup]
            ordered_ids = {device.device_id for device in ordered}
            ordered.extend(device for device in devices if device.device_id not in ordered_ids)
            return ordered

        return self._rank_for_profile(devices, profile.get(f"{kind}_keywords", []))

    def _apply_profile_disabled_sources(self, profile: dict) -> None:
        if "disabled_devices" not in profile:
            return

        previous = set(self.disabled_device_ids)
        target = set(profile.get("disabled_devices", []))
        all_ids = {device.device_id for device in self.outputs + self.inputs}
        to_disable = sorted((target - previous) & all_ids)
        to_enable = sorted((previous - target) & all_ids)

        self.disabled_device_ids = target
        for device in self.outputs + self.inputs:
            if device.device_id in target:
                device.hidden = True
                device.status = "Disabled"
            elif device.device_id in previous:
                device.hidden = False
                device.status = "Active"

        if self.backend.available:
            messages = []
            if to_disable:
                messages.append(self.backend.set_many_enabled(to_disable, False))
            if to_enable:
                messages.append(self.backend.set_many_enabled(to_enable, True))
            if messages:
                self.status_label.setText(" ".join(messages))
                self.disable_refresh_attempts = 8
                QTimer.singleShot(1000, self._refresh_after_disable_action)

    def _apply_profile_volumes(self, profile: dict) -> None:
        volumes = profile.get("volumes", {})
        if not volumes:
            return

        for device in self.outputs + self.inputs:
            if device.device_id not in volumes or device.hidden:
                continue

            value = int(volumes[device.device_id])
            device.level = value
            if self.backend.available:
                try:
                    self.backend.set_volume(device.device_id, value)
                except Exception:
                    pass

    def _rank_for_profile(self, devices: list[AudioDevice], keywords: list[str]) -> list[AudioDevice]:
        normalized = [keyword.lower() for keyword in keywords]

        def score(device: AudioDevice) -> tuple[int, int, str]:
            haystack = f"{device.name} {device.subtitle}".lower()
            match_index = next((index for index, keyword in enumerate(normalized) if keyword in haystack), len(normalized))
            return (device.hidden, match_index, device.name.lower())

        return sorted(devices, key=score)

    def _apply_profile_default(self, kind: str, profile: dict) -> None:
        devices = self.outputs if kind == "output" else self.inputs
        preferred_id = profile.get(f"{kind}_default", "")
        first_active = next((device for device in devices if device.device_id == preferred_id and not device.hidden), None)
        if first_active is None:
            first_active = next((device for device in devices if not device.hidden), None)
        if not first_active or not self.backend.available:
            return

        try:
            self.backend.set_default(first_active.device_id)
            volume = profile.get(f"{kind}_volume")
            if volume is not None:
                first_active.level = int(volume)
                self.backend.set_volume(first_active.device_id, int(volume))
        except Exception as exc:
            self.status_label.setText(f"Could not fully apply {kind} profile: {exc}")

    def move_device(self, device_id: str, direction: int) -> None:
        devices = self.current_devices()
        index = next((i for i, device in enumerate(devices) if device.device_id == device_id), -1)
        new_index = index + direction
        if index < 0 or new_index < 0 or new_index >= len(devices):
            return
        devices[index], devices[new_index] = devices[new_index], devices[index]
        self.set_current_devices(devices)
        self._apply_default_from_priority()
        self.refresh()

    def sort_current(self) -> None:
        devices = sorted(
            self.current_devices(),
            key=lambda device: (device.hidden, device.status != "Default", -device.level),
        )
        self.set_current_devices(devices)
        self._apply_default_from_priority()
        self.refresh()

    def _sync_from_list_order(self) -> None:
        visible_ids = [
            self.list_widget.item(index).data(Qt.UserRole)
            for index in range(self.list_widget.count())
        ]
        devices = self.current_devices()
        lookup = {device.device_id: device for device in devices}
        visible_devices = [lookup[device_id] for device_id in visible_ids if device_id in lookup]
        hidden_outside_filter = [device for device in devices if device.device_id not in visible_ids]
        self.set_current_devices(visible_devices + hidden_outside_filter)
        self._apply_default_from_priority()
        self.refresh()

    def _apply_default_from_priority(self) -> None:
        if not self.backend.available:
            return

        first_active = next((device for device in self.current_devices() if not device.hidden), None)
        if not first_active:
            self.status_label.setText("No active device is available to set as the Windows default.")
            return

        try:
            self.backend.set_default(first_active.device_id)
            self.status_label.setText(f"Windows default {self.current_kind} set to {first_active.name}.")
            QTimer.singleShot(500, lambda: self.reload_devices(refresh=True))
        except Exception as exc:
            self.status_label.setText(f"Windows rejected the default-device change: {exc}")

    def _apply_style(self) -> None:
        QApplication.instance().setFont(QFont("Segoe UI", 10))
        self.setStyleSheet(
            """
            QWidget#shell {
                background: #101216;
                color: #eef2f5;
            }

            QFrame#sidebar {
                background: #171a20;
                border-right: 1px solid #252a32;
            }

            QLabel#logo {
                background: #2dd4bf;
                color: #06120f;
                border-radius: 14px;
                font-weight: 800;
                font-size: 15px;
            }

            QLabel#brandTitle,
            QLabel#panelTitle {
                color: #f7fafc;
                font-size: 19px;
                font-weight: 800;
            }

            QLabel#brandSubtitle,
            QLabel#hintBody,
            QLabel#statusLabel,
            QLabel#deviceSubtitle,
            QLabel#ruleDetail,
            QLabel#activityText {
                color: #9aa5b1;
            }

            QLabel#statusLabel {
                background: #171b22;
                border: 1px solid #2b333d;
                border-radius: 10px;
                padding: 10px 12px;
            }

            QPushButton#navButton,
            QPushButton#profileButton {
                background: transparent;
                color: #b8c2cc;
                border: 0;
                border-radius: 8px;
                padding: 12px 14px;
                text-align: left;
                font-weight: 650;
            }

            QPushButton#navButton:hover,
            QPushButton#profileButton:hover {
                background: #222731;
                color: #ffffff;
            }

            QPushButton#navButton:checked {
                background: #e8fbf6;
                color: #071512;
            }

            QFrame#divider {
                background: #2a3039;
            }

            QLabel#sectionLabel,
            QLabel#eyebrow {
                color: #65e4d1;
                font-size: 12px;
                font-weight: 800;
                text-transform: uppercase;
                letter-spacing: 1px;
            }

            QFrame#sidebarHint,
            QFrame#panel,
            QFrame#rulesPanel {
                background: #191e26;
                border: 1px solid #29313b;
                border-radius: 14px;
            }

            QLabel#hintTitle,
            QLabel#ruleTitle,
            QLabel#deviceTitle {
                color: #f7fafc;
                font-weight: 750;
            }

            QWidget#workspace {
                background: #101216;
            }

            QLabel#pageTitle {
                color: #f8fafc;
                font-size: 31px;
                font-weight: 850;
            }

            QLineEdit#search {
                background: #171b22;
                color: #f8fafc;
                border: 1px solid #2b333d;
                border-radius: 10px;
                padding: 12px 14px;
                min-width: 280px;
            }

            QLineEdit#search:focus {
                border-color: #2dd4bf;
            }

            QCheckBox#showHidden,
            QCheckBox {
                color: #d7dde4;
                spacing: 8px;
            }

            QCheckBox::indicator {
                width: 42px;
                height: 22px;
                border-radius: 11px;
                background: #303743;
            }

            QCheckBox::indicator:checked {
                background: #2dd4bf;
            }

            QFrame#statCard {
                border-radius: 14px;
                border: 1px solid #2b333d;
                background: #171b22;
            }

            QFrame#statCard[tone="green"] {
                border-color: #1f766c;
            }

            QFrame#statCard[tone="amber"] {
                border-color: #9a6a20;
            }

            QFrame#statCard[tone="coral"] {
                border-color: #b75a4a;
            }

            QLabel#statValue {
                color: #ffffff;
                font-size: 23px;
                font-weight: 850;
            }

            QLabel#statLabel {
                color: #a4afbb;
                font-weight: 650;
            }

            QPushButton#primaryButton {
                background: #2dd4bf;
                color: #06120f;
                border: 0;
                border-radius: 9px;
                padding: 10px 14px;
                font-weight: 800;
            }

            QPushButton#primaryButton:hover {
                background: #5eead4;
            }

            QListWidget#deviceList {
                background: transparent;
                border: 0;
                outline: 0;
            }

            QListWidget#deviceList::item {
                background: transparent;
                border: 0;
            }

            QFrame#deviceCard {
                background: #202630;
                border: 1px solid #303946;
                border-radius: 14px;
            }

            QFrame#deviceCard[hiddenDevice="true"] {
                background: #171b22;
                border-style: dashed;
            }

            QLabel#dragHandle {
                color: #697586;
                font-weight: 900;
            }

            QLabel#priorityBadge {
                background: #101216;
                color: #65e4d1;
                border: 1px solid #2a625b;
                border-radius: 12px;
                font-weight: 900;
            }

            QLabel#statusPill {
                border-radius: 9px;
                padding: 4px 8px;
                font-size: 11px;
                font-weight: 800;
            }

            QLabel#statusPill[tone="default"] {
                background: #2dd4bf;
                color: #06120f;
            }

            QLabel#statusPill[tone="ready"] {
                background: #2a3441;
                color: #cbd5df;
            }

            QLabel#statusPill[tone="muted"] {
                background: #3b2a2a;
                color: #ffb4a6;
            }

            QFrame#meter {
                background: #11151a;
                border-radius: 3px;
            }

            QFrame#meterFill {
                background: #65e4d1;
                border-radius: 3px;
            }

            QPushButton#iconButton,
            QPushButton#dangerIconButton,
            QPushButton#restoreIconButton {
                background: #151a21;
                color: #dfe6ee;
                border: 1px solid #303946;
                border-radius: 9px;
                font-size: 16px;
                font-weight: 800;
            }

            QPushButton#iconButton:hover,
            QPushButton#restoreIconButton:hover {
                border-color: #2dd4bf;
                color: #ffffff;
            }

            QPushButton#dangerIconButton:hover {
                border-color: #ff8a78;
                color: #ffb4a6;
            }

            QFrame#ruleRow {
                background: #202630;
                border: 1px solid #303946;
                border-radius: 12px;
            }

            QLabel#activityTime {
                color: #65e4d1;
                font-weight: 800;
            }

            QLabel#volumeLabel {
                color: #cbd5df;
                font-weight: 800;
            }

            QSlider#volumeSlider::groove:horizontal {
                background: #11151a;
                border-radius: 4px;
                height: 8px;
            }

            QSlider#volumeSlider::sub-page:horizontal {
                background: #65e4d1;
                border-radius: 4px;
            }

            QSlider#volumeSlider::handle:horizontal {
                background: #f8fafc;
                border: 2px solid #65e4d1;
                border-radius: 8px;
                width: 16px;
                margin: -5px 0;
            }

            QScrollBar:vertical {
                background: #141820;
                border: 0;
                border-radius: 6px;
                width: 12px;
                margin: 2px;
            }

            QScrollBar::handle:vertical {
                background: #3a4654;
                border-radius: 6px;
                min-height: 42px;
            }

            QScrollBar::handle:vertical:hover {
                background: #65e4d1;
            }

            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0;
            }
            """
        )


def main() -> None:
    app = QApplication(sys.argv)
    window = SoundManagerWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
