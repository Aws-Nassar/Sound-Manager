import ctypes
import json
import sys
from dataclasses import dataclass
from pathlib import Path

from PyQt5.QtCore import Qt, QSize, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QFontMetrics, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QBoxLayout,
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
    QScrollArea,
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
    local_hidden: bool = False


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
CONFIG_PATH = APP_DIR / "config.json"
BUNDLED_CONFIG_PATH = RESOURCE_DIR / "config.json"
STATE_PATH = APP_DIR / "sound_manager_state.json"
APP_ICON_PATH = RESOURCE_DIR / "assets" / "sound_manager.ico"

DEFAULT_PROFILES = {
    "Default": {
        "description": "Default Sound Manager view with active devices first.",
        "output_keywords": [],
        "input_keywords": [],
        "hidden_devices": [],
        "disabled_devices": [],
        "volumes": {},
        "protected": True,
    },
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

DEFAULT_RULES = {
    "hide_unplugged": True,
    "prefer_headset_for_calls": False,
    "lock_studio_output": False,
}

PROTECTED_PROFILES = {"Default"}


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


class ElidedLabel(QLabel):
    def __init__(self, text: str = "", mode: Qt.TextElideMode = Qt.ElideRight):
        super().__init__()
        self.full_text = text
        self.mode = mode
        self.setText(text)

    def setText(self, text: str) -> None:
        self.full_text = text
        self.setToolTip(text)
        super().setText(self._elided_text())

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        super().setText(self._elided_text())

    def _elided_text(self) -> str:
        width = max(0, self.width() - 2)
        if width <= 0:
            return self.full_text
        return QFontMetrics(self.font()).elidedText(self.full_text, self.mode, width)


class DeviceCard(QFrame):
    hide_requested = pyqtSignal(str)
    disable_requested = pyqtSignal(str)
    move_requested = pyqtSignal(str, int)
    volume_requested = pyqtSignal(str, int)

    def __init__(self, device: AudioDevice, priority: int):
        super().__init__()
        self.device = device
        state = self._state_for(device)
        source_enabled = state not in {"disabled", "unplugged", "not_present"}
        self.setObjectName("deviceCard")
        self.setProperty("hiddenDevice", state in {"hidden", "disabled", "unplugged", "not_present"})

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 12, 12, 12)
        root.setSpacing(10)

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

        title = ElidedLabel(device.name)
        title.setObjectName("deviceTitle")
        title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        title.setMinimumWidth(60)

        status = QLabel(self._state_label(state, device))
        status.setObjectName("statusPill")
        status.setProperty("tone", state)

        title_row.addWidget(title)
        title_row.addWidget(status)

        subtitle = ElidedLabel(device.subtitle)
        subtitle.setObjectName("deviceSubtitle")
        subtitle.setMinimumWidth(60)

        copy.addLayout(title_row)
        copy.addWidget(subtitle)

        volume_row = QHBoxLayout()
        volume_row.setSpacing(10)
        volume_label = QLabel(f"{device.level}%")
        volume_label.setObjectName("volumeLabel")
        volume_label.setFixedWidth(42)
        volume_slider = QSlider(Qt.Horizontal)
        volume_slider.setObjectName("volumeSlider")
        volume_slider.setRange(0, 100)
        volume_slider.setValue(device.level)
        volume_slider.setEnabled(source_enabled)
        volume_slider.setCursor(Qt.PointingHandCursor)

        def update_volume(value: int) -> None:
            volume_label.setText(f"{value}%")
            self.volume_requested.emit(device.device_id, value)

        volume_slider.valueChanged.connect(update_volume)
        volume_row.addWidget(volume_slider, 1)
        volume_row.addWidget(volume_label)
        copy.addLayout(volume_row)

        actions_widget = QWidget()
        actions_widget.setObjectName("deviceActions")
        actions_widget.setFixedWidth(116)
        actions = QGridLayout(actions_widget)
        actions.setContentsMargins(0, 0, 0, 0)
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)

        up = self._icon_button("↑", "Move higher priority")
        down = self._icon_button("↓", "Move lower priority")
        hide = self._action_button("Show" if device.local_hidden else "Hide", "Show this source in this app" if device.local_hidden else "Hide this source only inside Sound Manager")
        hide.setObjectName("restoreActionButton" if device.local_hidden else "secondaryActionButton")
        disable = self._action_button("On" if state == "disabled" else "Off", "Enable this source in Windows" if state == "disabled" else "Disable this source in Windows")
        disable.setObjectName("restoreActionButton" if state == "disabled" else "dangerActionButton")
        disable.setEnabled(state not in {"unplugged", "not_present"})

        up.clicked.connect(lambda: self.move_requested.emit(device.device_id, -1))
        down.clicked.connect(lambda: self.move_requested.emit(device.device_id, 1))
        hide.clicked.connect(lambda: self.hide_requested.emit(device.device_id))
        disable.clicked.connect(lambda: self.disable_requested.emit(device.device_id))

        actions.addWidget(up, 0, 0)
        actions.addWidget(down, 0, 1)
        actions.addWidget(hide, 1, 0)
        actions.addWidget(disable, 1, 1)

        root.addWidget(handle)
        root.addWidget(badge)
        root.addLayout(copy, 1)
        root.addWidget(actions_widget)

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
        button.setFixedSize(52, 32)
        button.setCursor(Qt.PointingHandCursor)
        return button

    def _state_for(self, device: AudioDevice) -> str:
        status = device.status.lower().replace(" ", "_")
        if status in {"disabled", "unplugged", "not_present"}:
            return status
        if device.hidden:
            return "disabled"
        if device.local_hidden:
            return "hidden"
        if status == "default":
            return "default"
        return "active"

    def _state_label(self, state: str, device: AudioDevice) -> str:
        return {
            "default": "Default",
            "active": "Active",
            "hidden": "Hidden",
            "disabled": "Disabled",
            "unplugged": "Unplugged",
            "not_present": "Not present",
        }.get(state, device.status)


class StatCard(QFrame):
    def __init__(self, label: str, value: str, tone: str):
        super().__init__()
        self.setObjectName("statCard")
        self.setProperty("tone", tone)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)

        self.value = ElidedLabel(value)
        self.value.setObjectName("statValue")
        self.value.setMinimumWidth(40)
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
        self.hidden_device_ids = set(self.state.get("hidden_devices", []))
        self.deleted_profile_names = set(self.state.get("deleted_profiles", []))
        self.profiles = self._load_profiles()
        self.profile_buttons = {}
        self.profile_layout = None
        self.rules = {**DEFAULT_RULES, **self.state.get("rules", {})}
        self.rule_toggles = {}
        self.disable_refresh_attempts = 0
        self.last_list_signature = None
        self.volume_update_timer = QTimer(self)
        self.volume_update_timer.setSingleShot(True)
        self.pending_volume_change = None
        self.default_update_timer = QTimer(self)
        self.default_update_timer.setSingleShot(True)
        self.pending_default_kind = None
        self.show_hidden = False
        self.current_kind = "output"
        self.current_profile_name = self.state.get("current_profile", "Default")
        if self.current_profile_name not in self.profiles:
            self.current_profile_name = "Default"

        self.setWindowTitle("Sound Manager")
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))
        self.setMinimumSize(860, 560)

        self.setCentralWidget(self._build_ui())
        self._apply_style()
        self.reload_devices()
        self.refresh()
        self._fit_to_screen()

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(7000)
        self.poll_timer.timeout.connect(lambda: self.reload_devices(refresh=True))
        self.poll_timer.start()
        self.volume_update_timer.timeout.connect(self._flush_volume_change)
        self.default_update_timer.timeout.connect(self._flush_default_from_priority)

    def _fit_to_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            self.resize(1100, 760)
            return

        available = screen.availableGeometry()
        width = min(1220, max(self.minimumWidth(), int(available.width() * 0.86)))
        height = min(820, max(self.minimumHeight(), int(available.height() * 0.84)))
        self.resize(width, height)
        self.move(
            available.x() + (available.width() - width) // 2,
            available.y() + (available.height() - height) // 2,
        )
        self._update_responsive_layout(width, height)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._update_responsive_layout(self.width(), self.height())

    def _update_responsive_layout(self, width: int, height: int) -> None:
        if hasattr(self, "body_layout"):
            compact = width < 1120 or height < 700
            self.body_layout.setDirection(QBoxLayout.TopToBottom if compact else QBoxLayout.LeftToRight)
            if hasattr(self, "rules_panel"):
                self.rules_panel.setMaximumWidth(16777215 if compact else 320)

    def _build_ui(self) -> QWidget:
        shell = QWidget()
        shell.setObjectName("shell")
        root = QHBoxLayout(shell)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())
        workspace_scroll = QScrollArea()
        workspace_scroll.setObjectName("workspaceScroll")
        workspace_scroll.setWidgetResizable(True)
        workspace_scroll.setFrameShape(QFrame.NoFrame)
        workspace_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        workspace_scroll.setWidget(self._build_workspace())
        root.addWidget(workspace_scroll, 1)
        return shell

    def _build_sidebar(self) -> QWidget:
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setMinimumWidth(258)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(22, 24, 18, 22)
        layout.setSpacing(16)

        logo_row = QHBoxLayout()
        logo = QLabel("SM")
        logo.setObjectName("logo")
        logo.setAlignment(Qt.AlignCenter)
        logo.setFixedSize(48, 48)

        brand = QVBoxLayout()
        title = QLabel("Sound Manager")
        title.setObjectName("brandTitle")
        subtitle = QLabel("Windows audio")
        subtitle.setObjectName("brandSubtitle")
        subtitle.setWordWrap(False)
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
        save_profile.setMinimumHeight(44)
        save_profile.setCursor(Qt.PointingHandCursor)
        save_profile.clicked.connect(self.create_profile_from_current)
        layout.addWidget(save_profile)

        manage_row = QHBoxLayout()
        manage_row.setSpacing(8)
        rename_profile = QPushButton("Rename")
        rename_profile.setObjectName("secondarySideButton")
        rename_profile.setMinimumHeight(42)
        rename_profile.setCursor(Qt.PointingHandCursor)
        rename_profile.clicked.connect(self.rename_current_profile)
        delete_profile = QPushButton("Delete")
        delete_profile.setObjectName("dangerSideButton")
        delete_profile.setMinimumHeight(42)
        delete_profile.setCursor(Qt.PointingHandCursor)
        delete_profile.clicked.connect(self.delete_current_profile)
        manage_row.addWidget(rename_profile)
        manage_row.addWidget(delete_profile)
        layout.addLayout(manage_row)

        layout.addStretch()

        hint = QFrame()
        hint.setObjectName("sidebarHint")
        hint_layout = QVBoxLayout(hint)
        hint_layout.setContentsMargins(14, 14, 14, 14)
        hint_layout.setSpacing(6)
        hint_title = QLabel("Quiet picker")
        hint_title.setObjectName("hintTitle")
        hint_body = QLabel("Hide keeps a source out of this app. Disable keeps it out of Windows and other app input/output menus.")
        hint_body.setObjectName("hintBody")
        hint_body.setWordWrap(True)
        hint_layout.addWidget(hint_title)
        hint_layout.addWidget(hint_body)
        layout.addWidget(hint)

        scroll = QScrollArea()
        scroll.setObjectName("sidebarScroll")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setFixedWidth(258)
        scroll.setWidget(sidebar)
        return scroll

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
        layout.setContentsMargins(28, 26, 28, 26)
        layout.setSpacing(18)

        header = QVBoxLayout()
        header.setSpacing(12)

        headline = QVBoxLayout()
        eyebrow = QLabel("Device priority")
        eyebrow.setObjectName("eyebrow")
        title = ElidedLabel("Control Windows sound")
        title.setObjectName("pageTitle")
        headline.addWidget(eyebrow)
        headline.addWidget(title)

        controls = QHBoxLayout()
        controls.setSpacing(12)
        self.search = QLineEdit()
        self.search.setObjectName("search")
        self.search.setPlaceholderText("Search devices")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.refresh)

        self.hidden_toggle = QCheckBox("Show all")
        self.hidden_toggle.setObjectName("showHidden")
        self.hidden_toggle.setToolTip("Show hidden, disabled, unplugged, and not-present sources.")
        self.hidden_toggle.stateChanged.connect(self._toggle_hidden)

        controls.addWidget(self.search, 1)
        controls.addWidget(self.hidden_toggle)
        header.addLayout(headline)
        header.addLayout(controls)
        layout.addLayout(header)

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        stats = QGridLayout()
        stats.setHorizontalSpacing(14)
        stats.setVerticalSpacing(12)
        self.visible_stat = StatCard("Visible devices", "0", "green")
        self.hidden_stat = StatCard("Disabled sources", "0", "amber")
        self.default_stat = StatCard("Current default", "None", "coral")
        stats.addWidget(self.visible_stat, 0, 0)
        stats.addWidget(self.hidden_stat, 0, 1)
        stats.addWidget(self.default_stat, 0, 2)
        layout.addLayout(stats)

        body = QBoxLayout(QBoxLayout.LeftToRight)
        body.setSpacing(18)
        self.body_layout = body

        list_panel = QFrame()
        list_panel.setObjectName("panel")
        self.list_panel = list_panel
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

        self.rules_panel = self._build_rules_panel()
        body.addWidget(list_panel, 1)
        body.addWidget(self.rules_panel)
        layout.addLayout(body, 1)

        return workspace

    def _build_rules_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("rulesPanel")
        panel.setMinimumWidth(260)
        panel.setMaximumWidth(320)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel("Source Visibility")
        title.setObjectName("panelTitle")
        layout.addWidget(title)

        self.app_hidden_count_label = QLabel("0")
        self.app_hidden_count_label.setObjectName("visibilityCount")
        self.windows_disabled_count_label = QLabel("0")
        self.windows_disabled_count_label.setObjectName("visibilityCount")
        self.unavailable_count_label = QLabel("0")
        self.unavailable_count_label.setObjectName("visibilityCount")

        layout.addWidget(self._visibility_row("Hidden in app", "Hidden only inside Sound Manager.", self.app_hidden_count_label))
        layout.addWidget(self._visibility_row("Windows disabled", "Disabled sources stay out of other apps.", self.windows_disabled_count_label))
        layout.addWidget(self._visibility_row("Unavailable", "Unplugged or not-present sources from Windows.", self.unavailable_count_label))

        self.show_all_sources_button = QPushButton("Show hidden/disabled")
        self.show_all_sources_button.setObjectName("primaryButton")
        self.show_all_sources_button.setCursor(Qt.PointingHandCursor)
        self.show_all_sources_button.clicked.connect(self.toggle_show_all_sources)
        layout.addWidget(self.show_all_sources_button)

        layout.addStretch()
        return panel

    def _visibility_row(self, title: str, detail: str, count_label: QLabel) -> QWidget:
        row = QFrame()
        row.setObjectName("visibilityRow")
        row.setFixedHeight(46)
        row.setToolTip(detail)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(12, 7, 12, 7)
        layout.setSpacing(12)

        name = ElidedLabel(title)
        name.setObjectName("ruleTitle")
        name.setMinimumWidth(100)

        count_label.setAlignment(Qt.AlignCenter)
        count_label.setFixedSize(42, 32)
        layout.addWidget(name, 1)
        layout.addWidget(count_label)
        return row

    def _rule_toggle(self, title: str, detail: str, key: str, enabled: bool) -> QWidget:
        row = QFrame()
        row.setObjectName("ruleRow")
        row.setProperty("disabledRule", not enabled)
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
        toggle.setChecked(bool(self.rules.get(key, False)))
        toggle.setEnabled(enabled)
        toggle.setCursor(Qt.PointingHandCursor)
        toggle.stateChanged.connect(lambda _state, rule_key=key: self.set_rule(rule_key, toggle.isChecked()))
        self.rule_toggles[key] = toggle

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
            chip.setCheckable(True)
            chip.setChecked(text == self.current_profile_name)
            chip.setCursor(Qt.PointingHandCursor)
            chip.setMinimumHeight(32)
            chip.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            chip.setToolTip(self.profiles[text].get("description", "Apply this audio profile."))
            chip.clicked.connect(lambda checked=False, name=text: self.apply_profile(name))
            self.profile_buttons[text] = chip
            self.profile_layout.addWidget(chip)

    def set_rule(self, key: str, enabled: bool) -> None:
        self.rules[key] = bool(enabled)
        self.state["rules"] = self.rules
        self._save_state()
        self.last_list_signature = None
        if key == "hide_unplugged":
            self.status_label.setText("Unplugged and not-present source filtering updated.")
            self.refresh()

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
        self.state["hidden_devices"] = sorted(self.hidden_device_ids)
        self.state["rules"] = self.rules
        self.state["deleted_profiles"] = sorted(self.deleted_profile_names)
        self.state["current_profile"] = self.current_profile_name
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
            for name in self.state.get("deleted_profiles", []):
                if name not in PROTECTED_PROFILES:
                    profiles.pop(name, None)
            return profiles

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            profiles.update(data.get("profiles", {}))
            profiles.update(self.state.get("profiles", {}))
            for name in self.state.get("deleted_profiles", []):
                if name not in PROTECTED_PROFILES:
                    profiles.pop(name, None)
            return profiles
        except (OSError, json.JSONDecodeError):
            profiles.update(self.state.get("profiles", {}))
            for name in self.state.get("deleted_profiles", []):
                if name not in PROTECTED_PROFILES:
                    profiles.pop(name, None)
            return profiles

    def set_kind(self, kind: str) -> None:
        self.current_kind = kind
        self.output_nav.setChecked(kind == "output")
        self.input_nav.setChecked(kind == "input")
        self.refresh()

    def _toggle_hidden(self) -> None:
        self.show_hidden = self.hidden_toggle.isChecked()
        self.last_list_signature = None
        self.refresh()

    def toggle_show_all_sources(self) -> None:
        self.hidden_toggle.setChecked(not self.hidden_toggle.isChecked())

    def current_devices(self) -> list[AudioDevice]:
        return self.outputs if self.current_kind == "output" else self.inputs

    def set_current_devices(self, devices: list[AudioDevice]) -> None:
        if self.current_kind == "output":
            self.outputs = devices
        else:
            self.inputs = devices
        self.priority_order[self.current_kind] = [device.device_id for device in devices]
        self._save_state()

    def _save_current_profile_snapshot(self) -> None:
        if not self.current_profile_name or self.current_profile_name not in self.profiles:
            return

        profile = self._snapshot_current_profile(self.current_profile_name)
        self.profiles[self.current_profile_name] = profile
        self.state.setdefault("profiles", {})[self.current_profile_name] = profile
        self.deleted_profile_names.discard(self.current_profile_name)
        self._save_state()

    def reload_devices(self, refresh: bool = False) -> None:
        if refresh and QApplication.mouseButtons() != Qt.NoButton:
            return

        if self.backend.available:
            self.outputs = self._apply_state_overrides(self._apply_saved_order("output", self.backend.list_devices("output")))
            self.inputs = self._apply_state_overrides(self._apply_saved_order("input", self.backend.list_devices("input")))
            if hasattr(self, "status_label") and not self.status_label.text():
                self.status_label.setText("Connected to Windows Core Audio. Top priority changes update the real Windows default device.")
        else:
            self.outputs = self._apply_state_overrides(self._apply_saved_order("output", list(OUTPUT_DEVICES)))
            self.inputs = self._apply_state_overrides(self._apply_saved_order("input", list(INPUT_DEVICES)))
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

    def _apply_state_overrides(self, devices: list[AudioDevice]) -> list[AudioDevice]:
        for device in devices:
            device.local_hidden = device.device_id in self.hidden_device_ids
            if device.device_id in self.disabled_device_ids:
                device.hidden = True
                device.status = "Disabled"
                device.level = 0
        return devices

    def _is_hidden_or_disabled(self, device: AudioDevice) -> bool:
        transient_unavailable = device.status in {"Unplugged", "Not present"}
        return (
            device.local_hidden
            or device.hidden
            or device.status == "Disabled"
            or (self.rules.get("hide_unplugged", True) and transient_unavailable)
        )

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
            if (self.show_hidden or not self._is_hidden_or_disabled(device))
            and (not query or query in device.name.lower() or query in device.subtitle.lower())
        ]

        signature = (
            self.current_kind,
            self.show_hidden,
            query,
            tuple((device.device_id, device.status, device.level, device.hidden, device.local_hidden) for device in visible),
        )
        if signature == self.last_list_signature:
            self._update_stats()
            return
        self.last_list_signature = signature

        scroll_value = self.list_widget.verticalScrollBar().value()
        self.list_widget.setUpdatesEnabled(False)
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for priority, device in enumerate(visible, start=1):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, device.device_id)
            card = DeviceCard(device, priority)
            card.hide_requested.connect(self.toggle_device_hidden)
            card.disable_requested.connect(self.toggle_device_disabled)
            card.move_requested.connect(self.move_device)
            card.volume_requested.connect(self.set_device_volume)
            item.setSizeHint(QSize(100, 114))
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, card)
        self.list_widget.blockSignals(False)
        self.list_widget.verticalScrollBar().setValue(scroll_value)
        self.list_widget.setUpdatesEnabled(True)

        self._update_stats()

    def _update_stats(self) -> None:
        all_current = self.current_devices()
        visible = [device for device in all_current if not self._is_hidden_or_disabled(device)]
        hidden = [device for device in all_current if self._is_hidden_or_disabled(device)]
        default = next((device.name for device in all_current if device.status == "Default" and not self._is_hidden_or_disabled(device)), "None")

        self.visible_stat.value.setText(str(len(visible)))
        self.hidden_stat.value.setText(str(len(hidden)))
        self.default_stat.value.setText(default)
        if hasattr(self, "app_hidden_count_label"):
            self.app_hidden_count_label.setText(str(sum(1 for device in all_current if device.local_hidden)))
            self.windows_disabled_count_label.setText(str(sum(1 for device in all_current if device.status == "Disabled" or device.device_id in self.disabled_device_ids)))
            self.unavailable_count_label.setText(str(sum(1 for device in all_current if device.status in {"Unplugged", "Not present"})))
            self.show_all_sources_button.setText("Show active only" if self.show_hidden else "Show hidden/disabled")

    def toggle_device_hidden(self, device_id: str) -> None:
        device = next((item for item in self.current_devices() if item.device_id == device_id), None)
        if not device:
            return

        if device.local_hidden:
            self.hidden_device_ids.discard(device_id)
            device.local_hidden = False
            self.status_label.setText(f"{device.name} is visible in Sound Manager again.")
        else:
            self.hidden_device_ids.add(device_id)
            device.local_hidden = True
            self.status_label.setText(f"{device.name} hidden from Sound Manager.")

        self._save_state()
        self.last_list_signature = None
        self.refresh()
        self._save_current_profile_snapshot()

    def toggle_device_disabled(self, device_id: str) -> None:
        device = next((item for item in self.current_devices() if item.device_id == device_id), None)
        if not device:
            return

        enable_device = device.status == "Disabled" or device_id in self.disabled_device_ids

        if self.backend.available:
            message = self.backend.set_enabled(device_id, enable_device)
            self.status_label.setText(message)
            if "cancelled" not in message.lower() and "blocked" not in message.lower():
                self._set_device_disabled_locally(device_id, not enable_device)
                self.disable_refresh_attempts = 5
                QTimer.singleShot(1000, self._refresh_after_disable_action)
            return

        self._set_device_disabled_locally(device_id, not enable_device)

    def _set_device_disabled_locally(self, device_id: str, disabled: bool) -> None:
        if disabled:
            self.disabled_device_ids.add(device_id)
        else:
            self.disabled_device_ids.discard(device_id)

        for collection in (self.outputs, self.inputs):
            for device in collection:
                if device.device_id == device_id:
                    device.hidden = disabled
                    device.status = "Disabled" if disabled else "Active"
                    break
        self._save_state()
        self.last_list_signature = None
        self.refresh()
        self._save_current_profile_snapshot()

    def _refresh_after_disable_action(self) -> None:
        self.reload_devices(refresh=True)
        self.disable_refresh_attempts -= 1
        if self.disable_refresh_attempts > 0:
            QTimer.singleShot(2000, self._refresh_after_disable_action)

    def set_device_volume(self, device_id: str, value: int) -> None:
        for device in self.current_devices():
            if device.device_id == device_id:
                device.level = value
                break

        if not self.backend.available:
            self._save_current_profile_snapshot()
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
            self._save_current_profile_snapshot()
        except Exception as exc:
            self.status_label.setText(f"Windows rejected the volume change: {exc}")

    def apply_profile(self, profile_name: str) -> None:
        profile = self.profiles.get(profile_name)
        if not profile:
            return
        if profile_name != self.current_profile_name:
            self._save_current_profile_snapshot()

        previous_disabled = set(self.disabled_device_ids)
        self.current_profile_name = profile_name

        self.outputs = self._order_for_profile("output", self.outputs, profile)
        self.inputs = self._order_for_profile("input", self.inputs, profile)
        self.priority_order["output"] = [device.device_id for device in self.outputs]
        self.priority_order["input"] = [device.device_id for device in self.inputs]

        self._apply_profile_disabled_sources(profile, apply_windows=False, previous_disabled=previous_disabled)
        self._apply_profile_hidden_sources(profile)
        self._apply_profile_volumes(profile, apply_windows=False)
        self._apply_profile_default("output", profile, apply_windows=False)
        self._apply_profile_default("input", profile, apply_windows=False)
        self._save_state()
        self.status_label.setText(f"Applying {profile_name} profile to Windows...")
        self._refresh_profile_buttons()
        self.last_list_signature = None
        self.refresh()
        QTimer.singleShot(
            80,
            lambda name=profile_name, snapshot=dict(profile), disabled=previous_disabled: self._finish_profile_windows_apply(name, snapshot, disabled),
        )

    def _finish_profile_windows_apply(self, profile_name: str, profile: dict, previous_disabled: set[str]) -> None:
        if profile_name != self.current_profile_name:
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self._apply_profile_disabled_sources(profile, apply_windows=True, previous_disabled=previous_disabled)
            self._apply_profile_default("output", profile, apply_windows=True)
            self._apply_profile_default("input", profile, apply_windows=True)
            self._apply_profile_volumes(profile, apply_windows=True)
        finally:
            QApplication.restoreOverrideCursor()

        if profile_name != self.current_profile_name:
            return

        self._save_state()
        self.status_label.setText(f"{profile_name} profile applied. {profile.get('description', '')}".strip())
        self.last_list_signature = None
        self.refresh()

    def create_profile_from_current(self) -> None:
        name, accepted = QInputDialog.getText(self, "Save profile", "Profile name:")
        name = name.strip()
        if not accepted or not name:
            return
        if name in PROTECTED_PROFILES:
            self.status_label.setText("Default profile is protected. Choose a different profile name.")
            return
        if name in self.profiles:
            self.status_label.setText(f"A profile named {name} already exists.")
            return

        profile = self._snapshot_current_profile(name)
        self.profiles[name] = profile
        self.state.setdefault("profiles", {})[name] = profile
        self.deleted_profile_names.discard(name)
        self.current_profile_name = name
        self._save_state()
        self.status_label.setText(f"{name} profile saved with current order, volumes, defaults, and disabled sources.")
        self._refresh_profile_buttons()

    def rename_current_profile(self) -> None:
        old_name = self.current_profile_name
        if old_name in PROTECTED_PROFILES or self.profiles.get(old_name, {}).get("protected"):
            self.status_label.setText("Default profile cannot be renamed.")
            return
        if old_name not in self.profiles:
            self.status_label.setText("Select a custom profile before renaming.")
            return

        new_name, accepted = QInputDialog.getText(self, "Rename profile", "New profile name:", text=old_name)
        new_name = new_name.strip()
        if not accepted or not new_name or new_name == old_name:
            return
        if new_name in self.profiles:
            self.status_label.setText(f"A profile named {new_name} already exists.")
            return

        profile = self.profiles.pop(old_name)
        self.profiles[new_name] = profile
        user_profiles = self.state.setdefault("profiles", {})
        user_profiles.pop(old_name, None)
        user_profiles[new_name] = profile
        if old_name not in user_profiles:
            self.deleted_profile_names.add(old_name)
        self.deleted_profile_names.discard(new_name)
        self.current_profile_name = new_name
        self._save_state()
        self._refresh_profile_buttons()
        self.status_label.setText(f"Profile renamed to {new_name}.")

    def delete_current_profile(self) -> None:
        name = self.current_profile_name
        if name in PROTECTED_PROFILES or self.profiles.get(name, {}).get("protected"):
            self.status_label.setText("Default profile cannot be deleted.")
            return
        if name not in self.profiles:
            self.status_label.setText("Select a custom profile before deleting.")
            return

        self.profiles.pop(name, None)
        self.state.setdefault("profiles", {}).pop(name, None)
        self.deleted_profile_names.add(name)
        self.current_profile_name = "Default"
        self._save_state()
        self._refresh_profile_buttons()
        self.status_label.setText(f"{name} profile deleted.")

    def _snapshot_current_profile(self, name: str) -> dict:
        existing = self.profiles.get(name, {})
        disabled = set(self.disabled_device_ids)
        hidden = set(self.hidden_device_ids)
        volumes = {device.device_id: device.level for device in self.outputs + self.inputs}
        output_default = next((device.device_id for device in self.outputs if not self._is_hidden_or_disabled(device)), "")
        input_default = next((device.device_id for device in self.inputs if not self._is_hidden_or_disabled(device)), "")

        profile = {
            "description": existing.get("description", f"Saved snapshot for {name}."),
            "output_order": [device.device_id for device in self.outputs],
            "input_order": [device.device_id for device in self.inputs],
            "output_default": output_default,
            "input_default": input_default,
            "volumes": volumes,
            "hidden_devices": sorted(hidden),
            "disabled_devices": sorted(disabled),
        }
        if existing.get("protected") or name in PROTECTED_PROFILES:
            profile["protected"] = True
        return profile

    def _order_for_profile(self, kind: str, devices: list[AudioDevice], profile: dict) -> list[AudioDevice]:
        order = profile.get(f"{kind}_order")
        if order:
            lookup = {device.device_id: device for device in devices}
            ordered = [lookup[device_id] for device_id in order if device_id in lookup]
            ordered_ids = {device.device_id for device in ordered}
            ordered.extend(device for device in devices if device.device_id not in ordered_ids)
            return ordered

        return self._rank_for_profile(devices, profile.get(f"{kind}_keywords", []))

    def _apply_profile_disabled_sources(self, profile: dict, apply_windows: bool = True, previous_disabled: set[str] | None = None) -> None:
        if "disabled_devices" not in profile:
            return

        previous = set(self.disabled_device_ids if previous_disabled is None else previous_disabled)
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

        if apply_windows and self.backend.available:
            messages = []
            if to_disable:
                messages.append(self.backend.set_many_enabled(to_disable, False))
            if to_enable:
                messages.append(self.backend.set_many_enabled(to_enable, True))
            if messages:
                self.status_label.setText(" ".join(messages))
                self.disable_refresh_attempts = 5
                QTimer.singleShot(1000, self._refresh_after_disable_action)

    def _apply_profile_hidden_sources(self, profile: dict) -> None:
        if "hidden_devices" not in profile:
            return

        self.hidden_device_ids = set(profile.get("hidden_devices", []))
        for device in self.outputs + self.inputs:
            device.local_hidden = device.device_id in self.hidden_device_ids

    def _apply_profile_volumes(self, profile: dict, apply_windows: bool = True) -> None:
        volumes = profile.get("volumes", {})
        if not volumes:
            return

        for device in self.outputs + self.inputs:
            if device.device_id not in volumes or device.hidden:
                continue

            value = int(volumes[device.device_id])
            device.level = value
            if apply_windows and self.backend.available:
                try:
                    self.backend.set_volume(device.device_id, value)
                except Exception:
                    pass

    def _rank_for_profile(self, devices: list[AudioDevice], keywords: list[str]) -> list[AudioDevice]:
        normalized = [keyword.lower() for keyword in keywords]

        def score(device: AudioDevice) -> tuple[int, int, str]:
            haystack = f"{device.name} {device.subtitle}".lower()
            match_index = next((index for index, keyword in enumerate(normalized) if keyword in haystack), len(normalized))
            return (self._is_hidden_or_disabled(device), match_index, device.name.lower())

        return sorted(devices, key=score)

    def _apply_profile_default(self, kind: str, profile: dict, apply_windows: bool = True) -> None:
        devices = self.outputs if kind == "output" else self.inputs
        preferred_id = profile.get(f"{kind}_default", "")
        first_active = next((device for device in devices if device.device_id == preferred_id and not self._is_hidden_or_disabled(device)), None)
        if first_active is None:
            first_active = next((device for device in devices if not self._is_hidden_or_disabled(device)), None)

        for device in devices:
            if device.status == "Default" and device is not first_active:
                device.status = "Active"
        if first_active:
            first_active.status = "Default"

        if not first_active or not apply_windows or not self.backend.available:
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
        self._schedule_default_from_priority()
        self.last_list_signature = None
        self.refresh()
        self._save_current_profile_snapshot()

    def sort_current(self) -> None:
        devices = sorted(
            self.current_devices(),
            key=lambda device: (self._is_hidden_or_disabled(device), device.status != "Default", -device.level),
        )
        self.set_current_devices(devices)
        self._schedule_default_from_priority()
        self.last_list_signature = None
        self.refresh()
        self._save_current_profile_snapshot()

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
        self._schedule_default_from_priority()
        self.last_list_signature = None
        self.refresh()
        self._save_current_profile_snapshot()

    def _schedule_default_from_priority(self) -> None:
        if not self.backend.available:
            return

        self.pending_default_kind = self.current_kind
        self.default_update_timer.start(350)

    def _flush_default_from_priority(self) -> None:
        if not self.backend.available:
            return

        devices = self.outputs if self.pending_default_kind == "output" else self.inputs
        kind = self.pending_default_kind or self.current_kind
        self.pending_default_kind = None

        first_active = next((device for device in devices if not self._is_hidden_or_disabled(device)), None)
        if not first_active:
            self.status_label.setText("No active device is available to set as the Windows default.")
            return

        try:
            self.backend.set_default(first_active.device_id)
            self.status_label.setText(f"Windows default {kind} set to {first_active.name}.")
            self._save_current_profile_snapshot()
            QTimer.singleShot(1000, lambda: self.reload_devices(refresh=True))
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

            QScrollArea#sidebarScroll {
                background: #171a20;
                border: 0;
            }

            QScrollArea#workspaceScroll {
                background: #101216;
                border: 0;
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
            QPushButton#profileButton,
            QPushButton#primarySideButton {
                background: transparent;
                color: #b8c2cc;
                border: 0;
                border-radius: 8px;
                padding: 8px 14px;
                text-align: left;
                font-weight: 650;
            }

            QPushButton#navButton:hover,
            QPushButton#profileButton:hover,
            QPushButton#primarySideButton:hover {
                background: #222731;
                color: #ffffff;
            }

            QPushButton#navButton:checked {
                background: #e8fbf6;
                color: #071512;
            }

            QPushButton#profileButton:checked {
                background: #26313c;
                color: #65e4d1;
                border: 1px solid #2a625b;
            }

            QPushButton#primarySideButton {
                background: #2dd4bf;
                color: #06120f;
                border: 0;
                font-weight: 800;
            }

            QPushButton#secondarySideButton,
            QPushButton#dangerSideButton {
                background: #151a21;
                border: 1px solid #303946;
                color: #dfe6ee;
                border-radius: 8px;
                padding: 10px 12px;
                font-weight: 750;
            }

            QPushButton#dangerSideButton:hover {
                border-color: #ff8a78;
                color: #ffb4a6;
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
                min-width: 160px;
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

            QWidget#deviceActions {
                background: transparent;
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

            QLabel#statusPill[tone="active"] {
                background: #2a3441;
                color: #cbd5df;
            }

            QLabel#statusPill[tone="hidden"] {
                background: #2e3340;
                color: #f8d678;
            }

            QLabel#statusPill[tone="disabled"] {
                background: #3b2a2a;
                color: #ffb4a6;
            }

            QLabel#statusPill[tone="unplugged"],
            QLabel#statusPill[tone="not_present"] {
                background: #2d2634;
                color: #c4b5fd;
            }

            QPushButton#iconButton,
            QPushButton#secondaryActionButton,
            QPushButton#dangerActionButton,
            QPushButton#restoreActionButton {
                background: #151a21;
                color: #dfe6ee;
                border: 1px solid #303946;
                border-radius: 9px;
                font-weight: 800;
            }

            QPushButton#iconButton {
                font-size: 16px;
            }

            QPushButton#secondaryActionButton,
            QPushButton#dangerActionButton,
            QPushButton#restoreActionButton {
                font-size: 11px;
            }

            QPushButton#iconButton:hover,
            QPushButton#secondaryActionButton:hover,
            QPushButton#restoreActionButton:hover {
                border-color: #2dd4bf;
                color: #ffffff;
            }

            QPushButton#dangerActionButton:hover {
                border-color: #ff8a78;
                color: #ffb4a6;
            }

            QPushButton#dangerActionButton:disabled {
                color: #66717f;
                border-color: #252b34;
            }

            QFrame#ruleRow {
                background: #202630;
                border: 1px solid #303946;
                border-radius: 12px;
            }

            QFrame#visibilityRow {
                background: #202630;
                border: 1px solid #303946;
                border-radius: 12px;
            }

            QLabel#visibilityCount {
                background: #11151a;
                border: 1px solid #2a625b;
                border-radius: 10px;
                color: #65e4d1;
                font-size: 17px;
                font-weight: 850;
            }

            QFrame#ruleRow[disabledRule="true"] {
                background: #181d25;
                border-color: #252d37;
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
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ColdWorks.SoundManager")
    except Exception:
        pass

    app = QApplication(sys.argv)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    window = SoundManagerWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
