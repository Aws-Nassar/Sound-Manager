import sys
from dataclasses import dataclass

from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui import QFont
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
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)


@dataclass
class AudioDevice:
    device_id: str
    kind: str
    name: str
    subtitle: str
    status: str
    level: int
    hidden: bool = False


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

        status = QLabel("Hidden" if device.hidden else device.status)
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

        actions = QHBoxLayout()
        actions.setSpacing(6)

        up = self._icon_button("↑", "Move higher priority")
        down = self._icon_button("↓", "Move lower priority")
        hide = self._icon_button("↺" if device.hidden else "×", "Restore device" if device.hidden else "Hide device")
        hide.setObjectName("dangerIconButton" if not device.hidden else "restoreIconButton")

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
        self.outputs = list(OUTPUT_DEVICES)
        self.inputs = list(INPUT_DEVICES)
        self.show_hidden = False
        self.current_kind = "output"

        self.setWindowTitle("Sound Manager")
        self.setMinimumSize(1120, 760)
        self.resize(1220, 820)

        self.setCentralWidget(self._build_ui())
        self._apply_style()
        self.refresh()

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

        for text in ("Work Mode", "Gaming", "Recording"):
            chip = QPushButton(text)
            chip.setObjectName("profileButton")
            chip.setCursor(Qt.PointingHandCursor)
            layout.addWidget(chip)

        layout.addStretch()

        hint = QFrame()
        hint.setObjectName("sidebarHint")
        hint_layout = QVBoxLayout(hint)
        hint_layout.setContentsMargins(14, 14, 14, 14)
        hint_layout.setSpacing(6)
        hint_title = QLabel("Quiet picker")
        hint_title.setObjectName("hintTitle")
        hint_body = QLabel("Hidden devices stay out of your Windows input and output menus.")
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

        self.hidden_toggle = QCheckBox("Show hidden")
        self.hidden_toggle.setObjectName("showHidden")
        self.hidden_toggle.stateChanged.connect(self._toggle_hidden)

        header.addLayout(headline, 1)
        header.addWidget(self.search)
        header.addWidget(self.hidden_toggle)
        layout.addLayout(header)

        stats = QGridLayout()
        stats.setHorizontalSpacing(14)
        self.visible_stat = StatCard("Visible devices", "0", "green")
        self.hidden_stat = StatCard("Hidden devices", "0", "amber")
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

        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for priority, device in enumerate(visible, start=1):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, device.device_id)
            card = DeviceCard(device, priority)
            card.hide_requested.connect(self.toggle_device_hidden)
            card.move_requested.connect(self.move_device)
            item.setSizeHint(QSize(100, 92))
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, card)
        self.list_widget.blockSignals(False)

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
        for device in self.current_devices():
            if device.device_id == device_id:
                device.hidden = not device.hidden
                device.status = "Hidden" if device.hidden else "Ready"
                break
        self.refresh()

    def move_device(self, device_id: str, direction: int) -> None:
        devices = self.current_devices()
        index = next((i for i, device in enumerate(devices) if device.device_id == device_id), -1)
        new_index = index + direction
        if index < 0 or new_index < 0 or new_index >= len(devices):
            return
        devices[index], devices[new_index] = devices[new_index], devices[index]
        self.set_current_devices(devices)
        self.refresh()

    def sort_current(self) -> None:
        devices = sorted(
            self.current_devices(),
            key=lambda device: (device.hidden, device.status != "Default", -device.level),
        )
        self.set_current_devices(devices)
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
        self.refresh()

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
            QLabel#deviceSubtitle,
            QLabel#ruleDetail,
            QLabel#activityText {
                color: #9aa5b1;
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
            """
        )


def main() -> None:
    app = QApplication(sys.argv)
    window = SoundManagerWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
