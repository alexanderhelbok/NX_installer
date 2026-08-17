#!/usr/bin/env python3
"""
NX Installer GUI — PySide6 wizard for Siemens NX installation.
"""

import logging
import os
import re
import sys
import threading
from pathlib import Path
from typing import Optional

from PySide6.QtCore import (
    QObject, QThread, Signal, Slot, Qt, QTimer,
)
from PySide6.QtGui import (
    QColor, QFont, QPalette,
)
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QWidget,
    QFormLayout, QFrame, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox,  QPlainTextEdit, QProgressBar, QPushButton,
    QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QDialog,
)

if getattr(sys, 'frozen', False):
    sys.path.insert(0, str(Path(sys.executable).parent.resolve()))
else:
    sys.path.insert(0, str(Path(__file__).parent.resolve()))

from install_nx import (
    Config, FileDownloader, NXInstaller, PostInstallValidator,
    PrerequisitesInstaller, LicenseConfigurator,
    FEATURE_MAP, DEFAULT_FEATURES,
    check_admin_rights, check_license_server, find_msi,
    get_msi_features, get_msi_product_version,
    fix_nx_permissions, parse_nx_short_version, uninstall_nx,
    MSI_PRODUCT_CODE, is_nx_installed, get_msi_installed_location,
    configure_role, setup_logging,
)

STEPS = ["Configuration", "Features", "Install", "Post-Install", "Complete"]
UNINSTALL_STEPS = ["Uninstall", "Complete"]


def dark_palette() -> QPalette:
    p = QPalette()
    p.setColor(QPalette.Window, QColor(53, 53, 53))
    p.setColor(QPalette.WindowText, Qt.white)
    p.setColor(QPalette.Base, QColor(35, 35, 35))
    p.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    p.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
    p.setColor(QPalette.ToolTipText, Qt.white)
    p.setColor(QPalette.Text, Qt.white)
    p.setColor(QPalette.Button, QColor(53, 53, 53))
    p.setColor(QPalette.ButtonText, Qt.white)
    p.setColor(QPalette.BrightText, QColor(255, 100, 100))
    p.setColor(QPalette.Link, QColor(42, 130, 218))
    p.setColor(QPalette.Highlight, QColor(42, 130, 218))
    p.setColor(QPalette.HighlightedText, Qt.black)
    p.setColor(QPalette.Disabled, QPalette.WindowText, QColor(127, 127, 127))
    p.setColor(QPalette.Disabled, QPalette.Text, QColor(127, 127, 127))
    p.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(127, 127, 127))
    return p


DARK_STYLESHEET = """
QMainWindow { background-color: #353535; }
QGroupBox { border: 1px solid #555; border-radius: 4px; margin-top: 8px; padding-top: 16px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #aaa; }
QLineEdit, QComboBox { background: #2b2b2b; color: white; border: 1px solid #555; border-radius: 3px; padding: 4px 6px; }
QLineEdit:focus, QComboBox:focus { border-color: #2a82da; }
QPushButton { background: #3c3c3c; color: white; border: 1px solid #555; border-radius: 3px; padding: 6px 16px; }
QPushButton:hover { background: #4a4a4a; border-color: #2a82da; }
QPushButton:pressed { background: #2a2a2a; }
QPushButton:disabled { color: #666; }
QPushButton#primary { background: #2a82da; border-color: #2a82da; font-weight: bold; }
QPushButton#primary:hover { background: #3a92ea; border-color: #3a92ea; }
QPushButton#primary:disabled { background: #1a4a7a; border-color: #1a4a7a; }
QPushButton#danger { background: #8b3a3a; border-color: #8b3a3a; }
QPushButton#danger:hover { background: #9b4a4a; }
QProgressBar { background: #2b2b2b; border: 1px solid #555; border-radius: 3px; text-align: center; color: white; }
QProgressBar::chunk { background: #2a82da; border-radius: 2px; }
QListWidget { background: #2b2b2b; color: white; border: 1px solid #555; border-radius: 3px; outline: none; }
QListWidget::item { padding: 4px 6px; }
QListWidget::item:selected { background: #2a82da; color: white; }
QPlainTextEdit { background: #1e1e1e; color: #d4d4d4; border: 1px solid #555; border-radius: 3px; font-family: Consolas, Courier New, monospace; font-size: 12px; }
QTableWidget { background: #2b2b2b; color: white; border: 1px solid #555; gridline-color: #444; }
QHeaderView::section { background: #3c3c3c; color: white; border: 1px solid #555; padding: 4px; }
QFrame#sidebar { background: #2b2b2b; border-right: 1px solid #444; }
QFrame#step-box { border: none; margin: 0; padding: 8px; }
"""


class LogSignals(QObject):
    message = Signal(str, str)


log_signals = LogSignals()


class LogHandler(logging.Handler):
    def __init__(self, fmt: Optional[logging.Formatter] = None):
        super().__init__()
        self.setFormatter(fmt or logging.Formatter("%(message)s"))

    def emit(self, record):
        try:
            msg = self.format(record)
            log_signals.message.emit(record.levelname, msg)
        except Exception:
            pass


class InstallWorker(QObject):
    log_signal = Signal(str, str)
    stage_signal = Signal(str)
    finished = Signal(bool)

    def __init__(self, config: Config, selected: list, nx_version: str, parent=None):
        super().__init__(parent)
        self.config = config
        self.selected = selected
        self.nx_version = nx_version
        self.cancel_event = threading.Event()

    def cancel(self):
        self.cancel_event.set()

    def _log(self, level: str, msg: str):
        self.log_signal.emit(level, msg)

    def _stage(self, msg: str):
        self.stage_signal.emit(msg)
        self._log("INFO", msg)

    def run(self):
        try:
            logger = logging.getLogger("nx_install")
            handler = LogHandler()
            logger.addHandler(handler)
            handler.log_signal = self.log_signal
            try:
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Installing prerequisites...")
                prereq = PrerequisitesInstaller(
                    self.config.install_files, logger,
                    self.config.install_vcpp, self.config.install_dotnet,
                    self.config.temp_dir,
                )
                if not prereq.install_all(cancel_event=self.cancel_event):
                    if self.cancel_event.is_set():
                        self.finished.emit(False)
                        return
                    self.log_signal.emit("ERROR", "Prerequisites installation failed")
                    self.finished.emit(False)
                    return

                self._stage("Installing NX...")
                installer = NXInstaller(self.config, logger)
                if not installer.install(self.selected, cancel_event=self.cancel_event):
                    if self.cancel_event.is_set():
                        self.finished.emit(False)
                        return
                    self.log_signal.emit("ERROR", "NX installation failed")
                    self.finished.emit(False)
                    return
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Fixing permissions...")
                fix_nx_permissions(self.config.install_dir, logger)

                self.finished.emit(True)
            finally:
                logger.removeHandler(handler)
        except BaseException as e:
            self.log_signal.emit("ERROR", str(e))
            self.finished.emit(False)


class PostInstallWorker(QObject):
    log_signal = Signal(str, str)
    stage_signal = Signal(str)
    progress_signal = Signal(str, int, int)
    finished = Signal(bool)
    validation_signal = Signal(dict)

    def __init__(self, config: Config, nx_version: str, parent=None):
        super().__init__(parent)
        self.config = config
        self.nx_version = nx_version
        self.cancel_event = threading.Event()

    def cancel(self):
        self.cancel_event.set()

    def _log(self, level: str, msg: str):
        self.log_signal.emit(level, msg)

    def _stage(self, msg: str):
        self.stage_signal.emit(msg)
        self._log("INFO", msg)

    def _progress(self, fname: str, downloaded: int, total: int):
        self.progress_signal.emit(fname, downloaded, total)

    def run(self):
        try:
            logger = logging.getLogger("nx_install")
            handler = LogHandler()
            logger.addHandler(handler)
            handler.log_signal = self.log_signal
            try:
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Configuring license server...")
                lc = LicenseConfigurator(self.config, logger)
                if not lc.configure():
                    self._log("ERROR", "License configuration failed")
                    self.finished.emit(False)
                    return
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                dl = FileDownloader(self.config, logger)

                _progress_cb = lambda f, d, t: self._progress(f, d, t)

                self._stage("Downloading fcc.xml...")
                fcc = dl.download(self.config.fcc_url, "fcc.xml", _progress_cb, cancel_event=self.cancel_event)
                if not fcc:
                    if self.cancel_event.is_set():
                        self.finished.emit(False)
                        return
                    self._log("ERROR", "fcc.xml download failed")
                    self.finished.emit(False)
                    return
                target_fcc = Path(self.config.install_dir) / "UGMANAGER" / "tccs" / "fcc.xml"
                dl.move(fcc, target_fcc)
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Downloading java.zip...")
                target_java = Path(self.config.install_dir).parent / "java"
                if not target_java.exists():
                    java = dl.download(self.config.java_url, "java.zip", _progress_cb, cancel_event=self.cancel_event)
                    if not java:
                        if self.cancel_event.is_set():
                            self.finished.emit(False)
                            return
                        self._log("ERROR", "java.zip download failed")
                        self.finished.emit(False)
                        return
                    self._stage("Extracting java...")
                    java_extracted = Path(self.config.temp_dir) / "java_extracted"
                    dl.unzip(java, java_extracted)
                    fix_nx_permissions(str(target_java.parent), logger)
                    dl.move(java_extracted, target_java / "zulu17")
                else:
                    self._log("INFO", f"Java already exists at {target_java}, skipping")
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Downloading start_nx.bat...")
                start_nx = dl.download(self.config.start_nx_url, "start_nx.bat", _progress_cb, cancel_event=self.cancel_event)
                if not start_nx:
                    if self.cancel_event.is_set():
                        self.finished.emit(False)
                        return
                    self._log("ERROR", "start_nx.bat download failed")
                    self.finished.emit(False)
                    return

                def transform_start_nx(content: str) -> str:
                    name = self.config.name
                    surname = self.config.surname
                    if not name or not surname:
                        raise ValueError("name and surname must be set in config.ini")
                    username = f"{name}.{surname}"
                    password = f"{name}00"
                    install_dir = self.config.install_dir.rstrip("\\")
                    java_home = str(target_java / "zulu17")
                    result = []
                    for line in content.splitlines():
                        if line.startswith("set JAVA_HOME="):
                            result.append(f"set JAVA_HOME={java_home}")
                        elif line.startswith("set NX_HOME="):
                            result.append(f"set NX_HOME={install_dir}")
                        elif 'ugraf.exe' in line and '-u=' not in line:
                            idx = line.find('ugraf.exe"')
                            result.append(line[:idx + 10] + f' -u={username} -p={password} ' + line[idx + 10:])
                        else:
                            result.append(line)
                    return '\n'.join(result)

                target_start_nx = Path(self.config.install_dir).parent / "start_nx.bat"
                dl.transform(start_nx, target_start_nx, transform_start_nx)
                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Configuring role and preferences...")
                configure_role(self.config, logger, self.nx_version, cancel_event=self.cancel_event)

                if self.cancel_event.is_set():
                    self.finished.emit(False)
                    return

                self._stage("Running validation...")
                validator = PostInstallValidator(self.config, logger, self.nx_version)
                all_ok = validator.validate()
                self.validation_signal.emit(validator.results)

                self.finished.emit(all_ok)
            finally:
                logger.removeHandler(handler)
        except BaseException as e:
            self._log("ERROR", str(e))
            self.finished.emit(False)


class UninstallWorker(QObject):
    log_signal = Signal(str, str)
    stage_signal = Signal(str)
    finished = Signal(bool)

    def __init__(self, config: Config, parent=None):
        super().__init__(parent)
        self.config = config

    def _log(self, level: str, msg: str):
        self.log_signal.emit(level, msg)

    def _stage(self, msg: str):
        self.stage_signal.emit(msg)
        self._log("INFO", msg)

    def run(self):
        try:
            logger = logging.getLogger("nx_install")
            handler = LogHandler()
            logger.addHandler(handler)
            handler.log_signal = self.log_signal
            try:
                self._stage("Starting uninstall...")
                success = uninstall_nx(logger=logger, timeout=self.config.install_timeout, config=self.config)
                if success:
                    self._stage("Uninstall completed successfully.")
                else:
                    self._log("ERROR", "Uninstall failed. Check the log for details.")
                self.finished.emit(success)
            finally:
                logger.removeHandler(handler)
        except BaseException as e:
            self.log_signal.emit("ERROR", str(e))
            self.finished.emit(False)


class UninstallPage(QWidget):
    finished = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker = None
        self._thread = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Uninstalling NX...")
        title.setObjectName("uninstall_title")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 8px;")
        layout.addWidget(title)

        self.stage_label = QLabel("Waiting to start...")
        self.stage_label.setStyleSheet("color: #aaa; margin-bottom: 8px;")
        layout.addWidget(self.stage_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        layout.addWidget(self.progress_bar)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        layout.addWidget(self.log_view, 1)

    def start_uninstall(self, config: Config):
        self.log_view.clear()
        self.stage_label.setText("Starting uninstall...")
        self.progress_bar.setRange(0, 0)

        self._thread = QThread()
        self._worker = UninstallWorker(config)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log_signal.connect(self._on_log)
        self._worker.stage_signal.connect(self._on_stage)
        self._worker.finished.connect(self._on_uninstall_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()

    @Slot(str, str)
    def _on_log(self, level: str, msg: str):
        color_map = {
            "INFO": "#4ec94e",
            "DEBUG": "#888",
            "WARNING": "#e8c547",
            "ERROR": "#e84e4e",
            "CRITICAL": "#e84e4e",
        }
        color = color_map.get(level, "white")
        self.log_view.appendHtml(f'<span style="color:{color}">{msg}</span>')
        scrollbar = self.log_view.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @Slot(str)
    def _on_stage(self, msg: str):
        self.stage_label.setText(msg)

    @Slot(bool)
    def _on_uninstall_finished(self, success: bool):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100 if success else 0)
        if success:
            self.stage_label.setText("Uninstall complete!")
            self.stage_label.setStyleSheet("color: #4ec94e; font-weight: bold;")
            title = self.findChild(QLabel, "uninstall_title")
            if title:
                title.setText("Uninstall Complete")
        else:
            self.stage_label.setText("Uninstall failed!")
            self.stage_label.setStyleSheet("color: #e84e4e; font-weight: bold;")
        self.finished.emit(success)


class ConfigPage(QWidget):
    next_requested = Signal()
    version_changed = Signal(str)

    def __init__(self, config: Config, parent=None):
        super().__init__(parent)
        self.config = config
        self._setup_ui()
        self._load_config()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Configuration")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 8px;")
        layout.addWidget(title)

        desc = QLabel("Review and edit installation settings before proceeding.")
        desc.setStyleSheet("color: #aaa; margin-bottom: 16px;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        form = QFormLayout()
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight)

        self.media_edit = QLineEdit()
        self.media_browse = QPushButton("Browse...")
        self.media_browse.clicked.connect(self._on_browse_media)
        media_row = QHBoxLayout()
        media_row.addWidget(self.media_edit, 1)
        media_row.addWidget(self.media_browse)
        form.addRow("Media path:", media_row)

        self.install_edit = QLineEdit()
        self.install_browse = QPushButton("Browse...")
        self.install_browse.clicked.connect(self._on_browse_install)
        install_row = QHBoxLayout()
        install_row.addWidget(self.install_edit, 1)
        install_row.addWidget(self.install_browse)
        form.addRow("Install dir:", install_row)

        self.temp_edit = QLineEdit()
        form.addRow("Temp dir:", self.temp_edit)

        self.license_edit = QLineEdit()
        self.license_test = QPushButton("Test")
        self.license_test.clicked.connect(self._on_test_license)
        license_row = QHBoxLayout()
        license_row.addWidget(self.license_edit, 1)
        license_row.addWidget(self.license_test)
        form.addRow("License server:", license_row)

        self.name_edit = QLineEdit()
        self.surname_edit = QLineEdit()
        name_row = QHBoxLayout()
        name_row.addWidget(self.name_edit, 1)
        name_row.addWidget(self.surname_edit, 1)
        form.addRow("Name / Surname:", name_row)

        self.vcpp_cb = QCheckBox("Install VC++ Redistributable")
        self.vcpp_cb.setChecked(True)
        form.addRow("", self.vcpp_cb)

        self.dotnet_cb = QCheckBox("Install .NET Framework 4.8")
        self.dotnet_cb.setChecked(True)
        form.addRow("", self.dotnet_cb)

        layout.addLayout(form)
        layout.addStretch()

        self.version_label = QLabel()
        self.version_label.setStyleSheet("color: #888; font-size: 11px;")
        layout.addWidget(self.version_label)

    def _load_config(self):
        self.media_edit.setText(self.config.install_files)
        self.install_edit.setText(self.config.install_dir)
        self.temp_edit.setText(self.config.temp_dir)
        self.license_edit.setText(self.config.license_server)
        self.name_edit.setText(self.config.name)
        self.surname_edit.setText(self.config.surname)
        self.vcpp_cb.setChecked(self.config.install_vcpp)
        self.dotnet_cb.setChecked(self.config.install_dotnet)
        self._update_version()

    def _update_version(self):
        msi = find_msi(self.media_edit.text())
        ver = ""
        if msi:
            raw = get_msi_product_version(msi, self.config.temp_dir)
            if raw:
                ver = parse_nx_short_version(raw)
        if not ver:
            m = re.search(r"SiemensNX[._-](\d{4})", self.media_edit.text())
            if m:
                ver = m.group(1)
            else:
                ver = "?"
        self.version_label.setText(f"Detected NX version: {ver}")
        self._detected_version = ver
        if ver and ver != "?":
            self.version_changed.emit(ver)

    def get_version(self) -> str:
        return getattr(self, "_detected_version", "")

    def validate(self) -> bool:
        fields = {
            "Media path": self.media_edit,
            "Install dir": self.install_edit,
            "Temp dir": self.temp_edit,
            "License server": self.license_edit,
            "Name": self.name_edit,
            "Surname": self.surname_edit,
        }
        empty = []
        for label, field in fields.items():
            if not field.text().strip():
                empty.append(label)
                field.setStyleSheet("border: 1px solid #e74c3c;")
            else:
                field.setStyleSheet("")
        if empty:
            QMessageBox.warning(self, "Required Fields",
                                f"Please fill in: {', '.join(empty)}")
            return False
        return True

    def save_config(self):
        self.config.install_files = self.media_edit.text().strip()
        self.config.install_dir = self.install_edit.text().strip()
        self.config.temp_dir = self.temp_edit.text().strip()
        self.config.license_server = self.license_edit.text().strip()
        self.config.name = self.name_edit.text().strip().lower()
        self.config.surname = self.surname_edit.text().strip().lower()
        self.config.install_vcpp = self.vcpp_cb.isChecked()
        self.config.install_dotnet = self.dotnet_cb.isChecked()

    def _on_browse_media(self):
        d = QFileDialog.getExistingDirectory(self, "Select media directory", self.media_edit.text())
        if d:
            self.media_edit.setText(d)
            self._update_version()

    def _on_browse_install(self):
        d = QFileDialog.getExistingDirectory(self, "Select install directory", self.install_edit.text())
        if d:
            self.install_edit.setText(d)

    def _on_test_license(self):
        self.license_test.setEnabled(False)
        self.license_test.setText("Testing...")

        class TestThread(QThread):
            result = Signal(bool)

            def __init__(self, server, parent=None):
                super().__init__(parent)
                self.server = server

            def run(self):
                import logging
                ok = check_license_server(self.server, timeout=5, logger=logging.getLogger())
                self.result.emit(ok)

        self._test_thread = TestThread(self.license_edit.text().strip())
        self._test_thread.result.connect(self._on_test_result)
        self._test_thread.start()

    def _on_test_result(self, ok: bool):
        self.license_test.setEnabled(True)
        if ok:
            self.license_test.setText("OK")
            self.license_test.setStyleSheet("background: #2d6a2d; border-color: #2d6a2d; color: white;")
        else:
            self.license_test.setText("Failed")
            self.license_test.setStyleSheet("background: #8b3a3a; border-color: #8b3a3a; color: white;")
        QTimer.singleShot(3000, self._reset_license_test)

    def _reset_license_test(self):
        self.license_test.setText("Test")
        self.license_test.setStyleSheet("")


class FeaturePage(QWidget):
    next_requested = Signal()

    def __init__(self, config: Config, parent=None):
        super().__init__(parent)
        self.config = config
        self.msi_features: set = set()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Feature Selection")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 8px;")
        layout.addWidget(title)

        desc = QLabel("Choose which NX components to install.")
        desc.setStyleSheet("color: #aaa; margin-bottom: 8px;")
        layout.addWidget(desc)

        filter_row = QHBoxLayout()
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filter features...")
        self.filter_edit.textChanged.connect(self._on_filter)
        filter_row.addWidget(self.filter_edit, 1)
        layout.addLayout(filter_row)

        btn_row = QHBoxLayout()
        self.sel_all_btn = QPushButton("Select All")
        self.sel_all_btn.clicked.connect(self._on_select_all)
        self.defaults_btn = QPushButton("Defaults")
        self.defaults_btn.clicked.connect(self._on_defaults)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self._on_clear)
        btn_row.addWidget(self.sel_all_btn)
        btn_row.addWidget(self.defaults_btn)
        btn_row.addWidget(self.clear_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.feature_list = QListWidget()
        self.feature_list.setAlternatingRowColors(True)
        layout.addWidget(self.feature_list, 1)

        self.count_label = QLabel("0 features loaded")
        self.count_label.setStyleSheet("color: #888; font-size: 11px;")
        layout.addWidget(self.count_label)

    def load_features(self, msi_path: str):
        self.msi_features = get_msi_features(msi_path, self.config.temp_dir)
        self.feature_list.clear()
        matched = sorted(
            fid for fid in self.msi_features if fid in FEATURE_MAP
        )
        for fid in matched:
            item = QListWidgetItem(f"{fid:<30} {FEATURE_MAP.get(fid, fid)}")
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if fid in DEFAULT_FEATURES else Qt.Unchecked)
            item.setData(Qt.UserRole, fid)
            self.feature_list.addItem(item)
        self._update_count()

    def _update_count(self):
        total = self.feature_list.count()
        checked = sum(
            1 for i in range(self.feature_list.count())
            if self.feature_list.item(i).checkState() == Qt.Checked
        )
        self.count_label.setText(f"{checked} / {total} selected")

    def _on_filter(self, text: str):
        for i in range(self.feature_list.count()):
            item = self.feature_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def _on_select_all(self):
        for i in range(self.feature_list.count()):
            if not self.feature_list.item(i).isHidden():
                self.feature_list.item(i).setCheckState(Qt.Checked)
        self._update_count()

    def _on_defaults(self):
        for i in range(self.feature_list.count()):
            fid = self.feature_list.item(i).data(Qt.UserRole)
            checked = Qt.Checked if fid in DEFAULT_FEATURES else Qt.Unchecked
            self.feature_list.item(i).setCheckState(checked)
        self._update_count()

    def _on_clear(self):
        for i in range(self.feature_list.count()):
            self.feature_list.item(i).setCheckState(Qt.Unchecked)
        self._update_count()

    def get_selected(self) -> list:
        return sorted(
            self.feature_list.item(i).data(Qt.UserRole)
            for i in range(self.feature_list.count())
            if self.feature_list.item(i).checkState() == Qt.Checked
        )


class InstallPage(QWidget):
    finished = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker = None
        self._thread = None
        self._config = None
        self._selected = None
        self._nx_version = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Installing NX...")
        title.setObjectName("status_title")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 8px;")
        layout.addWidget(title)

        self.stage_label = QLabel("Ready to start installation.")
        self.stage_label.setStyleSheet("color: #aaa; margin-bottom: 8px;")
        layout.addWidget(self.stage_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        layout.addWidget(self.log_view, 1)

        self.cancel_btn = QPushButton("Cancel Installation")
        self.cancel_btn.setObjectName("danger")
        self.cancel_btn.clicked.connect(self._on_cancel)
        self.cancel_btn.setVisible(False)
        layout.addWidget(self.cancel_btn, alignment=Qt.AlignRight)

        self.skip_btn = QPushButton("Skip to Post-Install (NX already installed)")
        self.skip_btn.setObjectName("primary")
        self.skip_btn.clicked.connect(self._on_skip)
        self.skip_btn.setVisible(False)
        layout.addWidget(self.skip_btn, alignment=Qt.AlignRight)

        self.start_btn = QPushButton("Start Installation")
        self.start_btn.setObjectName("primary")
        self.start_btn.clicked.connect(self._on_start_click)
        layout.addWidget(self.start_btn, alignment=Qt.AlignRight)

    def set_install_params(self, config, selected, nx_version):
        self._config = config
        self._selected = selected
        self._nx_version = nx_version
        nx_installed = is_nx_installed()
        self.skip_btn.setVisible(nx_installed)
        self.start_btn.setText("Reinstall NX" if nx_installed else "Start Installation")

    def _on_start_click(self):
        self.start_install(self._config, self._selected, self._nx_version)

    def _on_skip(self):
        self.log_view.clear()
        self.stage_label.setText("Skipping install — NX already present.")
        self.stage_label.setStyleSheet("color: #4ec94e; font-weight: bold;")
        title = self.findChild(QLabel, "status_title")
        if title:
            title.setText("Installation Skipped")
        self.start_btn.setVisible(False)
        self.skip_btn.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.finished.emit(True)

    def reset_ui(self):
        self.log_view.clear()
        self.stage_label.setText("Ready to start installation.")
        self.stage_label.setStyleSheet("color: #aaa;")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        nx_installed = is_nx_installed()
        self.start_btn.setText("Reinstall NX" if nx_installed else "Start Installation")
        self.start_btn.setVisible(True)
        self.skip_btn.setVisible(nx_installed)
        self.cancel_btn.setVisible(False)
        title = self.findChild(QLabel, "status_title")
        if title:
            title.setText("Installing NX...")

    def start_install(self, config: Config, selected: list, nx_version: str):
        self.log_view.clear()
        self.stage_label.setText("Starting installation...")
        self.progress_bar.setRange(0, 0)
        self.start_btn.setVisible(False)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel Installation")

        self._thread = QThread()
        self._worker = InstallWorker(config, selected, nx_version)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log_signal.connect(self._on_log)
        self._worker.stage_signal.connect(self._on_stage)
        self._worker.finished.connect(self._on_install_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()

    @Slot(str, str)
    def _on_log(self, level: str, msg: str):
        color_map = {
            "INFO": "#4ec94e",
            "DEBUG": "#888",
            "WARNING": "#e8c547",
            "ERROR": "#e84e4e",
            "CRITICAL": "#e84e4e",
        }
        color = color_map.get(level, "white")
        self.log_view.appendHtml(f'<span style="color:{color}">{msg}</span>')
        scrollbar = self.log_view.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @Slot(str)
    def _on_stage(self, msg: str):
        self.stage_label.setText(msg)

    def _on_cancel(self):
        if self._worker:
            self._worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Cancelling...")
        self._on_log("WARNING", "Cancellation requested, stopping...")
        self.stage_label.setText("Cancelling...")
        self.stage_label.setStyleSheet("color: #e8c547; font-weight: bold;")

    @Slot(bool)
    def _on_install_finished(self, success: bool):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100 if success else 0)
        self.cancel_btn.setVisible(False)
        if success:
            self.stage_label.setText("Installation complete!")
            self.stage_label.setStyleSheet("color: #4ec94e; font-weight: bold;")
            title = self.findChild(QLabel, "status_title")
            if title:
                title.setText("Installation Complete")
            self.start_btn.setVisible(False)
        else:
            self.stage_label.setText("Installation failed!")
            self.stage_label.setStyleSheet("color: #e84e4e; font-weight: bold;")
            self.start_btn.setText("Retry Installation")
            self.start_btn.setVisible(True)
        self.finished.emit(success)


class PostInstallPage(QWidget):
    finished = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._worker = None
        self._thread = None
        self._dl_bars: dict = {}
        self._config = None
        self._nx_version = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Post-Install Configuration")
        title.setObjectName("post_title")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: white; margin-bottom: 8px;")
        layout.addWidget(title)

        self.stage_label = QLabel("Ready for post-install configuration.")
        self.stage_label.setStyleSheet("color: #aaa; margin-bottom: 8px;")
        layout.addWidget(self.stage_label)

        self.dl_group = QGroupBox("Downloads")
        self.dl_layout = QVBoxLayout(self.dl_group)
        self.dl_group.setVisible(False)
        layout.addWidget(self.dl_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        layout.addWidget(self.log_view, 1)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("danger")
        self.cancel_btn.clicked.connect(self._on_cancel)
        self.cancel_btn.setVisible(False)
        layout.addWidget(self.cancel_btn, alignment=Qt.AlignRight)

        self.start_btn = QPushButton("Start Post-Install")
        self.start_btn.setObjectName("primary")
        self.start_btn.clicked.connect(self._on_start_click)
        layout.addWidget(self.start_btn, alignment=Qt.AlignRight)

    def set_postinstall_params(self, config, nx_version):
        self._config = config
        self._nx_version = nx_version

    def _on_start_click(self):
        self.start_postinstall(self._config, self._nx_version)

    def reset_ui(self):
        self.log_view.clear()
        self.stage_label.setText("Ready for post-install configuration.")
        self.stage_label.setStyleSheet("color: #aaa;")
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.start_btn.setText("Start Post-Install")
        self.start_btn.setVisible(True)
        self.cancel_btn.setVisible(False)
        self.dl_group.setVisible(False)

    def _add_dl_bar(self, name: str):
        row = QHBoxLayout()
        label = QLabel(name)
        label.setFixedWidth(160)
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setValue(0)
        bar.setTextVisible(True)
        bar.setFixedHeight(20)
        row.addWidget(label)
        row.addWidget(bar, 1)
        self.dl_layout.addLayout(row)
        self._dl_bars[name] = bar

    def start_postinstall(self, config: Config, nx_version: str):
        self.log_view.clear()
        self.stage_label.setText("Starting post-install...")
        self.progress_bar.setRange(0, 0)
        self.start_btn.setVisible(False)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel")
        self.dl_group.setVisible(True)

        while self.dl_layout.count():
            item = self.dl_layout.takeAt(0)
            if item.layout():
                while item.layout().count():
                    child = item.layout().takeAt(0)
                    if child.widget():
                        child.widget().deleteLater()
                item.layout().deleteLater()
        self._dl_bars.clear()

        for dname in ["fcc.xml", "java.zip", "start_nx.bat", "user.mtx", "NX_user.dpv", "feature_toggle_user.fcg"]:
            self._add_dl_bar(dname)

        self._thread = QThread()
        self._worker = PostInstallWorker(config, nx_version)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log_signal.connect(self._on_log)
        self._worker.stage_signal.connect(self._on_stage)
        self._worker.progress_signal.connect(self._on_progress)
        self._worker.finished.connect(self._on_post_finished)
        self._worker.validation_signal.connect(self._on_validation)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()

    @Slot(str, str)
    def _on_log(self, level: str, msg: str):
        color_map = {
            "INFO": "#4ec94e",
            "DEBUG": "#888",
            "WARNING": "#e8c547",
            "ERROR": "#e84e4e",
        }
        color = color_map.get(level, "white")
        self.log_view.appendHtml(f'<span style="color:{color}">{msg}</span>')
        scrollbar = self.log_view.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @Slot(str)
    def _on_stage(self, msg: str):
        self.stage_label.setText(msg)

    @Slot(str, int, int)
    def _on_progress(self, fname: str, downloaded: int, total: int):
        bar = self._dl_bars.get(fname)
        if bar:
            pct = int(downloaded / max(total, 1) * 100)
            bar.setValue(pct)
            bar.setFormat(f"{pct}% ({downloaded // 1024}KB / {total // 1024}KB)")

    def _on_cancel(self):
        if self._worker:
            self._worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setText("Cancelling...")
        self._on_log("WARNING", "Cancellation requested, stopping...")
        self.stage_label.setText("Cancelling...")
        self.stage_label.setStyleSheet("color: #e8c547; font-weight: bold;")

    @Slot(bool)
    def _on_post_finished(self, success: bool):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.cancel_btn.setVisible(False)
        if success:
            self.stage_label.setText("All post-install steps completed!")
            self.stage_label.setStyleSheet("color: #4ec94e; font-weight: bold;")
            self.start_btn.setVisible(False)
        else:
            self.stage_label.setText("Post-install failed!")
            self.stage_label.setStyleSheet("color: #e84e4e; font-weight: bold;")
            self.start_btn.setText("Retry Post-Install")
            self.start_btn.setVisible(True)
        self.finished.emit(success)

    @Slot(dict)
    def _on_validation(self, results: dict):
        self.validation_results = results


class LogViewerDialog(QDialog):
    def __init__(self, log_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Installation Log")
        self.setMinimumSize(800, 500)
        self.resize(900, 600)
        self.setStyleSheet("background: #1e1e1e; color: #d4d4d4;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setMaximumBlockCount(10000)
        self.text_edit.setStyleSheet(
            "QPlainTextEdit { background: #1e1e1e; color: #d4d4d4; border: 1px solid #333; "
            "font-family: Consolas, Courier New, monospace; font-size: 12px; }"
        )
        layout.addWidget(self.text_edit, 1)

        btn_row = QHBoxLayout()
        self.close_btn = QPushButton("Close")
        self.close_btn.setObjectName("primary")
        self.close_btn.clicked.connect(self.accept)

        line_count_label = QLabel("")
        line_count_label.setStyleSheet("color: #888;")

        btn_row.addWidget(line_count_label)
        btn_row.addStretch()
        btn_row.addWidget(self.close_btn)
        layout.addLayout(btn_row)

        self._load_log(log_path)
        line_count_label.setText(f"{self.text_edit.blockCount()} lines")

        scrollbar = self.text_edit.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @staticmethod
    def _ansi_to_html(text: str) -> str:
        ansi_color_map = {
            "0": None,
            "30": "#000", "31": "#e84e4e", "32": "#4ec94e",
            "33": "#e8c547", "34": "#2a82da", "35": "#c586c0",
            "36": "#888", "37": "#d4d4d4",
            "90": "#888", "91": "#e84e4e", "92": "#4ec94e",
            "93": "#e8c547", "94": "#2a82da", "95": "#c586c0",
            "96": "#888", "97": "#d4d4d4",
        }
        parts = re.split(r'(\x1B\[[0-9;]*m)', text)
        current_color = None
        html_parts = []
        for part in parts:
            m = re.match(r'\x1B\[([0-9;]*)m', part)
            if m:
                codes = m.group(1).split(";") if m.group(1) else ["0"]
                for code in codes:
                    if code in ("", "0"):
                        current_color = None
                    elif code == "1":
                        pass
                    elif code in ansi_color_map:
                        current_color = ansi_color_map[code]
                continue
            if not part:
                continue
            escaped = part.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if current_color:
                html_parts.append(f'<span style="color:{current_color}">{escaped}</span>')
            else:
                html_parts.append(escaped)
        return "".join(html_parts)

    def _load_log(self, log_path: str):
        try:
            with open(log_path, encoding="utf-8", errors="replace") as f:
                for line in f:
                    line = line.rstrip("\n\r")
                    html = self._ansi_to_html(line)
                    self.text_edit.appendHtml(html)
        except Exception as e:
            self.text_edit.setPlainText(f"Failed to read log: {e}")


class CompletePage(QWidget):
    def __init__(self, nx_version: str, parent=None):
        super().__init__(parent)
        self.nx_version = nx_version
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)

        header_row = QHBoxLayout()
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(48, 48)
        header_row.addWidget(self.icon_label)

        header_text = QVBoxLayout()
        self.status_title = QLabel("Installation Complete")
        self.status_title.setStyleSheet("font-size: 20px; font-weight: bold; color: white;")
        header_text.addWidget(self.status_title)
        self.status_sub = QLabel("")
        self.status_sub.setStyleSheet("color: #aaa;")
        header_text.addWidget(self.status_sub)
        header_row.addLayout(header_text, 1)
        layout.addLayout(header_row)

        info_group = QGroupBox("Installation Info")
        info_layout = QFormLayout(info_group)
        self.info_table = QTableWidget(0, 2)
        self.info_table.setColumnCount(2)
        self.info_table.setHorizontalHeaderLabels(["Check", "Result"])
        self.info_table.horizontalHeader().setStretchLastSection(True)
        self.info_table.setColumnWidth(0, 200)
        info_layout.addWidget(self.info_table)
        layout.addWidget(info_group)

        btn_row = QHBoxLayout()
        self.log_btn = QPushButton("View Log")
        self.log_btn.clicked.connect(self._on_view_log)
        self.finish_btn = QPushButton("Finish")
        self.finish_btn.setObjectName("primary")
        self.finish_btn.clicked.connect(self._on_finish)
        btn_row.addStretch()
        btn_row.addWidget(self.log_btn)
        btn_row.addWidget(self.finish_btn)
        layout.addLayout(btn_row)

    def show_results(self, success: bool, config: Config, results: dict, operation: str = "Installation"):
        if success:
            self.icon_label.setText("✅")
            self.status_title.setText(f"{operation} Complete")
            if operation == "Uninstall":
                self.status_sub.setText("NX has been removed from this system.")
            else:
                self.status_sub.setText(f"NX {self.nx_version} installed at {config.install_dir}")
            self.status_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #4ec94e;")
        else:
            self.icon_label.setText("❌")
            self.status_title.setText(f"{operation} Failed")
            self.status_sub.setText("One or more steps failed. Check the log for details.")
            self.status_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #e84e4e;")

        self.info_table.setRowCount(len(results))
        for i, (check, passed) in enumerate(sorted(results.items())):
            status = "PASS" if passed else "FAIL"
            self.info_table.setItem(i, 0, QTableWidgetItem(check))
            item = QTableWidgetItem(status)
            if passed:
                item.setForeground(QColor("#4ec94e"))
            else:
                item.setForeground(QColor("#e84e4e"))
            self.info_table.setItem(i, 1, item)

        self._config = config

    def _on_view_log(self):
        from install_nx import LOG_FILE
        if LOG_FILE and hasattr(LOG_FILE, 'baseFilename'):
            log_path = LOG_FILE.baseFilename
            if Path(log_path).exists():
                dlg = LogViewerDialog(log_path, self)
                dlg.exec()
                return
        QMessageBox.information(self, "Log", "No log file available.")

    def _on_finish(self):
        QApplication.instance().quit()


class NXInstallWizard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config: Optional[Config] = None
        self.nx_version = ""
        self.selected_features: list = []
        self._pages: list[QWidget] = []
        self._current_page = 0
        self._mode = "install"
        self._install_state = "pending"
        self._postinstall_state = "pending"

        self._setup_ui()
        self._load_initial()

    def _setup_ui(self):
        self.setWindowTitle("Siemens NX Installer")
        self.setMinimumSize(900, 680)
        self.resize(960, 720)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(200)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(12, 20, 12, 20)
        self.sidebar_layout.setSpacing(2)

        sidebar_title = QLabel("NX Installer")
        sidebar_title.setStyleSheet("font-size: 16px; font-weight: bold; color: white; margin-bottom: 16px;")
        self.sidebar_layout.addWidget(sidebar_title)

        self.steps_container = QVBoxLayout()
        self.steps_container.setSpacing(2)
        self.sidebar_layout.addLayout(self.steps_container)

        self.step_labels: list[QPushButton] = []

        self.sidebar_layout.addStretch()
        self.version_info = QLabel()
        self.version_info.setStyleSheet("color: #555; font-size: 10px; padding: 4px;")
        self.sidebar_layout.addWidget(self.version_info)
        main_layout.addWidget(self.sidebar)

        content_area = QVBoxLayout()
        content_area.setContentsMargins(0, 0, 0, 0)
        content_area.setSpacing(0)

        self.stack = QStackedWidget()
        content_area.addWidget(self.stack, 1)

        nav_bar = QFrame()
        nav_bar.setStyleSheet("QFrame { background: #2b2b2b; border-top: 1px solid #444; }")
        nav_layout = QHBoxLayout(nav_bar)
        nav_layout.setContentsMargins(16, 10, 16, 10)

        self.back_btn = QPushButton("< Back")
        self.back_btn.clicked.connect(self._on_back)
        nav_layout.addWidget(self.back_btn)

        nav_layout.addStretch()

        self.next_btn = QPushButton("Next >")
        self.next_btn.setObjectName("primary")
        self.next_btn.clicked.connect(self._on_next)
        nav_layout.addWidget(self.next_btn)

        content_area.addWidget(nav_bar)
        main_layout.addLayout(content_area, 1)

    def _rebuild_sidebar_steps(self, steps: list[str]):
        while self.steps_container.count():
            item = self.steps_container.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.step_labels = []
        for i, name in enumerate(steps):
            btn = QPushButton(f"  {i + 1}.  {name}")
            btn.setObjectName("step-label")
            btn.setFlat(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                "QPushButton { padding: 8px 12px; border-radius: 4px; color: #888; font-size: 13px;"
                " text-align: left; border: none; }"
                "QPushButton:hover { background: #3a3a3a; }"
            )
            btn.clicked.connect(lambda checked, idx=i: self._navigate_to(idx))
            self.steps_container.addWidget(btn)
            self.step_labels.append(btn)

    def _load_initial(self):
        if getattr(sys, 'frozen', False):
            exe_dir = Path(sys.executable).parent.resolve()
            embedded = Path(sys._MEIPASS) / "config.ini"
            override = exe_dir / "config.ini"
            config_path = override if override.exists() else embedded
        else:
            config_path = Path(__file__).parent.resolve() / "config.ini"
        if not config_path.exists():
            QMessageBox.critical(self, "Error", f"Config not found: {config_path}")
            sys.exit(1)

        self.config = Config(str(config_path))

        if not check_admin_rights():
            QMessageBox.critical(self, "Error", "Must be run as Administrator.")
            sys.exit(1)

        msi = find_msi(self.config.install_files)
        if msi:
            raw = get_msi_product_version(msi, self.config.temp_dir)
            if raw:
                self.nx_version = parse_nx_short_version(raw)
        if not self.nx_version:
            m = re.search(r"SiemensNX[._-](\d{4})", self.config.install_files)
            self.nx_version = m.group(1) if m else "2506"

        self.version_info.setText(f"NX {self.nx_version}")
        self.setWindowTitle(f"Siemens NX {self.nx_version} Installer")

        if is_nx_installed():
            msg = QMessageBox(self)
            msg.setWindowTitle("NX Installation Detected")
            msg.setText("NX is already installed on this system.")
            btn_install = msg.addButton("Reinstall / Upgrade", QMessageBox.ActionRole)
            btn_uninstall = msg.addButton("Uninstall", QMessageBox.ActionRole)
            btn_cancel = msg.addButton("Cancel", QMessageBox.RejectRole)
            msg.exec()
            if msg.clickedButton() == btn_uninstall:
                self._mode = "uninstall"
            elif msg.clickedButton() == btn_cancel:
                sys.exit(0)

        self._rebuild_sidebar_steps(STEPS if self._mode == "install" else UNINSTALL_STEPS)

        if self._mode == "uninstall":
            self.uninstall_page = UninstallPage()
            self.stack.addWidget(self.uninstall_page)
            self._pages.append(self.uninstall_page)

            self.complete_page = CompletePage(self.nx_version)
            self.stack.addWidget(self.complete_page)
            self._pages.append(self.complete_page)

            self.uninstall_page.finished.connect(self._on_uninstall_done)
            self._show_page(0)
            QTimer.singleShot(500, lambda: self.uninstall_page.start_uninstall(self.config))
        else:
            self.config_page = ConfigPage(self.config)
            self.config_page.version_changed.connect(self._on_version_changed)
            self.stack.addWidget(self.config_page)
            self._pages.append(self.config_page)

            self.feature_page = FeaturePage(self.config)
            self.stack.addWidget(self.feature_page)
            self._pages.append(self.feature_page)

            self.install_page = InstallPage()
            self.stack.addWidget(self.install_page)
            self._pages.append(self.install_page)

            self.postinstall_page = PostInstallPage()
            self.stack.addWidget(self.postinstall_page)
            self._pages.append(self.postinstall_page)

            self.complete_page = CompletePage(self.nx_version)
            self.stack.addWidget(self.complete_page)
            self._pages.append(self.complete_page)

            self.install_page.finished.connect(self._on_install_done)
            self.postinstall_page.finished.connect(self._on_postinstall_done)

            self._show_page(0)

    def _show_page(self, idx: int):
        self._current_page = idx
        self.stack.setCurrentIndex(idx)
        total_pages = len(self._pages)
        if self._mode == "uninstall":
            self.back_btn.setEnabled(idx > 0)
            self.next_btn.setVisible(False)
        else:
            self.back_btn.setEnabled(idx > 0)
            if idx == 0:
                self.next_btn.setText("Next >")
                self.next_btn.setVisible(True)
            elif idx == 1:
                self.next_btn.setText("Next >")
                self.next_btn.setVisible(True)
            elif idx == 2:
                if self._install_state == "completed":
                    self.next_btn.setText("Continue >")
                    self.next_btn.setVisible(True)
                else:
                    self.next_btn.setVisible(False)
            elif idx == 3:
                if self._postinstall_state == "completed":
                    self.next_btn.setText("Review Results >")
                    self.next_btn.setVisible(True)
                else:
                    self.next_btn.setVisible(False)
            elif idx == total_pages - 1:
                self.next_btn.setVisible(False)

        for i, lbl in enumerate(self.step_labels):
            if i == idx:
                lbl.setStyleSheet(
                    "QPushButton { padding: 8px 12px; border-radius: 4px; color: white;"
                    " background: #2a82da; font-size: 13px; font-weight: bold;"
                    " text-align: left; border: none; }"
                    " QPushButton:hover { background: #3a92ea; }"
                )
            elif i < idx:
                lbl.setStyleSheet(
                    "QPushButton { padding: 8px 12px; border-radius: 4px; color: #4ec94e;"
                    " font-size: 13px; text-align: left; border: none; }"
                    " QPushButton:hover { background: #3a3a3a; }"
                )
            else:
                lbl.setStyleSheet(
                    "QPushButton { padding: 8px 12px; border-radius: 4px; color: #666;"
                    " font-size: 13px; text-align: left; border: none; }"
                    " QPushButton:hover { background: #3a3a3a; }"
                )

    def _on_version_changed(self, ver: str):
        self.nx_version = ver
        self.version_info.setText(f"NX {self.nx_version}")
        self.setWindowTitle(f"Siemens NX {self.nx_version} Installer")

    def _navigate_to(self, idx: int):
        if idx < 0 or idx >= len(self._pages) or idx == self._current_page:
            return
        if self._current_page == 2 and self._install_state == "running":
            reply = QMessageBox.question(
                self, "Cancel Installation",
                "An installation is in progress. Cancel and navigate away?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            self.install_page._on_cancel()
            self._install_state = "pending"
        elif self._current_page == 3 and self._postinstall_state == "running":
            reply = QMessageBox.question(
                self, "Cancel Post-Install",
                "Post-install is in progress. Cancel and navigate away?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            self.postinstall_page._on_cancel()
            self._postinstall_state = "pending"

        if idx == 2:
            self.install_page.set_install_params(
                self.config, self.selected_features, self.nx_version
            )
            if self._install_state in ("pending", "failed"):
                self.install_page.reset_ui()
        elif idx == 3:
            self.postinstall_page.set_postinstall_params(
                self.config, self.nx_version
            )
            if self._postinstall_state in ("pending", "failed"):
                self.postinstall_page.reset_ui()

        self._show_page(idx)

    def _on_next(self):
        if self._current_page == 0:
            if not self.config_page.validate():
                return
            self.config_page.save_config()
            nx_ver = self.config_page.get_version()
            if nx_ver:
                self.nx_version = nx_ver
                self.version_info.setText(f"NX {self.nx_version}")
                self.setWindowTitle(f"Siemens NX {self.nx_version} Installer")
            msi = find_msi(self.config.install_files)
            if msi:
                self.feature_page.load_features(msi)
            self._navigate_to(1)

        elif self._current_page == 1:
            self.selected_features = self.feature_page.get_selected()
            if not self.selected_features:
                QMessageBox.warning(self, "No Features", "Select at least one feature, or use Defaults.")
                return
            self._navigate_to(2)

        elif self._current_page == 2 and self._install_state == "completed":
            self._navigate_to(3)

        elif self._current_page == 3 and self._postinstall_state == "completed":
            self._show_complete(True)

    def _on_back(self):
        if self._current_page > 0:
            self._navigate_to(self._current_page - 1)

    def _on_uninstall_done(self, success: bool):
        self._show_complete(success)

    def _on_install_done(self, success: bool):
        self._install_state = "completed" if success else "failed"
        if self._current_page == 2:
            self._show_page(2)

    def _on_postinstall_done(self, success: bool):
        self._postinstall_state = "completed" if success else "failed"
        if self._current_page == 3:
            self._show_page(3)

    def _show_complete(self, success: bool):
        operation = "Uninstall" if self._mode == "uninstall" else "Installation"
        page_idx = 1 if self._mode == "uninstall" else 4
        self._show_page(page_idx)
        results = {}
        if success and self._mode == "install":
            logger = logging.getLogger("nx_install")
            logger.setLevel(logging.DEBUG)
            validator = PostInstallValidator(self.config, logger, self.nx_version)
            validator.validate()
            results = validator.results
        self.complete_page.show_results(success, self.config, results, operation)


def main():
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent.resolve()
        embedded = Path(sys._MEIPASS) / "config.ini"
        override = exe_dir / "config.ini"
        config_path = override if override.exists() else embedded
    else:
        config_path = Path(__file__).parent.resolve() / "config.ini"
    log_dir = os.environ.get("TEMP", "C:\\Temp")
    log_level = "INFO"
    if config_path.exists():
        try:
            cfg = Config(str(config_path))
            log_level = cfg.log_level
            log_dir = cfg.temp_dir
        except Exception:
            pass
    setup_logging(log_level, log_dir)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setPalette(dark_palette())
    app.setStyleSheet(DARK_STYLESHEET)

    font = QFont("Segoe UI", 10)
    app.setFont(font)

    wizard = NXInstallWizard()
    wizard.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
