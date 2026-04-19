#!/usr/bin/env python3
"""
Headless Siemens NX2506 Installation Script for Windows.

Automates unattended installation of Siemens NX 2506 via msiexec.
Media: D:\\Install files\\SiemensNX-2506.7002_wntx64\\SiemensNX-2506.7002_wntx64

MSI properties (confirmed from actual MSI log):
  INSTALLDIR       - installation directory
  SPLMLICENSESERVER - license server (custom action reads SPLM_LICENSE_SERVER env var)

Usage:
    python install_nx.py [--config CONFIG_PATH] [--unattend]
    python install_nx.py --uninstall
    python install_nx.py --validate-only

Interactive mode: select features via menu (SPACE to toggle, ENTER to confirm).
--unattend mode: installs default feature set automatically.
"""

import argparse
import ctypes
import configparser
import hashlib
import logging
import os
import re
import shutil
import subprocess
import sys
import textwrap
import time
import urllib.request
import winreg
import zipfile
from pathlib import Path
from typing import Optional

import colorama
from prompt_toolkit import Application
from prompt_toolkit.key_binding import KeyBindings
from prompt_toolkit.layout import Layout
from prompt_toolkit.layout.containers import HSplit, Window
from prompt_toolkit.layout.controls import FormattedTextControl
from prompt_toolkit.styles import Style
from prompt_toolkit.widgets import Frame

colorama.init(convert=True, strip=False)

SCRIPT_VERSION = "1.5.0"
LOG_FILE: Optional[logging.Handler] = None

MSI_PRODUCT_CODE = "{CB259176-5CE9-40A9-8B3D-8EBF288E80AB}"

FEATURE_MAP = {
    "FEAT_NXPLATFORM":              "NX CAD (Core)",
    "FEAT_CORE":                    "NX Core",
    "FEAT_VSIX":                    "NX Visual Studio Extensions",
    "FEAT_MANUFACTURING":           "NX Manufacturing (CAM)",
    "FEAT_MANUFACTURING_PLANNING":  "NX Manufacturing Planning",
    "FEAT_MESHINGSRV":              "NX Meshing Services",
    "FEAT_MECHATRONICS":            "NX Mechatronics Concept Designer",
    "FEAT_SIMULATION":              "NX Simcenter 3D (CAE)",
    "FEAT_NXNASTRAN":               "Simcenter Nastran Solver",
    "FEAT_NXTMG":                   "Simcenter 3D Thermal Flow",
    "FEAT_NXACOUSTICS":             "Simcenter 3D Acoustics Solver",
    "FEAT_SAMCEF":                  "Samcef Solver",
    "FEAT_CFDD":                    "CFD Designer",
    "FEAT_ROUTING":                 "NX Routing (Cabinet/Harness/Piping)",
    "FEAT_DRAFTING":                "NX Drafting & 2D",
    "FEAT_TRANSLATORS":             "NX Translators (all formats)",
    "FEAT_UGFLEXLM":                "NX Flexible Modeling",
    "FEAT_COMPOSITES":              "NX Composites",
    "FEAT_MOLDED_PART_DESIGN":      "NX Molded Part Design",
    "FEAT_TOOLING_DESIGN":          "NX Tooling Design & NX Join",
    "FEAT_IMMERSIVE":               "NX Immersive",
    "FEAT_DIAGRAMMING":             "NX Diagramming",
    "FEAT_AUTOMOTIVE":              "NX Automotive",
    "FEAT_SHIP_BUILDING":           "NX Ship Building",
    "FEAT_INDUSTRIAL_ELECTRICAL_DESIGN": "NX Industrial Electrical Design",
    "FEAT_FABRICMODELER":           "NX Fabric Modeler",
    "FEAT_ADDMANSIM":               "NX Additive Manufacturing Simulation",
    "FEAT_AUTOMATED_TESTING_STUDIO": "NX Automated Testing Studio",
    "FEAT_MODELBASEDPARTMANUFACTURING": "NX Model-Based Part Manufacturing",
    "FEAT_OPTIMIZATION_TOOLS":      "NX Optimization Tools",
    "FEAT_PROGRAMMING_TOOLS":       "NX Programming Tools",
    "FEAT_NXREPORTING":             "NX Reporting",
    "FEAT_STUDIO_RENDER":           "Photo Realistic Rendering",
    "FEAT_ECLASS_NX_AUTHOR":        "NX Author for ECLASS",
    "FEAT_GREATERCHINATOOLS":       "Greater China Tools",
    "FEAT_VALIDATION":              "NX Validation",
    "FEAT_OPTIONAL":                "Optional Components",
    "FEAT_LOCALIZATION":            "NX Localization (base)",
}

DEFAULT_FEATURES = {
    "FEAT_NXPLATFORM", "FEAT_PROGRAMMING_TOOLS", "FEAT_COMPOSITES",
    "FEAT_MECHATRONICS", "FEAT_SIMULATION", "FEAT_NXNASTRAN",
    "FEAT_TRANSLATORS", "FEAT_STUDIO_RENDER",
}


class ColoredFormatter(logging.Formatter):
    COLORS = {
        "DEBUG": colorama.Fore.CYAN,
        "INFO": colorama.Fore.GREEN,
        "WARNING": colorama.Fore.YELLOW,
        "ERROR": colorama.Fore.RED,
        "CRITICAL": colorama.Fore.MAGENTA,
    }
    RESET = colorama.Style.RESET_ALL

    def format(self, record):
        record.levelname = (
            f"{self.COLORS.get(record.levelname, self.RESET)}"
            f"{record.levelname}{self.RESET}"
        )
        return super().format(record)


def setup_logging(level: str = "INFO", log_dir: Optional[str] = None) -> logging.Logger:
    global LOG_FILE
    logger = logging.getLogger("nx_install")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(ColoredFormatter("%(levelname)s  %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(console)
    if log_dir:
        Path(log_dir).mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        log_path = Path(log_dir) / f"nx_install_{timestamp}.log"
        LOG_FILE = logging.FileHandler(log_path, encoding="utf-8")
        LOG_FILE.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(LOG_FILE)
    return logger


def check_admin_rights() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def check_license_server(license_server: str, timeout: int = 10, logger: Optional[logging.Logger] = None) -> bool:
    if not license_server:
        if logger:
            logger.error("License server is not configured.")
        return False

    license_server = license_server.strip()
    if "@" not in license_server:
        if logger:
            logger.error(f"Invalid license server format: '{license_server}' (expected PORT@HOST)")
        return False

    port_str, host = license_server.rsplit("@", 1)
    host = host.strip()
    port_str = port_str.strip()

    try:
        port = int(port_str)
    except ValueError:
        if logger:
            logger.error(f"Invalid port '{port_str}' in license server: '{license_server}'")
        return False

    if logger:
        logger.info(f"Checking license server connectivity: {host}:{port}")

    import socket
    try:
        sock = socket.create_connection((host, port), timeout=timeout)
        sock.close()
        if logger:
            logger.info(f"License server {host}:{port} is reachable.")
        return True
    except socket.timeout:
        if logger:
            logger.error(f"License server {host}:{port} timed out after {timeout}s.")
        return False
    except socket.gaierror as e:
        if logger:
            logger.error(f"Cannot resolve hostname '{host}': {e}")
        return False
    except OSError as e:
        if logger:
            logger.error(f"Cannot connect to {host}:{port} — {e}")
        return False



def get_free_disk_space(path: str) -> int:
    free_bytes = ctypes.c_ulonglong(0)
    ctypes.windll.kernel32.GetDiskFreeSpaceExW(ctypes.c_wchar_p(path), None, None, ctypes.pointer(free_bytes))
    return free_bytes.value


def wait_for_process(proc: subprocess.Popen, logger: logging.Logger, timeout: int, poll_interval: int = 30) -> int:
    start = time.time()
    while True:
        retcode = proc.poll()
        if retcode is not None:
            return retcode
        elapsed = time.time() - start
        if elapsed >= timeout:
            logger.error(f"Installation timed out after {timeout}s")
            proc.kill()
            return -1
        if int(elapsed) % 60 == 0 and int(elapsed) > 0:
            logger.info(f"  Still running... ({int(elapsed // 60)}m elapsed)")
        time.sleep(poll_interval)


class Config:
    def __init__(self, config_path: str):
        cp = configparser.ConfigParser()
        cp.read(config_path, encoding="utf-8")
        self.install_files = cp.get("paths", "install_files").strip()
        self.install_dir = cp.get("paths", "install_dir").strip()
        self.temp_dir = cp.get("paths", "temp_dir", fallback=os.environ.get("TEMP", "C:\\Temp")).strip()
        self.install_vcpp = cp.getboolean("prerequisites", "install_vcpp", fallback=True)
        self.install_dotnet = cp.getboolean("prerequisites", "install_dotnet", fallback=True)
        self.license_server = cp.get("license", "splm_license_server").strip()
        self.name = cp.get("user", "name", fallback="").strip()
        self.surname = cp.get("user", "surname", fallback="").strip()
        self.fcc_url = cp.get("downloads", "fcc_url").strip()
        self.java_url = cp.get("downloads", "java_url").strip()
        self.start_nx_url = cp.get("downloads", "start_nx_url").strip()
        self.role_url = cp.get("downloads", "role_url", fallback="").strip()
        self.dpv_url = cp.get("downloads", "dpv_url", fallback="").strip()
        self.fcg_url = cp.get("downloads", "fcg_url", fallback="").strip()
        self.log_level = cp.get("logging", "log_level", fallback="INFO").strip()
        self.install_timeout = cp.getint("timeouts", "install_timeout_seconds", fallback=5400)


class PrerequisitesInstaller:
    VCRUNTIME_URL = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    DOTNET48_URL = "https://go.microsoft.com/fwlink/?linkid=2088631"

    def __init__(self, install_files: str, logger: logging.Logger,
                 install_vcpp: bool = True, install_dotnet: bool = True,
                 temp_dir: Optional[str] = None):
        self.install_files = Path(install_files)
        self.logger = logger
        self.install_vcpp = install_vcpp
        self.install_dotnet = install_dotnet
        self.temp_dir = temp_dir or os.environ.get("TEMP", "C:\\Temp")
        self._found_vcredist = None
        self._found_dotnet48 = None
        self._found_webview2 = None
        self._found_aspnetcore = None
        self._found_desktop_runtime = None
        self._find_local()

    def _find_local(self):
        self.logger.debug("Scanning media for prerequisite installers...")
        for p in self.install_files.rglob("*"):
            if not p.is_file():
                continue
            ln = p.name.lower()
            if "vc_redist" in ln and ("x64" in ln or "amd64" in ln) and ln.endswith(".exe"):
                self._found_vcredist = str(p)
            elif "ndp48" in ln and ln.endswith(".exe"):
                self._found_dotnet48 = str(p)
            elif "webview2" in ln and ln.endswith(".exe"):
                self._found_webview2 = str(p)
            elif "aspnetcore" in ln and ln.endswith(".exe"):
                self._found_aspnetcore = str(p)
            elif "windowsdesktop-runtime" in ln and ln.endswith(".exe"):
                self._found_desktop_runtime = str(p)

    def _run(self, path: str, args: list, timeout: int, label: str) -> bool:
        if not Path(path).exists():
            self.logger.warning(f"{label} installer not found: {path}")
            return False
        self.logger.info(f"Installing {label}...")
        try:
            proc = subprocess.Popen([path] + args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            retcode = wait_for_process(proc, self.logger, timeout, poll_interval=10)
            ok = retcode in (0, 1638, 3010)
            self.logger.info(f"{label}: {'OK' if ok else f'warning (exit {retcode})'}")
            return ok
        except Exception as e:
            self.logger.warning(f"{label} error: {e}")
            return False

    def _check_vcpp(self) -> bool:
        for root in [r"HKLM\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
                     r"HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"]:
            try:
                r = subprocess.run(["reg", "query", root, "/v", "Installed"],
                                   capture_output=True, text=True)
                if "Installed" in r.stdout and "0x1" in r.stdout:
                    return True
            except Exception:
                pass
        for root in [r"HKLM\SOFTWARE\Microsoft\VisualStudio", r"HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio"]:
            try:
                r = subprocess.run(["reg", "query", root], capture_output=True, text=True)
                if "VC_RUNTIME" in r.stdout and "14." in r.stdout:
                    return True
            except Exception:
                pass
        for uk in [r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                   r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"]:
            try:
                r = subprocess.run(["reg", "query", uk], capture_output=True, text=True)
                if re.search(r"Microsoft Visual C\+\+.*2015.*2022", r.stdout):
                    return True
            except Exception:
                pass
        return False

    def _check_dotnet48(self) -> bool:
        try:
            r = subprocess.run(
                ["reg", "query", r"HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full", "/v", "Version"],
                capture_output=True, text=True)
            m = re.search(r"Version\s+REG_SZ\s+(\S+)", r.stdout)
            if not m:
                return False
            parts = [int(x) for x in m.group(1).split(".")]
            return parts[0] >= 4 and parts[1] >= 8
        except Exception:
            return False

    def _do_vcpp(self) -> bool:
        if self._check_vcpp():
            self.logger.info("VC++ 2015-2022 Redistributable already present, skipping.")
            return True
        path = self._found_vcredist
        if not path:
            path = str(Path(self.temp_dir) / "vc_redist.x64.exe")
            if not Path(path).exists():
                self.logger.info("Downloading VC++ Redistributable...")
                try:
                    urllib.request.urlretrieve(self.VCRUNTIME_URL, path)
                except Exception as e:
                    self.logger.error(f"Failed to download VC++ redist: {e}")
                    return False
        return self._run(path, ["/install", "/quiet", "/norestart"], 300, "VC++ Redistributable")

    def _do_dotnet48(self) -> bool:
        if self._check_dotnet48():
            self.logger.info(".NET Framework 4.8 already present, skipping.")
            return True
        path = self._found_dotnet48
        if not path:
            path = str(Path(self.temp_dir) / "ndp48-x86-x64-allos-enu.exe")
            if not Path(path).exists():
                self.logger.info("Downloading .NET Framework 4.8...")
                try:
                    urllib.request.urlretrieve(self.DOTNET48_URL, path)
                except Exception as e:
                    self.logger.error(f"Failed to download .NET 4.8: {e}")
                    return False
        return self._run(path, ["/quiet", "/norestart"], 600, ".NET 4.8")

    def install_webview2(self) -> bool:
        if self._check_webview2():
            self.logger.info("WebView2 Runtime already installed, skipping.")
            return True
        if self._found_webview2:
            return self._run(self._found_webview2, ["--do-not-launch-edge"], 300, "WebView2 Runtime")
        return True

    def _check_webview2(self) -> bool:
        for root in [r"HKLM\SOFTWARE\Microsoft\EdgeUpdate\Clients",
                     r"HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients",
                     r"HKCU\SOFTWARE\Microsoft\EdgeUpdate\Clients"]:
            try:
                r = subprocess.run(["reg", "query", root], capture_output=True, text=True)
                if "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" in r.stdout:
                    return True
            except Exception:
                pass
        return False

    def install_aspnetcore(self) -> bool:
        if self._found_aspnetcore:
            return self._run(self._found_aspnetcore, ["/quiet", "/norestart"], 300, "ASP.NET Core Runtime")
        return True

    def install_desktop_runtime(self) -> bool:
        if self._found_desktop_runtime:
            return self._run(self._found_desktop_runtime, ["/quiet", "/norestart"], 300, "Windows Desktop Runtime")
        return True

    def install_all(self) -> bool:
        ok = True
        for impl in [
            self._do_vcpp, self._do_dotnet48,
            self.install_webview2, self.install_aspnetcore, self.install_desktop_runtime,
        ]:
            if not impl():
                ok = False
        return ok


def _check_registry_install(product_code: str) -> tuple[bool, Optional[str]]:
    UNINSTALL_KEY = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        for view in [0, winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]:
            try:
                subkey = f"{UNINSTALL_KEY}\\{product_code}"
                with winreg.OpenKey(hive, subkey, 0, view | winreg.KEY_READ) as key:
                    try:
                        loc, _ = winreg.QueryValueEx(key, "InstallLocation")
                    except FileNotFoundError:
                        loc = None
                    return True, loc
            except FileNotFoundError:
                pass
            except Exception:
                pass
    return False, None


def get_msi_installed_location(product_code: str) -> Optional[str]:
    _, loc = _check_registry_install(product_code)
    return loc


def is_nx_installed(product_code: str = MSI_PRODUCT_CODE) -> bool:
    installed, _ = _check_registry_install(product_code)
    return installed


def uninstall_nx(product_code: str, logger: logging.Logger, timeout: int = 600) -> bool:
    logger.info("Uninstalling existing NX installation...")
    log_path = Path(os.environ.get("TEMP", "C:\\Temp")) / "nx_uninstall.log"
    cmd = ["msiexec.exe", "/x", product_code, "/qn", "/norestart", "/l*v", str(log_path)]
    logger.info(f"Command: {' '.join(cmd[:3])} ...")
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                creationflags=subprocess.CREATE_NO_WINDOW)
        retcode = wait_for_process(proc, logger, timeout)
        if retcode in (0, 3010):
            logger.info("Uninstall completed successfully.")
            return True
        logger.error(f"Uninstall failed with exit code: {retcode}")
        if log_path.exists():
            with open(log_path, encoding="utf-8", errors="ignore") as f:
                for line in f.read().splitlines()[-20:]:
                    if line.strip():
                        logger.debug(f"  {line}")
        return False
    except Exception as e:
        logger.error(f"Uninstall error: {e}")
        return False


def fix_nx_permissions(install_dir: str, logger: logging.Logger, timeout: int = 120) -> bool:
    logger.info(f"Fixing ownership and permissions on {install_dir}...")
    install_path = install_dir.rstrip("\\")
    try:
        subprocess.run(["takeown", "/F", install_path, "/R", "/A", "/D", "Y"],
                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=timeout)
        subprocess.run(["icacls", install_path, "/grant:r", "%USERNAME%:(OI)(CI)F", "/T"],
                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=timeout)

        fcc_path = str(Path(install_path) / "UGMANAGER" / "tccs" / "fcc.xml")
        subprocess.run(["attrib", "-R", fcc_path],
                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=30)

        logger.info("Ownership set to Administrators, permissions granted, attributes cleaned.")
        return True
    except subprocess.TimeoutExpired:
        logger.error(f"Permission fix timed out after {timeout}s.")
        return False
    except Exception as e:
        logger.warning(f"Permission fix failed: {e}")
        return False


def configure_role(config: Config, logger: logging.Logger) -> bool:
    role_url = config.role_url.strip()
    dpv_url = config.dpv_url.strip()
    fcg_url = config.fcg_url.strip()
    if not role_url and not dpv_url and not fcg_url:
        logger.debug("No role/dpv/fcg URL configured, skipping.")
        return True
    logger.info("Configuring NX role and preferences...")
    target_dir = Path(os.environ["LOCALAPPDATA"]) / "Siemens" / "NX2506"
    target_dir.mkdir(parents=True, exist_ok=True)
    dl = FileDownloader(config, logger)

    if role_url:
        role_file = dl.download(role_url, "user.mtx")
        if role_file:
            shutil.copy2(role_file, target_dir / "user.mtx")
            prefs = target_dir / "UserPreferences.txt"
            mtx_path = str(target_dir / "user.mtx").replace("\\", "\\\\")
            prefs_entry = '[HKEY_CURRENT_USER\\Software\\Unigraphics Solutions\\NX\\2506\\Layout]\n"LastRole"="{0}"\n'.format(mtx_path)
            with open(prefs, "a", encoding="utf-8") as f:
                f.write(prefs_entry)
            logger.info(f"Role configured: {target_dir / 'user.mtx'}")
        else:
            logger.warning("Role download failed, skipping.")

    if dpv_url:
        dpv_file = dl.download(dpv_url, "NX_user.dpv")
        if dpv_file:
            shutil.copy2(dpv_file, target_dir / "NX_user.dpv")
            logger.info(f"Preferences configured: {target_dir / 'NX_user.dpv'}")
        else:
            logger.warning("dpv download failed, skipping.")

    if fcg_url:
        fcg_file = dl.download(fcg_url, "feature_toggle_user.fcg")
        if fcg_file:
            shutil.copy2(fcg_file, target_dir / "feature_toggle_user.fcg")
            logger.info(f"Feature toggle configured: {target_dir / 'feature_toggle_user.fcg'}")
        else:
            logger.warning("fcg download failed, skipping.")

    return True


def get_msi_features(msi_path: str, temp_dir: str) -> set:
    vbs = textwrap.dedent(f"""\
        Set msi = CreateObject("WindowsInstaller.Installer")
        Set db = msi.OpenDatabase("{msi_path}", 0)
        Set view = db.OpenView("SELECT Feature FROM Feature")
        view.Execute
        Do
            Set rec = view.Fetch
            If rec Is Nothing Then Exit Do
            WScript.Echo rec.StringData(1)
        Loop
        view.Close
    """)
    tmp = Path(temp_dir) / "get_msi_features.vbs"
    try:
        tmp.write_text(vbs, encoding="utf-16le")
        result = subprocess.run(
            ["cscript", "//Nologo", "//E:vbscript", str(tmp)],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        return set(line.strip() for line in result.stdout.splitlines() if line.strip())
    except Exception:
        return set()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


class NXInstaller:
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self._find_msi()

    def _get_msi_features(self) -> set:
        return get_msi_features(self.msi_path, self.config.temp_dir)

    def _find_msi(self):
        for p in [
            Path(self.config.install_files) / "nx" / "SiemensNX.msi",
            Path(self.config.install_files) / "SiemensNX.msi",
        ]:
            if p.exists():
                self.msi_path = str(p)
                self.logger.info(f"Found MSI: {self.msi_path}")
                return
        for p in Path(self.config.install_files).rglob("SiemensNX.msi"):
            self.msi_path = str(p)
            self.logger.info(f"Found MSI: {self.msi_path}")
            return
        raise FileNotFoundError(f"SiemensNX.msi not found in {self.config.install_files}")

    def _check_disk(self) -> bool:
        drive = Path(self.config.install_dir).drive or Path(self.config.install_dir).anchor
        req_bytes = 28 * (1024**3)
        free = get_free_disk_space(drive)
        if free < req_bytes:
            self.logger.error(f"Insufficient disk space on {drive}: "
                              f"{free / (1024**3):.1f} GB free, 30 GB required")
            return False
        self.logger.info(f"Disk space: {free / (1024**3):.1f} GB available on {drive}")
        return True

    def _read_msi_log(self, log_path: str):
        try:
            with open(log_path, encoding="utf-8", errors="ignore") as f:
                lines = f.read().splitlines()
            errors = [l for l in lines if "Error" in l or "error" in l or "1639" in l or "1603" in l]
            for l in errors[-15:]:
                if l.strip():
                    self.logger.debug(f"  msilog: {l.strip()}")
        except Exception:
            pass

    def install(self, features: list) -> bool:
        self.logger.info("=" * 60)
        self.logger.info("Starting NX2506 Installation")
        self.logger.info("=" * 60)

        if is_nx_installed():
            self.logger.warning("NX 2506 is already installed. Use --uninstall first to reinstall.")
            self.logger.warning(f"Existing install location: {get_msi_installed_location(MSI_PRODUCT_CODE)}")
            return False

        if not self._check_disk():
            return False

        if self.config.install_vcpp or self.config.install_dotnet:
            self.logger.info("Installing prerequisites...")
            prereq = PrerequisitesInstaller(self.config.install_files, self.logger,
                                           self.config.install_vcpp, self.config.install_dotnet,
                                           self.config.temp_dir)
            prereq.install_all()
            self.logger.info("Prerequisites processed.")

        self.logger.info("-" * 40)
        resolved = list(features)

        msi_features = self._get_msi_features()
        if msi_features:
            valid = [f for f in resolved if f in msi_features]
            invalid = [f for f in resolved if f not in msi_features]
            for f in invalid:
                self.logger.warning(f"Skipping {f} — not found in MSI (FEATURE_MAP says: {FEATURE_MAP.get(f, f)})")
            if valid:
                resolved = valid
            else:
                self.logger.warning("None of the selected features found in MSI. Falling back to ALL.")
        else:
            self.logger.debug("Could not read MSI feature table. Using feature list as-is.")

        if resolved:
            for f in resolved:
                self.logger.info(f"  Feature: {f} ({FEATURE_MAP.get(f, f)})")

        addlocal = ",".join(resolved) if resolved else "ALL"
        self.logger.info(f"ADDLOCAL: {addlocal}")

        log_path = Path(self.config.temp_dir) / "nx_msi_install.log"
        Path(self.config.temp_dir).mkdir(parents=True, exist_ok=True)
        Path(self.config.install_dir).mkdir(parents=True, exist_ok=True)

        env = os.environ.copy()
        env["SPLM_LICENSE_SERVER"] = self.config.license_server

        install_dir = self.config.install_dir.rstrip("\\")
        cmd = [
            "msiexec.exe",
            "/i", self.msi_path,
            "/qn",
            "/norestart",
            "/l*v", str(log_path),
            f"INSTALLDIR={install_dir}",
            f"SPLMLICENSESERVER={self.config.license_server}",
            f"ADDLOCAL={addlocal}",
        ]
        self.logger.info(f"Command: msiexec /i <msi> /qn /norestart [properties...]")

        try:
            proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                    env=env, creationflags=subprocess.CREATE_NO_WINDOW)
            retcode = wait_for_process(proc, self.logger, self.config.install_timeout)

            self._read_msi_log(str(log_path))

            if retcode == 0:
                self.logger.info("Installation completed successfully.")
                return True
            if retcode == 3010:
                self.logger.info("Installation completed — reboot required (3010).")
                return True
            self.logger.error(f"Installation failed with exit code: {retcode}")
            return False
        except Exception as e:
            self.logger.error(f"Installation failed: {e}")
            return False


class LicenseConfigurator:
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger

    def configure(self) -> bool:
        self.logger.info("Setting SPLM_LICENSE_SERVER environment variable...")
        env_var = "SPLM_LICENSE_SERVER"
        value = self.config.license_server
        os.environ[env_var] = value

        for cmd_args in [["setx", env_var, value, "/M"], ["setx", env_var, value]]:
            try:
                r = subprocess.run(cmd_args, capture_output=True, text=True,
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                if r.returncode == 0:
                    scope = "system" if "/M" in cmd_args else "user"
                    self.logger.info(f"Set {env_var}={value} ({scope}-level)")
                    return True
            except Exception as e:
                self.logger.warning(f"setx failed: {e}")
        self.logger.error(f"Failed to set {env_var}")
        return False


class PostInstallValidator:
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.results: dict = {}

    def validate(self) -> bool:
        self.logger.info("-" * 40)
        self.logger.info("Running post-install validation...")

        self._check_binary(self.config.install_dir + "\\UGII\\ugraf.exe", "ugraf.exe")
        self._check_env_var()
        self._check_fcc()
        self._check_nx_prefs()

        self.logger.info("-" * 40)
        self.logger.info("Validation Summary:")
        all_passed = True
        for check, passed in self.results.items():
            status = "PASS" if passed else "FAIL"
            self.logger.info(f"  [{status}] {check}")
            if not passed:
                all_passed = False

        if all_passed:
            self.logger.info("All validation checks passed.")
        else:
            self.logger.error("One or more validation checks failed.")
        return all_passed

    def _check_binary(self, path: str, label: str):
        exists = Path(path).exists()
        self.results[label] = exists
        if not exists:
            self.logger.warning(f"Not found: {path}")

    def _check_env_var(self):
        val = os.environ.get("SPLM_LICENSE_SERVER", "")
        self.results["SPLM_LICENSE_SERVER Env"] = bool(val)

    def _check_fcc(self):
        ref_fcc = Path(self.config.temp_dir) / "fcc.xml"
        inst_fcc = Path(self.config.install_dir) / "UGMANAGER" / "tccs" / "fcc.xml"
        self._check_file("fcc.xml checksum", ref_fcc, inst_fcc)

        self._check_java_zip(Path(self.config.temp_dir) / "java.zip",
                             Path(self.config.install_dir).parent / "java")

    def _check_java_zip(self, ref_zip: Path, inst_dir: Path):
        try:
            zip_ok = False
            if ref_zip.exists():
                try:
                    with zipfile.ZipFile(ref_zip, "r") as zf:
                        zf.testzip()
                    self.results["java.zip integrity"] = True
                    self.logger.info(f"  java.zip integrity: OK")
                except zipfile.BadZipFile:
                    self.results["java.zip integrity"] = False
                    self.logger.warning(f"  java.zip integrity: BAD ZIP FILE")
                except Exception as e:
                    self.results["java.zip integrity"] = False
                    self.logger.warning(f"  java.zip integrity: {e}")
            else:
                self.results["java.zip integrity"] = False
                self.logger.warning(f"  java.zip integrity: not found in temp")

            zulu = inst_dir / "zulu11"
            zulu_ok = zulu.exists()
            self.results["java dir"] = zulu_ok
            if zulu_ok:
                self.logger.info(f"  java dir: OK ({zulu})")
            else:
                self.logger.warning(f"  java dir: not found ({zulu})")
        except Exception as e:
            self.logger.warning(f"  java validation error: {e}")

    def _check_nx_prefs(self):
        prefs_dir = Path(os.environ["LOCALAPPDATA"]) / "Siemens" / "NX2506"
        for filename in ["user.mtx", "NX_user.dpv", "feature_toggle_user.fcg"]:
            path = prefs_dir / filename
            exists = path.exists()
            self.results[f"{filename}"] = exists
            if exists:
                self.logger.info(f"  {filename}: OK")
            else:
                self.logger.warning(f"  {filename}: not found")

    def _check_zip(self, target_name: str, ref_zip: Path, ref_dir: Path, inst_dir: Path):
        zip_label = f"{target_name}.zip checksum"
        dir_label = f"{target_name} dir checksum"

        try:
            ref_zip_exists = ref_zip.exists()
            inst_dir_exists = inst_dir.exists()

            if not inst_dir_exists:
                self.results[zip_label] = False
                self.results[dir_label] = False
                self.logger.warning(f"  {dir_label}: installed dir not found")
                return

            if ref_zip_exists:
                ref_zip_hash = hashlib.sha256(ref_zip.read_bytes()).hexdigest()
                self.results[zip_label] = True
                self.logger.info(f"  {zip_label}: OK ({ref_zip_hash[:16]}...)")
            else:
                self.results[zip_label] = False
                self.logger.warning(f"  {zip_label}: reference zip not in temp")

            if ref_dir.exists():
                ref_hash = FileDownloader.dir_checksum(ref_dir)
                inst_hash = FileDownloader.dir_checksum(inst_dir)
                dir_match = ref_hash == inst_hash
                self.results[dir_label] = dir_match
                if dir_match:
                    self.logger.info(f"  {dir_label}: OK ({inst_hash[:16]}...)")
                else:
                    self.logger.warning(f"  {dir_label}: MISMATCH")
                    self.logger.warning(f"    Reference: {ref_hash[:16]}...")
                    self.logger.warning(f"    Installed: {inst_hash[:16]}...")
            else:
                self.results[dir_label] = None
                self.logger.debug(f"  {dir_label}: reference dir not in temp")
        except Exception as e:
            self.results[zip_label] = False
            self.results[dir_label] = False
            self.logger.warning(f"  {target_name}: check error — {e}")

    def _check_file(self, label: str, ref_path: Path, inst_path: Path):
        try:
            ref_exists = ref_path.exists()
            inst_exists = inst_path.exists()

            if not inst_exists:
                self.results[label] = False
                self.logger.warning(f"  {label}: installed file not found")
                return

            if not ref_exists:
                inst_hash = hashlib.sha256(inst_path.read_bytes()).hexdigest()
                self.results[label] = True
                self.logger.info(f"  {label}: OK ({inst_hash[:16]}...)")
                return

            ref_hash = hashlib.sha256(ref_path.read_bytes()).hexdigest()
            inst_hash = hashlib.sha256(inst_path.read_bytes()).hexdigest()
            match = ref_hash == inst_hash
            self.results[label] = match
            if match:
                self.logger.info(f"  {label}: OK ({inst_hash[:16]}...)")
            else:
                self.logger.warning(f"  {label}: MISMATCH")
                self.logger.warning(f"    Reference: {ref_hash[:16]}...")
                self.logger.warning(f"    Installed: {inst_hash[:16]}...")
        except Exception as e:
            self.results[label] = False
            self.logger.warning(f"  {label}: check error — {e}")


class FileDownloader:
    def __init__(self, config: Config, logger: logging.Logger):
        self.config = config
        self.logger = logger

    def download(self, url: str, filename: str) -> Optional[Path]:
        Path(self.config.temp_dir).mkdir(parents=True, exist_ok=True)
        dest = Path(self.config.temp_dir) / filename
        self.logger.info(f"Downloading {filename}...")
        try:
            r = subprocess.run(
                ["curl", "-fsSL", "-o", str(dest), url],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW, timeout=600,
            )
            if r.returncode != 0:
                self.logger.error(f"Download failed: {r.stderr.strip()}")
                return None
            return dest
        except subprocess.TimeoutExpired:
            self.logger.error(f"Download timed out: {url}")
            return None
        except Exception as e:
            self.logger.error(f"Download error: {e}")
            return None

    def unzip(self, zip_path: Path, dest_dir: Optional[Path] = None) -> bool:
        dest_dir = dest_dir or Path(self.config.install_dir)
        self.logger.info(f"Unzipping {zip_path.name} to {dest_dir}...")
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(dest_dir)
            self.logger.info(f"Extracted: {zip_path.stem}/")
            return True
        except zipfile.BadZipFile as e:
            self.logger.error(f"Invalid zip file: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Unzip error: {e}")
            return False

    def transform(self, src: Path, dest: Path, transform_func) -> bool:
        self.logger.info(f"Transforming {src.name} to {dest.name}...")
        try:
            content = src.read_text(encoding="utf-8")
            transformed = transform_func(content)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(transformed, encoding="utf-8")
            self.logger.info(f"Created: {dest}")
            return True
        except Exception as e:
            self.logger.error(f"Transform error: {e}")
            return False

    def move(self, src: Path, dest: Path) -> bool:
        self.logger.info(f"Moving {src.name} to {dest}...")
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                try:
                    if dest.is_dir():
                        shutil.rmtree(dest)
                    else:
                        dest.unlink()
                except Exception:
                    try:
                        subprocess.run(["takeown", "/F", str(dest), "/R", "/A", "/D", "Y"],
                                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=120)
                        subprocess.run(["icacls", str(dest), "/grant:r", "%USERNAME%:(OI)(CI)F", "/T"],
                                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=120)
                        subprocess.run(["attrib", "-R", "-S", "-H", str(dest)],
                                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=30)
                        if dest.is_dir():
                            shutil.rmtree(dest)
                        else:
                            dest.unlink()
                    except Exception as e:
                        self.logger.warning(f"Could not remove existing {dest}: {e}")
            src.replace(dest)
            self.logger.info(f"Installed: {dest}")
            return True
        except Exception as e:
            self.logger.error(f"Move error: {e}")
            return False

    @staticmethod
    def dir_checksum(dir_path: Path) -> str:
        h = hashlib.sha256()
        for f in sorted(dir_path.rglob("*")):
            if f.is_file():
                h.update(f.relative_to(dir_path).as_posix().encode())
                h.update(f.read_bytes())
        return h.hexdigest()


def select_features(msi_features: set, logger: logging.Logger) -> list:
    matched = sorted(
        (fid for fid in msi_features if fid in FEATURE_MAP),
        key=lambda x: (0 if x in DEFAULT_FEATURES else 1, x)
    )

    selected: set = {fid for fid in matched if fid in DEFAULT_FEATURES}

    class Item:
        def __init__(self, fid, label):
            self.fid = fid
            self.label = label

    default_item = Item(None, "Default (all recommended features)")
    items = [default_item] + [Item(fid, FEATURE_MAP.get(fid, fid)) for fid in matched]
    cursor = [0]

    def all_defaults():
        return all(fid in selected for fid in matched if fid in DEFAULT_FEATURES)

    def format_list():
        formatted = []
        formatted.append(('class:header', '  Select NX features to install\n'))
        formatted.append(('class:header', '  ' + '\u2500' * 46 + '\n'))
        for i, item in enumerate(items):
            if item.fid is None:
                is_sel = all_defaults()
            else:
                is_sel = item.fid in selected
            is_cur = i == cursor[0]

            fid = item.fid if item.fid else "FEAT_DEFAULT"
            desc = item.label
            star = "*" if is_sel else " "
            cls = 'selected' if is_cur else ('default-item' if item.fid is None else ('selected-item' if is_sel else ''))
            style = f'class:{cls}' if cls else ''

            formatted.append(('class:checkbox', "["))
            formatted.append(('class:checkbox', star))
            formatted.append(('class:checkbox', "]  "))
            formatted.append((style, f"{fid:<30} {desc}\n"))

        count = sum(1 for it in items if (all_defaults() if it.fid is None else it.fid in selected))
        formatted.append(('class:footer', f'\n  {count} selected    ENTER confirm    UP/DOWN navigate    SPACE toggle\n'))
        return formatted

    def get_formatted_text():
        return format_list()

    layout = Layout(
        HSplit([
            Frame(
                Window(
                    FormattedTextControl(get_formatted_text),
                    wrap_lines=False,
                ),
                title="NX Feature Selection",
            ),
        ])
    )

    style = Style.from_dict({
        'header': 'bold',
        'footer': 'italic',
        'selected': 'reverse bg:#4444ff',
        'checkbox': 'fg:#007acc',
        'selected-item': 'fg:#00aa00',
        'default-item': 'fg:#00ff00',
    })

    kb = KeyBindings()
    app_ref = [None]

    @kb.add('up')
    def on_up(event):
        cursor[0] = max(0, cursor[0] - 1)
        app_ref[0].invalidate()

    @kb.add('down')
    def on_down(event):
        cursor[0] = min(len(items) - 1, cursor[0] + 1)
        app_ref[0].invalidate()

    @kb.add(' ')
    def on_space(event):
        item = items[cursor[0]]
        if item.fid is None:
            if all_defaults():
                for fid in matched:
                    if fid in DEFAULT_FEATURES:
                        selected.discard(fid)
            else:
                for fid in matched:
                    if fid in DEFAULT_FEATURES:
                        selected.add(fid)
        else:
            if item.fid in selected:
                selected.discard(item.fid)
            else:
                selected.add(item.fid)
        app_ref[0].invalidate()

    @kb.add('enter')
    def on_enter(event):
        event.app.exit(result=sorted(selected))

    @kb.add('c-c')
    @kb.add('escape')
    def on_quit(event):
        event.app.exit(result=None)

    app = Application(
        layout=layout,
        key_bindings=kb,
        style=style,
        full_screen=False,
        mouse_support=True,
        erase_when_done=True,
    )
    app_ref[0] = app
    result = app.run()

    if result is None:
        logger.warning("Feature selection cancelled.")
        return sorted(fid for fid in matched if fid in DEFAULT_FEATURES)
    if not result:
        logger.warning("No features selected. Will install ALL features.")
    return result


def parse_args():
    parser = argparse.ArgumentParser(
        description="Headless Siemens NX2506 Installation Script",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--config", "-c", default=None)
    parser.add_argument("--unattend", "-y", action="store_true")
    parser.add_argument("--validate-only", action="store_true")
    parser.add_argument("--uninstall", action="store_true",
                        help="Uninstall existing NX 2506 before installing")
    parser.add_argument("--skip-nx", action="store_true",
                        help="Skip NX installation (post-install steps only)")

    parser.add_argument("--version", action="version", version=f"%(prog)s {SCRIPT_VERSION}")
    return parser.parse_args()


def main():
    args = parse_args()
    script_dir = Path(__file__).parent.resolve()
    config_path = args.config or str(script_dir / "config.ini")

    if not Path(config_path).exists():
        print(f"ERROR: Config not found: {config_path}")
        sys.exit(1)

    config = Config(config_path)
    logger = setup_logging(config.log_level, config.temp_dir)
    logger.info(f"NX2506 Installer v{SCRIPT_VERSION}")
    logger.info(f"Config: {config_path}")

    if not check_admin_rights():
        logger.error("Must be run as Administrator.")
        sys.exit(1)

    logger.info("Checking license server...")
    if not check_license_server(config.license_server, logger=logger):
        logger.error("License server is not reachable. Fix the connection and try again.")
        sys.exit(1)

    if is_nx_installed() and not args.uninstall and not args.validate_only and not args.skip_nx:
        existing = get_msi_installed_location(MSI_PRODUCT_CODE)
        logger.warning(f"NX 2506 is already installed at: {existing}")
        logger.warning("Use --uninstall to remove it first.")
        sys.exit(1)

    if args.validate_only:
        sys.exit(0 if PostInstallValidator(config, logger).validate() else 1)

    if args.uninstall:
        if not uninstall_nx(MSI_PRODUCT_CODE, logger):
            logger.error("Uninstall failed.")
            sys.exit(1)
        logger.info("Uninstall complete. You can now run install.")
        sys.exit(0)

    if args.skip_nx:
        logger.info("Skipping NX installation (--skip-nx)")
    elif not args.unattend:
        installer = NXInstaller(config, logger)
        msi_features = installer._get_msi_features()
        if msi_features:
            selected = select_features(msi_features, logger)
        else:
            logger.debug("Could not read MSI feature table. Using default features.")
            selected = sorted(DEFAULT_FEATURES)
        logger.info("-" * 40)
    else:
        installer = NXInstaller(config, logger)
        selected = sorted(DEFAULT_FEATURES)
        logger.info("Installing default features (--unattend mode)")

    if args.skip_nx:
        logger.info("Skipping NX installation (--skip-nx)")
        fix_nx_permissions(config.install_dir, logger)
    else:
        if not installer.install(selected):
            logger.error("Installation failed.")
            sys.exit(1)
        fix_nx_permissions(config.install_dir, logger)

    if not LicenseConfigurator(config, logger).configure():
        logger.error("License configuration failed.")
        sys.exit(1)

    dl = FileDownloader(config, logger)

    fcc = dl.download(config.fcc_url, "fcc.xml")
    if not fcc:
        logger.error("fcc.xml download failed.")
        sys.exit(1)

    target_fcc = Path(config.install_dir) / "UGMANAGER" / "tccs" / "fcc.xml"
    if not dl.move(fcc, target_fcc):
        logger.error("fcc.xml move failed.")
        sys.exit(1)

    target_java = Path(config.install_dir).parent / "java"
    if target_java.exists():
        logger.info(f"Java directory already exists at {target_java}, skipping java installation.")
    else:
        java = dl.download(config.java_url, "java.zip")
        if not java:
            logger.error("java.zip download failed.")
            sys.exit(1)

        target_java.mkdir(parents=True, exist_ok=True)
        fix_nx_permissions(str(target_java.parent), logger)

        java_extracted = Path(config.temp_dir) / "java_extracted"
        if not dl.unzip(java, java_extracted):
            logger.error("java.zip unzip failed.")
            sys.exit(1)

        zulu_src = java_extracted
        if not dl.move(zulu_src, target_java / "zulu11"):
            logger.error("zulu11 move failed.")
            sys.exit(1)

    start_nx = dl.download(config.start_nx_url, "start_nx.bat")
    if not start_nx:
        logger.error("start_nx.bat download failed.")
        sys.exit(1)

    def transform_start_nx(content: str) -> str:
        name = config.name
        surname = config.surname
        if not name or not surname:
            logger.error("name and surname must both be set in config.ini")
            sys.exit(1)
        username = f"{name}.{surname}"
        password = f"{name}00"
        install_dir = config.install_dir.rstrip("\\")
        java_home = str(Path(install_dir) / "java" / "zulu11")
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

    target_start_nx = Path(config.install_dir).parent / "start_nx.bat"
    if not dl.transform(start_nx, target_start_nx, transform_start_nx):
        logger.error("start_nx.bat transform failed.")
        sys.exit(1)

    configure_role(config, logger)

    if not PostInstallValidator(config, logger).validate():
        logger.error("Post-install validation failed.")
        sys.exit(1)

    logger.info("Cleaning up temp directory...")
    shutil.rmtree(config.temp_dir, ignore_errors=True)

    logger.info("=" * 60)
    logger.info("Installation complete!")
    logger.info("=" * 60)
    sys.exit(0)


if __name__ == "__main__":
    main()
