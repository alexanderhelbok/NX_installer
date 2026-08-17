"""Build nx_gui.exe from nx_gui.py using PyInstaller."""

import subprocess
import sys
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
BUILD = ROOT / "build"
SPEC = ROOT / "nx_gui.spec"

PYINSTALLER_ARGS = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onefile",
    "--uac-admin",
    "--name", "nx_gui",
    "--add-data", f"config.ini{';' if sys.platform == 'win32' else ':'}.",
    "--hidden-import", "gdown",
    "--exclude-module", "PyQt6",
    "--distpath", str(DIST),
    "--workpath", str(BUILD),
    "--specpath", str(ROOT),
    str(ROOT / "nx_gui.py"),
]


def clean():
    for d in (DIST, BUILD):
        if d.exists():
            shutil.rmtree(d)
    if SPEC.exists():
        SPEC.unlink()


def build():
    print("Building nx_gui.exe ...")
    result = subprocess.run(PYINSTALLER_ARGS)
    if result.returncode != 0:
        print(f"Build failed (exit code {result.returncode})")
        sys.exit(result.returncode)
    exe = DIST / "nx_gui.exe"
    if exe.exists():
        size_mb = exe.stat().st_size / (1024 * 1024)
        print(f"Done: {exe} ({size_mb:.1f} MB)")
    else:
        print("Build completed but exe not found")
        sys.exit(1)


if __name__ == "__main__":
    clean()
    build()
