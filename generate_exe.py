"""
generate_exe.py -- builds the standalone Windows .exe for the SAF-T ->
Excel GUI (saft_to_excel_gui_tk.py) via PyInstaller.

    .venv\\Scripts\\python generate_exe.py

A onefile, windowed build (no console window). saft_to_excel.py is a
plain top-level `import` in the GUI's own code (see saft_to_excel_gui_tk.py's
_run_worker/_worker_command -- the GUI relaunches itself as a child process
with `--convert <in> <out>` rather than shelling out to saft_to_excel.py by
path, since a frozen onefile exe has no separate interpreter or script
file to point at), so PyInstaller's own import analysis bundles it
automatically -- no --add-data needed and no hand-written .spec file.

Excludes PIL/ssl/_hashlib (see EXCLUDED_MODULES): openpyxl's optional
image support pulls in Pillow defensively even though the GUI never
touches an image, and it does no networking/hashing that needs OpenSSL --
together a measured ~16MB of dead weight in the uncompressed bundle (PIL
10.7MB, the OpenSSL DLLs 5.9MB) before this. Verify after any further
dependency exclusion that both the --convert worker and the GUI itself
still run -- an exclude that turns out to be load-bearing fails at import
time inside the frozen exe, not at build time.

No --icon is passed, so the built .exe uses Tk's own default window icon.

Safe to run from any working directory: every path below is anchored on
this file's own location, not the cwd. Re-running overwrites the previous
build without prompting.
"""

import shutil
import subprocess
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
GUI_SCRIPT = TOOLS_DIR / "saft_to_excel_gui_tk.py"
APP_NAME = "Conversor SAFT em Excel"

# Bundled by PyInstaller's default hooks despite the GUI never using them
# (confirmed by testing: both --convert and the GUI itself still work
# without them) -- see this file's own docstring for the measured savings.
EXCLUDED_MODULES = ["PIL", "ssl", "_hashlib"]


def build_exe() -> None:
    # `sys.executable -m PyInstaller`, not a bare `pyinstaller`, so the
    # build always uses the PyInstaller installed in whichever interpreter
    # launched this script (the .venv), regardless of what's on PATH.
    # `--noconfirm` skips the "wipe dist/?" prompt on a rebuild.
    command = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", APP_NAME,
    ]
    for module in EXCLUDED_MODULES:
        command += ["--exclude-module", module]
    command.append(str(GUI_SCRIPT))

    subprocess.run(command, cwd=TOOLS_DIR, check=True)

    built = TOOLS_DIR / "dist" / f"{APP_NAME}.exe"
    if not built.is_file():
        raise SystemExit(f"PyInstaller reported success but {built} is missing.")

    destino = TOOLS_DIR / f"{APP_NAME}.exe"
    shutil.copy2(built, destino)
    size_mb = destino.stat().st_size / 1024 / 1024
    print(f"\nCopied {built.relative_to(TOOLS_DIR)} -> {destino.relative_to(TOOLS_DIR)} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    build_exe()
