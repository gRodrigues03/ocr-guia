import os
import sys
import shutil
import subprocess
import ctypes
from pathlib import Path
from dulwich import porcelain

PROJECT_NAME = "ocr-guia"
REPO_URL = b"https://github.com/gRodrigues03/ocr-guia.git"

# Keep mutex global so it isn't garbage collected
BOOTSTRAP_MUTEX = None

FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def run_hidden(cmd):
    subprocess.check_call(cmd, creationflags=FLAGS)


def acquire_lock():
    global BOOTSTRAP_MUTEX
    BOOTSTRAP_MUTEX = ctypes.windll.kernel32.CreateMutexW(None, False, "MyBootstrapMutex")

    if ctypes.windll.kernel32.GetLastError() == 183:
        sys.exit(0)


def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

LOCAL_PATH = BASE_DIR / PROJECT_NAME
VENV_PATH = BASE_DIR / "venv"

if os.name == "nt":
    VENV_PYTHON = VENV_PATH / "Scripts" / "python.exe"
    VENV_PYTHONW = VENV_PATH / "Scripts" / "pythonw.exe"
else:
    VENV_PYTHON = VENV_PATH / "bin" / "python"
    VENV_PYTHONW = VENV_PATH / "bin" / "python"


def find_system_python():

    candidates = ["py", "python3", "python"]

    for c in candidates:
        try:
            subprocess.check_call(
                [c, "--version"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=FLAGS
            )
            return c
        except Exception:
            pass

    return None


def ensure_venv():

    if VENV_PYTHON.exists():
        return

    if VENV_PATH.exists():
        shutil.rmtree(VENV_PATH)

    python = find_system_python()

    if not python:
        sys.exit(1)

    run_hidden([
        python,
        "-m",
        "venv",
        str(VENV_PATH)
    ])


def update_repo():

    if not LOCAL_PATH.exists():
        porcelain.clone(REPO_URL, str(LOCAL_PATH))
        return

    try:
        shutil.rmtree(LOCAL_PATH)
        porcelain.clone(REPO_URL, str(LOCAL_PATH))
    except Exception:
        pass


def install_packages():

    run_hidden([
        str(VENV_PYTHON),
        "-m",
        "ensurepip",
        "--upgrade"
    ])

    run_hidden([
        str(VENV_PYTHON),
        "-m",
        "pip",
        "install",
        "--upgrade",
        "--upgrade-strategy",
        "only-if-needed",
        "-r",
        str(LOCAL_PATH / "requirements.txt")
    ])


def launch_app():

    main_script = LOCAL_PATH / "ocr-guia.py"

    if not main_script.exists():
        sys.exit(1)

    subprocess.Popen(
        [str(VENV_PYTHONW), str(main_script)],
        creationflags=FLAGS
    )


def main():

    acquire_lock()

    ensure_venv()
    update_repo()
    install_packages()
    launch_app()


if __name__ == "__main__":
    main()