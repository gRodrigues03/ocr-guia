import os
import sys
import shutil
import subprocess
from pathlib import Path
from dulwich import porcelain
from dulwich.repo import Repo
from dulwich.index import build_index_from_tree

PROJECT_NAME = "ocr-guia"
REPO_URL = b"https://github.com/gRodrigues03/ocr-guia.git"


def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

LOCAL_PATH = BASE_DIR / PROJECT_NAME
VENV_PATH = BASE_DIR / "venv"

if os.name == "nt":
    VENV_PYTHON = VENV_PATH / "Scripts" / "python.exe"
else:
    VENV_PYTHON = VENV_PATH / "bin" / "python"


def find_system_python():
    """
    Find a real Python interpreter to create the venv.
    """

    candidates = [
        "python3",
        "python",
        "py"
    ]

    for c in candidates:
        try:
            subprocess.check_output([c, "--version"])
            return c
        except Exception:
            pass

    return None


def ensure_venv():

    if VENV_PYTHON.exists():
        print("[BOOTSTRAP] Virtual environment exists")
        return

    if VENV_PATH.exists():
        shutil.rmtree(VENV_PATH)

    python = find_system_python()

    if not python:
        print("[BOOTSTRAP] ERROR: No Python interpreter found to create venv")
        sys.exit(1)

    print("[BOOTSTRAP] Creating virtual environment...")

    subprocess.check_call([
        python,
        "-m",
        "venv",
        str(VENV_PATH)
    ])


def install_packages():

    subprocess.check_call([
        str(VENV_PYTHON),
        "-m",
        "ensurepip",
        "--upgrade"
    ])

    print("[BOOTSTRAP] Installing")

    subprocess.check_call([
        str(VENV_PYTHON),
        "-c",
        "import sys; print(sys.version)"
    ])

    subprocess.check_call([
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


def update_repo():

    if not LOCAL_PATH.exists():
        print("[BOOTSTRAP] Cloning project...")
        porcelain.clone(REPO_URL, str(LOCAL_PATH))
        return

    print("[BOOTSTRAP] Re-cloning repository...")

    try:
        shutil.rmtree(LOCAL_PATH)
        porcelain.clone(REPO_URL, str(LOCAL_PATH))
        print("[BOOTSTRAP] Update complete")

    except Exception as e:
        print("[BOOTSTRAP] Update failed:", e)


def launch_app():

    main_script = LOCAL_PATH / "ocr-guia.py"

    if not main_script.exists():
        print("[BOOTSTRAP] ERROR: ocr-guia.py missing")
        sys.exit(1)

    subprocess.check_call([
        str(VENV_PYTHON),
        str(main_script)
    ])


def main():

    print("[BOOTSTRAP] Base directory:", BASE_DIR)
    print(sys.version)

    ensure_venv()
    update_repo()
    install_packages()
    launch_app()


if __name__ == "__main__":
    main()