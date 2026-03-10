import os
import sys
import shutil
import subprocess
import ctypes
import hashlib
import json
import zipfile
from pathlib import Path

import requests


PROJECT_NAME = "ocr-guia"
REPO_ZIP = "https://github.com/gRodrigues03/ocr-guia/archive/refs/heads/main.zip"
COMMIT_API = "https://api.github.com/repos/gRodrigues03/ocr-guia/commits/main"

BOOTSTRAP_MUTEX = None

FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


# ---------------- PATHS ----------------

def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

LOCAL_PATH = BASE_DIR / PROJECT_NAME
VENV_PATH = BASE_DIR / "venv"

STATE_FILE = BASE_DIR / ".bootstrap_state.json"


if os.name == "nt":
    VENV_PYTHON = VENV_PATH / "Scripts" / "python.exe"
    VENV_PYTHONW = VENV_PATH / "Scripts" / "pythonw.exe"
else:
    VENV_PYTHON = VENV_PATH / "bin" / "python"
    VENV_PYTHONW = VENV_PATH / "bin" / "python"


# ---------------- UTILS ----------------

def run_hidden(cmd):
    subprocess.check_call(cmd, creationflags=FLAGS)


def acquire_lock():
    global BOOTSTRAP_MUTEX

    BOOTSTRAP_MUTEX = ctypes.windll.kernel32.CreateMutexW(None, False, "OCR_GUIA_MUTEX")

    if ctypes.windll.kernel32.GetLastError() == 183:
        sys.exit(0)


def find_system_python():

    for c in ["py", "python3", "python"]:
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


# ---------------- STATE ----------------

def load_state():

    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text())
        except Exception:
            pass

    return {}


def save_state(state):
    STATE_FILE.write_text(json.dumps(state))


# ---------------- HASH ----------------

def hash_file(path):

    h = hashlib.sha256()

    with open(path, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)

    return h.hexdigest()


# ---------------- VENV ----------------

def ensure_venv():

    if VENV_PYTHON.exists():
        return

    if VENV_PATH.exists():
        shutil.rmtree(VENV_PATH)

    python = find_system_python()

    if not python:
        sys.exit(1)

    run_hidden([python, "-m", "venv", str(VENV_PATH)])


# ---------------- GITHUB ----------------

def get_remote_commit():

    try:
        r = requests.get(COMMIT_API, timeout=10)
        r.raise_for_status()
        return r.json()["sha"]
    except Exception:
        return None


def download_repo():

    tmp_zip = BASE_DIR / "repo.zip"
    tmp_extract = BASE_DIR / "repo_tmp"

    if tmp_zip.exists():
        tmp_zip.unlink()

    if tmp_extract.exists():
        shutil.rmtree(tmp_extract)

    r = requests.get(REPO_ZIP, timeout=60)

    with open(tmp_zip, "wb") as f:
        f.write(r.content)

    with zipfile.ZipFile(tmp_zip, "r") as z:
        z.extractall(tmp_extract)

    extracted = next(tmp_extract.iterdir())

    if LOCAL_PATH.exists():
        shutil.rmtree(LOCAL_PATH)

    extracted.rename(LOCAL_PATH)

    shutil.rmtree(tmp_extract)
    tmp_zip.unlink()


# ---------------- UPDATE ----------------

def update_repo():

    state = load_state()

    remote_commit = get_remote_commit()

    if not remote_commit:
        return

    if state.get("commit") == remote_commit and LOCAL_PATH.exists():
        return

    download_repo()

    state["commit"] = remote_commit
    save_state(state)


# ---------------- PIP ----------------

def install_packages():

    req = LOCAL_PATH / "requirements.txt"

    if not req.exists():
        return

    state = load_state()

    new_hash = hash_file(req)

    if state.get("req_hash") == new_hash:
        return

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
        str(req)
    ])

    state["req_hash"] = new_hash
    save_state(state)


# ---------------- LAUNCH ----------------

def launch_app():

    main_script = LOCAL_PATH / "ocr-guia.py"

    if not main_script.exists():
        sys.exit(1)

    subprocess.Popen(
        [str(VENV_PYTHONW), str(main_script)],
        creationflags=FLAGS
    )


# ---------------- MAIN ----------------

def main():

    acquire_lock()

    ensure_venv()
    update_repo()
    install_packages()
    launch_app()


if __name__ == "__main__":
    main()