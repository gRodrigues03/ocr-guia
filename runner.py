import os
import sys
import shutil
import subprocess
import ctypes
import json
import zipfile
from pathlib import Path

import requests


PROJECT_NAME = "ocr-guia"
REPO_ZIP = "https://github.com/gRodrigues03/ocr-guia/archive/refs/heads/main.zip"
COMMIT_API = "https://api.github.com/repos/gRodrigues03/ocr-guia/commits/main"

UV_URL = "https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip"

BOOTSTRAP_MUTEX = None

FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


# ---------------- PATHS ----------------

def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()

LOCAL_PATH = BASE_DIR / PROJECT_NAME

STATE_FILE = BASE_DIR / ".bootstrap_state.json"

UV_DIR = BASE_DIR / "uv"
UV_EXE = UV_DIR / "uv.exe"


# ---------------- UTILS ----------------

def run_hidden(cmd, cwd=None):
    subprocess.check_call(cmd, cwd=cwd, creationflags=FLAGS)


def acquire_lock():
    global BOOTSTRAP_MUTEX

    BOOTSTRAP_MUTEX = ctypes.windll.kernel32.CreateMutexW(None, False, "OCR_GUIA_MUTEX")

    if ctypes.windll.kernel32.GetLastError() == 183:
        sys.exit(0)


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


# ---------------- UV ----------------

def ensure_uv():

    if UV_EXE.exists():
        return

    tmp_zip = BASE_DIR / "uv.zip"

    r = requests.get(UV_URL, timeout=60)

    with open(tmp_zip, "wb") as f:
        f.write(r.content)

    with zipfile.ZipFile(tmp_zip, "r") as z:
        z.extractall(UV_DIR)

    tmp_zip.unlink()


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


# ---------------- UV SYNC ----------------

def sync_env():

    run_hidden(
        [str(UV_EXE), "sync"],
        cwd=LOCAL_PATH
    )


# ---------------- LAUNCH ----------------

def launch_app():

    main_script = LOCAL_PATH / "ocr-guia.py"

    if not main_script.exists():
        sys.exit(1)

    subprocess.Popen(
        [str(UV_EXE), "run", str(main_script)],
        cwd=LOCAL_PATH,
        creationflags=FLAGS
    )


# ---------------- MAIN ----------------

def main():

    acquire_lock()

    ensure_uv()
    update_repo()
    sync_env()
    launch_app()


if __name__ == "__main__":
    main()