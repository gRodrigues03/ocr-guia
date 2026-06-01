import time
import re
import threading
import queue
import socket
import json
import struct
import uuid
from pathlib import Path

import fitz
import numpy as np
from rapidocr_onnxruntime import RapidOCR

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import customtkinter as ctk
from tkinter import filedialog

import requests
import sys


# ---------------- THEME ----------------

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

empresas_index = {
    'PONTE': '001',
    'GLORIA': '002',
    'GARDEL': '003'
}

# Canonical UNC root that all nodes can reach regardless of local drive mapping.
# Paths sent over the network are always expressed relative to this base.
UNC_BASE = Path(r"\\148.1.1.231\shared\ARQUIVO\!GUIAS")


def _resolve_unc_share(drive: str) -> "Path | None":
    """
    Use WNetGetConnection to find the UNC share a drive letter maps to.
    Returns the UNC share root (e.g. the path to //148.1.1.231/shared) for 'X:'.
    Returns None on any failure (non-Windows, unmapped drive, etc.).
    """
    try:
        import ctypes
        buf  = ctypes.create_unicode_buffer(512)
        size = ctypes.c_ulong(512)
        ret  = ctypes.windll.mpr.WNetGetConnectionW(drive, buf, ctypes.byref(size))
        if ret == 0:
            return Path(buf.value)
    except Exception:
        pass
    return None


def to_unc(p: Path) -> Path:
    """
    Rewrite a local path to its canonical UNC form under UNC_BASE.

    All valid mount layouts produce a unc_full that falls under UNC_BASE:
      X: = UNC_BASE     ->  X:/GLORIA/...            -> UNC_BASE/GLORIA/...
      X: = share root   ->  X:/ARQUIVO/!GUIAS/GLORIA -> UNC_BASE/GLORIA/...
      X: = one up       ->  X:/!GUIAS/GLORIA/...     -> UNC_BASE/GLORIA/...
      Already UNC_BASE  ->  returned as-is

    If the path is not under our share hierarchy, it is returned unchanged.
    """
    # 1. Already correct
    try:
        return UNC_BASE / p.relative_to(UNC_BASE)
    except ValueError:
        pass

    # 2. Build the UNC equivalent of p
    unc_full: "Path | None" = None

    if p.drive and not p.drive.startswith("\\\\"):
        # Drive-letter path: resolve via WNetGetConnection
        share = _resolve_unc_share(p.drive)
        if share is not None:
            unc_full = share / p.relative_to(p.anchor)
    elif p.drive.startswith("\\\\"):
        # Already a UNC path on some share
        unc_full = p

    # 3. Relativise against UNC_BASE
    if unc_full is not None:
        try:
            return UNC_BASE / unc_full.relative_to(UNC_BASE)
        except ValueError:
            pass

    # Not under our share hierarchy — return as-is
    return p


def resource_path(filename):
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / filename


# ---------------- API ----------------

def consultar_api(id_, empresa, mes):
    url = f"http://148.1.1.239:8501/nguia?id={id_}&mes={mes}&empresa={empresas_index.get(empresa, None)}"
    r = requests.get(url, timeout=5)
    r.raise_for_status()
    return r.text


def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    else:
        return Path(sys.argv[0]).resolve().parent.parent


def get_threads():
    threads_file = get_base_dir() / "threads.txt"
    if threads_file.exists():
        try:
            thread_count = int(threads_file.read_text().strip())
            return 12 if thread_count > 12 else thread_count if thread_count > 0 else 2
        except Exception:
            return 2
    return 2


# ---------------- CONFIG ----------------

regex_guia = re.compile(r"\d{5,6}", re.IGNORECASE | re.DOTALL)
regex_mes  = re.compile(r"[\\/](\d{4})[\\/](\d{2})\s*-")

ocr = RapidOCR(use_cls=False)

observer    = None
pasta_atual = None

fila     = queue.Queue()
ui_queue = queue.Queue()

stop_event = threading.Event()

# Node identity
NODE_ID   = str(uuid.uuid4())[:8]
NODE_NAME = socket.gethostname()

# Network ports
UDP_PORT = 55100   # broadcast / discovery
TCP_PORT = 55101   # work distribution


# ---------------- PATH DATE ----------------

def extrair_data_do_path(path: Path):
    match = regex_mes.search(str(path))
    if not match:
        return None, None
    ano = match.group(1)
    mes = match.group(2)
    dia = path.parent.name
    if not dia.isdigit():
        return None, None
    data      = f"{ano}-{mes}-{int(dia):02d}"
    parts     = path.parts
    idx_ano   = parts.index(ano)
    pasta_raiz = parts[idx_ano - 1]
    return pasta_raiz, data


# ---------------- OCR ----------------

def extrair_guia(pdf_path, empresa, data_ref):
    try:
        doc  = fitz.open(pdf_path)
        page = doc[0]
        rect = page.rect

        if data_ref >= "2025-08-01":
            clip = fitz.Rect(rect.width * 0.1, 0, rect.width * 0.9, rect.height * 0.4)
        else:
            clip = fitz.Rect(rect.width * 0.55, 0, rect.width, rect.height * 0.28)

        pix   = fitz.Pixmap(doc, page.get_images()[0][0])
        scale = pix.width / rect.width
        x0, y0, x1, y1 = (clip * scale)

        buf = np.frombuffer(pix.samples_mv, dtype=np.uint8)

        img = buf.reshape(pix.height, pix.width, pix.n)
        img = img[int(y0):int(y1), int(x0):int(x1)]

        resultado, _ = ocr(img)
        texto = "\n".join([r[1] for r in resultado]) if resultado else ""
        match = regex_guia.findall(texto)

        if match:
            for m in match:
                text = consultar_api(m, empresa, data_ref)
                if text and len(text):
                    return text.split(',')[0], texto

        return None, texto

    except Exception as e:
        ui_queue.put(("log", f"Erro em {Path(pdf_path).name}: {e}"))

    return None, None


# ---------------- PROCESSING ----------------

def esperar_arquivo_finalizar(path, retries: int = 6, retry_delay: float = 1.0):
    """
    Wait until the file stops growing (write complete).
    On network paths a FileNotFoundError may be transient, so retry a few
    times before giving up.
    """
    not_found_count = 0
    tamanho = -1
    while True:
        try:
            novo = path.stat().st_size
            not_found_count = 0
        except FileNotFoundError:
            not_found_count += 1
            if not_found_count >= retries:
                return False
            time.sleep(retry_delay)
            continue
        if novo == tamanho:
            return True
        tamanho = novo
        time.sleep(0.5)


def processar_pdf(pdf: Path, *, silent: bool = False) -> tuple[str, str, str]:
    if not esperar_arquivo_finalizar(pdf):
        if not silent:
            ui_queue.put(("error", f"Arquivo não encontrado: {pdf.name}"))
        return "error", str(pdf), "Arquivo não encontrado"

    empresa, data_ref = extrair_data_do_path(pdf)
    if not data_ref:
        if not silent:
            ui_queue.put(("error", f"Data não encontrada: {pdf.name}"))
        return "error", str(pdf), "Data não encontrada"

    guia, texto = extrair_guia(pdf, empresa, data_ref)

    if guia:
        novo_nome = pdf.with_name(f"{guia}.pdf")
        contador  = 1
        while novo_nome.exists():
            novo_nome = pdf.with_name(f"{guia} ({contador}).pdf")
            contador += 1
        pdf.rename(novo_nome)
        if not silent:
            ui_queue.put(("renamed", pdf.name, novo_nome.name))
        return "renamed", str(pdf), novo_nome.name
    else:
        novo_nome = pdf.with_name(f"NO_OCR {pdf.name}")
        contador  = 1
        while novo_nome.exists():
            novo_nome = pdf.with_name(f"NO_OCR {contador} {pdf.name}")
            contador += 1
        pdf.rename(novo_nome)
        if not silent:
            ui_queue.put(("notfound", novo_nome.name))
        return "notfound", str(pdf), novo_nome.name


def worker():
    while not stop_event.is_set():
        # Reserve 8 slots per connected collaborator so they always have work.
        # Local workers only consume the surplus beyond that reservation.
        n_collabs = len(master_server._workers) if master_server else 0
        reserved  = n_collabs * 8

        if reserved > 0 and fila.qsize() <= reserved:
            time.sleep(1)
            continue

        try:
            pdf = fila.get(timeout=0.5)
        except queue.Empty:
            continue
        processar_pdf(pdf)
        fila.task_done()


# ---------------- WATCHDOG ----------------

class Handler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if (
            path.suffix.lower() == ".pdf"
            and path.name
            and not path.name[0].isdigit()
            and not path.name.startswith("NO_OCR ")
        ):
            fila.put(path)


def iniciar_observer():
    global observer
    if observer:
        observer.stop()
        observer.join()

    handler  = Handler()
    observer = Observer()
    observer.schedule(handler, str(pasta_atual), recursive=True)
    observer.start()

    ui_queue.put(("log", f"Observando: {pasta_atual}"))

    for pdf in pasta_atual.glob("*/*.pdf"):
        if pdf.name and not pdf.name[0].isdigit() and not pdf.name.startswith("NO_OCR "):
            fila.put(pdf)


# ================================================================
#  P2P NETWORK — MASTER SIDE
# ================================================================

class MasterServer:
    """
    Runs two background threads:
      1. UDP broadcaster  — announces this master every 3 s
      2. TCP work server  — accepts worker connections, hands out jobs, receives reports
    """

    def __init__(self):
        self._stop  = threading.Event()
        self._lock  = threading.Lock()
        self._in_flight: dict[str, Path] = {}   # token -> pdf path
        self._workers: list[str] = []            # connected worker names

    # ---- public ----

    def start(self):
        threading.Thread(target=self._udp_broadcaster, daemon=True, name="udp-bc").start()
        threading.Thread(target=self._tcp_server,      daemon=True, name="tcp-srv").start()
        ui_queue.put(("log", f"[Rede] Mestre ativo — {NODE_NAME} ({self._local_ip()})"))

    def stop(self):
        self._stop.set()

    def connected_workers(self):
        return list(self._workers)

    # ---- UDP broadcast ----

    def _local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    def _udp_broadcaster(self):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        payload = json.dumps({
            "type": "MASTER_ANNOUNCE",
            "id":   NODE_ID,
            "name": NODE_NAME,
            "ip":   self._local_ip(),
            "port": TCP_PORT,
        }).encode()

        while not self._stop.is_set():
            try:
                sock.sendto(payload, ("<broadcast>", UDP_PORT))
            except Exception:
                pass
            time.sleep(3)

        sock.close()

    # ---- TCP work server ----

    def _tcp_server(self):
        srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        srv.bind(("0.0.0.0", TCP_PORT))
        srv.listen(16)
        srv.settimeout(1.0)

        while not self._stop.is_set():
            try:
                conn, addr = srv.accept()
            except socket.timeout:
                continue
            threading.Thread(
                target=self._handle_worker,
                args=(conn, addr),
                daemon=True
            ).start()

        srv.close()

    def _handle_worker(self, conn: socket.socket, addr):
        worker_name = f"{addr[0]}:{addr[1]}"
        with self._lock:
            self._workers.append(worker_name)
        ui_queue.put(("log", f"[Rede] Colaborador conectado: {worker_name}"))
        ui_queue.put(("net_workers", self._workers))

        try:
            conn.settimeout(60)
            while not self._stop.is_set():
                msg = self._recv_msg(conn)
                if msg is None:
                    break

                if msg.get("cmd") == "GET_WORK":
                    jobs = self._pop_jobs(8)
                    self._send_msg(conn, {"jobs": jobs})

                elif msg.get("cmd") == "REPORT":
                    for r in msg.get("results", []):
                        status   = r.get("status")
                        original = r.get("original", "")
                        result   = r.get("result",   "")
                        if status == "renamed":
                            ui_queue.put(("renamed", Path(original).name, result))
                        elif status == "notfound":
                            ui_queue.put(("notfound", result))
                        elif status == "error":
                            ui_queue.put(("error", result))

        except Exception as e:
            ui_queue.put(("log", f"[Rede] Conexão perdida {worker_name}: {e}"))
        finally:
            with self._lock:
                if worker_name in self._workers:
                    self._workers.remove(worker_name)
            ui_queue.put(("log", f"[Rede] Colaborador desconectado: {worker_name}"))
            ui_queue.put(("net_workers", self._workers))
            conn.close()

    def _pop_jobs(self, n: int) -> list[str]:
        jobs = []
        for _ in range(n):
            try:
                pdf: Path = fila.get_nowait()
                jobs.append(str(to_unc(pdf)))   # normalise to UNC for collaborators
                fila.task_done()
            except queue.Empty:
                break
        return jobs

    # ---- framing helpers (4-byte length prefix) ----

    @staticmethod
    def _send_msg(sock: socket.socket, obj: dict):
        data = json.dumps(obj).encode()
        sock.sendall(struct.pack(">I", len(data)) + data)

    @staticmethod
    def _recv_msg(sock: socket.socket):
        raw = MasterServer._recvn(sock, 4)
        if not raw:
            return None
        length = struct.unpack(">I", raw)[0]
        data   = MasterServer._recvn(sock, length)
        if not data:
            return None
        return json.loads(data.decode())

    @staticmethod
    def _recvn(sock: socket.socket, n: int):
        buf = b""
        while len(buf) < n:
            chunk = sock.recv(n - len(buf))
            if not chunk:
                return None
            buf += chunk
        return buf


# ================================================================
#  P2P NETWORK — WORKER (COLLABORATOR) SIDE
# ================================================================

# Shared state visible to UI
discovered_masters: dict[str, dict] = {}   # id -> {name, ip, port, last_seen}
_disc_lock = threading.Lock()

active_master_id: str | None = None        # which master this node is serving
_worker_stop = threading.Event()


def _udp_listener():
    """Listens for UDP broadcasts and updates discovered_masters."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEPORT, 1)
    except AttributeError:
        pass
    sock.bind(("", UDP_PORT))
    sock.settimeout(1.0)

    while not stop_event.is_set():
        try:
            data, addr = sock.recvfrom(1024)
            msg = json.loads(data.decode())
            if msg.get("type") == "MASTER_ANNOUNCE" and msg.get("id") != NODE_ID:
                with _disc_lock:
                    discovered_masters[msg["id"]] = {
                        "name":      msg["name"],
                        "ip":        msg["ip"],
                        "port":      msg["port"],
                        "last_seen": time.time(),
                    }
                ui_queue.put(("masters_updated", dict(discovered_masters)))
        except (socket.timeout, json.JSONDecodeError):
            pass
        except Exception as e:
            ui_queue.put(("log", f"[Rede] Erro UDP listener: {e}"))

    sock.close()


def _purge_old_masters():
    """Removes masters not seen for > 10 s."""
    while not stop_event.is_set():
        time.sleep(5)
        now = time.time()
        with _disc_lock:
            stale = [k for k, v in discovered_masters.items() if now - v["last_seen"] > 10]
            for k in stale:
                del discovered_masters[k]
        if stale:
            ui_queue.put(("masters_updated", dict(discovered_masters)))


def start_discovery():
    threading.Thread(target=_udp_listener,    daemon=True, name="udp-listen").start()
    threading.Thread(target=_purge_old_masters, daemon=True, name="purge").start()


def connect_to_master(master_id: str):
    """
    Connects to the selected master and processes work in a loop.
    Runs in its own daemon thread.
    """
    global active_master_id
    active_master_id = master_id
    _worker_stop.clear()
    threading.Thread(
        target=_collaborator_loop,
        args=(master_id,),
        daemon=True,
        name="collab-loop"
    ).start()


def disconnect_from_master():
    global active_master_id
    _worker_stop.set()
    active_master_id = None


def _send_msg_raw(sock, obj):
    data = json.dumps(obj).encode()
    sock.sendall(struct.pack(">I", len(data)) + data)

def _recv_msg_raw(sock):
    raw = _recvn_raw(sock, 4)
    if not raw:
        return None
    length = struct.unpack(">I", raw)[0]
    data   = _recvn_raw(sock, length)
    if not data:
        return None
    return json.loads(data.decode())

def _recvn_raw(sock, n):
    buf = b""
    while len(buf) < n:
        chunk = sock.recv(n - len(buf))
        if not chunk:
            return None
        buf += chunk
    return buf


def _collaborator_loop(master_id: str):
    with _disc_lock:
        info = discovered_masters.get(master_id)

    if not info:
        ui_queue.put(("log", "[Rede] Mestre não encontrado."))
        return

    ip, port = info["ip"], info["port"]
    ui_queue.put(("log", f"[Rede] Conectando ao mestre {info['name']} ({ip}:{port})…"))

    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        sock.settimeout(30)
        ui_queue.put(("log", f"[Rede] Conectado a {info['name']}"))
        ui_queue.put(("collab_status", "connected", info["name"]))
    except Exception as e:
        ui_queue.put(("log", f"[Rede] Falha ao conectar: {e}"))
        ui_queue.put(("collab_status", "disconnected", ""))
        return

    try:
        while not _worker_stop.is_set():
            # Ask for work
            _send_msg_raw(sock, {"cmd": "GET_WORK"})
            response = _recv_msg_raw(sock)

            if response is None:
                ui_queue.put(("log", "[Rede] Conexão encerrada pelo mestre."))
                break

            jobs: list[str] = response.get("jobs", [])

            if not jobs:
                time.sleep(2)
                continue

            ui_queue.put(("log", f"[Rede] Recebidos {len(jobs)} trabalho(s)."))

            # Process jobs using local worker threads.
            # Explicit argument passing avoids closure-capture bugs across iterations.
            batch_q       = queue.Queue()
            batch_results: list[dict] = []
            batch_lock    = threading.Lock()

            for j in jobs:
                batch_q.put(j)

            def _collab_worker(work_q, out, lock):
                while True:
                    try:
                        pdf_str = work_q.get(block=False)
                    except queue.Empty:
                        break
                    status, original, result = processar_pdf(Path(pdf_str), silent=True)
                    with lock:
                        out.append({
                            "status":   status,
                            "original": original,
                            "result":   result,
                        })
                    # Mirror result to the network log on the collaborator side
                    if status == "renamed":
                        ui_queue.put(("net_log", f"{Path(original).name} → {result}"))
                    elif status == "notfound":
                        ui_queue.put(("net_log", f"NO OCR: {result}"))
                    elif status == "error":
                        ui_queue.put(("net_log", f"Erro: {result}"))
                    work_q.task_done()

            threads_n     = get_threads()
            batch_threads = [
                threading.Thread(
                    target=_collab_worker,
                    args=(batch_q, batch_results, batch_lock),
                    daemon=True,
                )
                for _ in range(min(threads_n, len(jobs)))
            ]
            for t in batch_threads:
                t.start()
            for t in batch_threads:
                t.join()

            # Report results back to master
            _send_msg_raw(sock, {"cmd": "REPORT", "results": batch_results})

    except Exception as e:
        ui_queue.put(("log", f"[Rede] Erro no colaborador: {e}"))
    finally:
        sock.close()
        ui_queue.put(("collab_status", "disconnected", ""))
        ui_queue.put(("log", "[Rede] Desconectado do mestre."))


# ================================================================
#  UI
# ================================================================

master_server: MasterServer | None = None


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("OCR Renomeador")
        self.geometry("680x520")

        try:
            ico = resource_path("exeicon.ico")
            png = resource_path("trayicon.png")
            if ico.exists():
                self.iconbitmap(ico)
            elif png.exists():
                from PIL import ImageTk, Image
                img = ImageTk.PhotoImage(Image.open(png))
                self.iconphoto(True, img)
                self._icon = img
        except Exception:
            pass

        self.processed = 0
        self.errors    = 0

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabs = ctk.CTkTabview(self, command=self._on_tab_change)
        self.tabs.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tabs.add("Mestre")
        self.tabs.add("Colaborador")

        self._build_master_tab()
        self._build_collab_tab()

        self.after(200, self.update_ui)

    # ------------------------------------------------------------------ MASTER TAB

    def _build_master_tab(self):
        tab = self.tabs.tab("Mestre")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(4, weight=1)

        # Top bar: folder button + stop button side by side
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(10, 6))

        self.btn_pasta = ctk.CTkButton(
            top, text="Selecionar pasta", command=self.selecionar_pasta
        )
        self.btn_pasta.pack(side="left")

        self.btn_master_stop = ctk.CTkButton(
            top, text="Parar",
            width=80, fg_color="#c0392b", hover_color="#922b21",
            state="disabled", command=self._stop_master
        )
        self.btn_master_stop.pack(side="left", padx=(8, 0))

        self.status = ctk.CTkLabel(
            tab,
            text="Selecione uma pasta para começar e anunciar-se como mestre na rede."
        )
        self.status.grid(row=1, column=0, padx=0, sticky="w")

        self.queue_label = ctk.CTkLabel(tab, text="Fila: 0")
        self.queue_label.grid(row=2, column=0, padx=0, sticky="w")

        self.counter = ctk.CTkLabel(tab, text="Processados: 0 | Erros: 0")
        self.counter.grid(row=3, column=0, padx=0, sticky="w")

        self.log = ctk.CTkTextbox(tab)
        self.log.grid(row=4, column=0, padx=0, pady=(6, 0), sticky="nsew")

    def selecionar_pasta(self):
        global pasta_atual, fila, master_server
        pasta = filedialog.askdirectory()
        if not pasta:
            return
        pasta_atual = Path(pasta)
        fila = queue.Queue()
        iniciar_observer()
        # Automatically become master as soon as a folder is chosen
        if master_server:
            master_server.stop()
        master_server = MasterServer()
        master_server.start()
        self.btn_pasta.configure(text="Trocar pasta")
        self.btn_master_stop.configure(state="normal")

    def _stop_master(self):
        global master_server, pasta_atual
        if master_server:
            master_server.stop()
            master_server = None
        if observer:
            observer.stop()
        pasta_atual = None
        self.status.configure(
            text="Selecione uma pasta para começar e anunciar-se como mestre na rede."
        )
        self.btn_pasta.configure(text="Selecionar pasta")
        self.btn_master_stop.configure(state="disabled")

    # ------------------------------------------------------------------ COLLAB TAB

    def _build_collab_tab(self):
        tab = self.tabs.tab("Colaborador")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        self.collab_status_label = ctk.CTkLabel(
            tab, text="Buscando mestres na rede…"
        )
        self.collab_status_label.grid(row=0, column=0, sticky="w", pady=(10, 4))

        # Scrollable list of discovered masters
        self.master_list_frame = ctk.CTkScrollableFrame(tab, label_text="Mestres disponíveis")
        self.master_list_frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        self.master_list_frame.grid_columnconfigure(0, weight=1)
        self._master_btns: dict[str, ctk.CTkButton] = {}

        self.net_log = ctk.CTkTextbox(tab)
        self.net_log.grid(row=2, column=0, sticky="nsew", pady=(0, 4))

    # ------------------------------------------------------------------ TAB SWITCH

    def _on_tab_change(self):
        tab = self.tabs.get()
        if tab == "Colaborador":
            # Start discovery passively whenever the collab tab is visible
            start_discovery()
        # No forced disconnect on tab switch — user may switch just to check logs

    # ------------------------------------------------------------------ MASTER LIST (collab)

    def _refresh_master_list(self, masters: dict):
        for mid, btn in list(self._master_btns.items()):
            if mid not in masters:
                btn.destroy()
                del self._master_btns[mid]

        for mid, info in masters.items():
            label     = f"{info['name']}  ({info['ip']})"
            connected = (mid == active_master_id)
            fg        = ("#2ecc71", "#1a7a44") if connected else ("#3B8ED0", "#1F6AA5")
            text      = f"✓  {label}" if connected else label

            if mid not in self._master_btns:
                btn = ctk.CTkButton(
                    self.master_list_frame,
                    text=text,
                    fg_color=fg,
                    anchor="w",
                    command=lambda m=mid: self._toggle_master(m),
                )
                btn.grid(sticky="ew", padx=4, pady=2)
                self._master_btns[mid] = btn
            else:
                self._master_btns[mid].configure(text=text, fg_color=fg)

    def _toggle_master(self, master_id: str):
        """Connect to a master, or disconnect if it's already the active one."""
        if active_master_id == master_id:
            disconnect_from_master()
            self.collab_status_label.configure(text="Buscando mestres na rede…")
            if master_id in self._master_btns:
                with _disc_lock:
                    info = discovered_masters.get(master_id, {})
                label = f"{info.get('name', master_id)}  ({info.get('ip', '')})"
                self._master_btns[master_id].configure(
                    text=label, fg_color=("#3B8ED0", "#1F6AA5")
                )
        else:
            prev = active_master_id
            if prev:
                disconnect_from_master()
                if prev in self._master_btns:
                    with _disc_lock:
                        pinfo = discovered_masters.get(prev, {})
                    self._master_btns[prev].configure(
                        text=f"{pinfo.get('name', prev)}  ({pinfo.get('ip', '')})",
                        fg_color=("#3B8ED0", "#1F6AA5"),
                    )
            connect_to_master(master_id)
            with _disc_lock:
                info = discovered_masters.get(master_id, {})
            self.collab_status_label.configure(
                text=f"Colaborando com {info.get('name', master_id)}…"
            )
            if master_id in self._master_btns:
                label = f"✓  {info.get('name', master_id)}  ({info.get('ip', '')})"
                self._master_btns[master_id].configure(
                    text=label, fg_color=("#2ecc71", "#1a7a44")
                )

    # ------------------------------------------------------------------ UI LOOP

    def update_ui(self):
        while True:
            try:
                msg = ui_queue.get_nowait()
            except queue.Empty:
                break

            tipo = msg[0]

            if tipo == "renamed":
                self.processed += 1
                self.log.insert("end", f"{msg[1]} → {msg[2]}\n")
                self.log.see("end")

            elif tipo == "notfound":
                self.errors += 1
                self.log.insert("end", f"NO OCR: {msg[1]}\n")
                self.log.see("end")

            elif tipo == "error":
                self.errors += 1
                self.log.insert("end", msg[1] + "\n")
                self.log.see("end")

            elif tipo == "log":
                text = msg[1]
                if text.startswith("[Rede]"):
                    self.net_log.insert("end", text + "\n")
                    self.net_log.see("end")
                else:
                    self.log.insert("end", text + "\n")
                    self.log.see("end")

            elif tipo == "net_log":
                self.net_log.insert("end", msg[1] + "\n")
                self.net_log.see("end")

            elif tipo == "masters_updated":
                self._refresh_master_list(msg[1])

            elif tipo == "net_workers":
                workers = msg[1]
                n = len(workers)
                suffix = ("  —  " + ", ".join(workers)) if workers else ""
                self.log.insert("end", f"[Rede] Colaboradores conectados: {n}{suffix}\n")
                self.log.see("end")

            elif tipo == "collab_status":
                _, status, name = msg
                if status == "connected":
                    self.collab_status_label.configure(text=f"Colaborando com {name}")
                else:
                    self.collab_status_label.configure(text="Desconectado — buscando mestres…")

            # Keep master tab status in sync
            if pasta_atual and master_server:
                collab_count = len(master_server._workers)
                collab_info  = f"  ({collab_count} colaborador(es))" if collab_count else ""
                self.status.configure(text=f"{pasta_atual}{collab_info}")
            elif pasta_atual:
                self.status.configure(text=str(pasta_atual))

        self.queue_label.configure(text=f"Arquivos na fila: {fila.qsize()}")
        self.counter.configure(text=f"Processados: {self.processed} | Erros: {self.errors}")

        self.after(200, self.update_ui)


# ------------------------------------------------------------------ MAIN

def main():
    threads = get_threads()
    for _ in range(threads):
        threading.Thread(target=worker, daemon=True).start()

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()