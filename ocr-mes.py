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

def resource_path(filename):
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / filename


# ---------------- API ----------------

def consultar_api(id_, empresa, mes):
    url = f"http://148.1.1.11:6969/nguia?id={id_}&mes={mes}&empresa={empresas_index.get(empresa, None)}"
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

        img = np.frombuffer(pix.samples, np.uint8).reshape(
            pix.height, pix.width
        )[int(y0):int(y1), int(x0):int(x1)]

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

def esperar_arquivo_finalizar(path):
    tamanho = -1
    while True:
        try:
            novo = path.stat().st_size
        except FileNotFoundError:
            return False
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
                jobs.append(str(pdf))
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

        # ---- Tab view ----
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.tabs.add("Local")
        self.tabs.add("Rede")

        self._build_local_tab()
        self._build_net_tab()

        self.after(200, self.update_ui)

    # ------------------------------------------------------------------ LOCAL TAB

    def _build_local_tab(self):
        tab = self.tabs.tab("Local")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(4, weight=1)

        self.btn = ctk.CTkButton(tab, text="Selecionar pasta", command=self.selecionar_pasta)
        self.btn.grid(row=0, column=0, padx=0, pady=(10, 6), sticky="w")

        self.status = ctk.CTkLabel(
            tab,
            text="Nenhuma pasta selecionada" if pasta_atual is None else str(pasta_atual)
        )
        self.status.grid(row=1, column=0, padx=0, sticky="w")

        self.queue_label = ctk.CTkLabel(tab, text="Fila: 0")
        self.queue_label.grid(row=2, column=0, padx=0, sticky="w")

        self.counter = ctk.CTkLabel(tab, text="Processados: 0 | Erros: 0")
        self.counter.grid(row=3, column=0, padx=0, sticky="w")

        self.log = ctk.CTkTextbox(tab)
        self.log.grid(row=4, column=0, padx=0, pady=(6, 0), sticky="nsew")

    def selecionar_pasta(self):
        global pasta_atual, fila
        pasta = filedialog.askdirectory()
        if not pasta:
            return
        pasta_atual = Path(pasta)
        fila = queue.Queue()
        iniciar_observer()

    # ------------------------------------------------------------------ NETWORK TAB

    def _build_net_tab(self):
        tab = self.tabs.tab("Rede")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(3, weight=1)

        # Role selector
        role_frame = ctk.CTkFrame(tab, fg_color="transparent")
        role_frame.grid(row=0, column=0, sticky="ew", pady=(10, 6))

        ctk.CTkLabel(role_frame, text="Papel neste nó:").pack(side="left", padx=(0, 10))

        self.role_var = ctk.StringVar(value="none")

        self.btn_master = ctk.CTkButton(
            role_frame, text="Ser Mestre",
            width=120, command=self._activate_master
        )
        self.btn_master.pack(side="left", padx=4)

        self.btn_worker = ctk.CTkButton(
            role_frame, text="Ser Colaborador",
            width=140, fg_color="gray", command=self._activate_worker
        )
        self.btn_worker.pack(side="left", padx=4)

        self.btn_net_stop = ctk.CTkButton(
            role_frame, text="Parar",
            width=80, fg_color="#c0392b", hover_color="#922b21",
            command=self._deactivate_net
        )
        self.btn_net_stop.pack(side="left", padx=4)
        self.btn_net_stop.configure(state="disabled")

        # Status line
        self.net_status_label = ctk.CTkLabel(tab, text="Inativo")
        self.net_status_label.grid(row=1, column=0, sticky="w")

        # Middle area: master shows worker list; worker shows master list
        self.net_mid_frame = ctk.CTkFrame(tab)
        self.net_mid_frame.grid(row=2, column=0, sticky="ew", pady=4)
        self.net_mid_frame.grid_columnconfigure(0, weight=1)

        self.net_mid_label = ctk.CTkLabel(self.net_mid_frame, text="")
        self.net_mid_label.grid(row=0, column=0, sticky="w", padx=8, pady=4)

        # Master list (for worker mode)
        self.master_list_frame = ctk.CTkScrollableFrame(tab, height=100)
        self.master_list_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 4))
        self.master_list_frame.grid_columnconfigure(0, weight=1)
        self._master_btns: dict[str, ctk.CTkButton] = {}

        # Network log
        self.net_log = ctk.CTkTextbox(tab, height=120)
        self.net_log.grid(row=4, column=0, sticky="ew", pady=(0, 4))

    def _activate_master(self):
        global master_server
        if master_server:
            return
        master_server = MasterServer()
        master_server.start()
        self.role_var.set("master")
        self.net_status_label.configure(text=f"Mestre ativo — {NODE_NAME}")
        self.btn_master.configure(state="disabled")
        self.btn_worker.configure(state="disabled")
        self.btn_net_stop.configure(state="normal")
        self.net_mid_label.configure(text="Colaboradores conectados:")

    def _activate_worker(self):
        self.role_var.set("worker")
        start_discovery()
        self.net_status_label.configure(text="Buscando mestres na rede…")
        self.btn_master.configure(state="disabled")
        self.btn_worker.configure(state="disabled")
        self.btn_net_stop.configure(state="normal")
        self.net_mid_label.configure(text="Mestres disponíveis:")

    def _deactivate_net(self):
        global master_server
        if master_server:
            master_server.stop()
            master_server = None
        disconnect_from_master()
        stop_event.set()
        self.role_var.set("none")
        self.net_status_label.configure(text="Inativo")
        self.btn_master.configure(state="normal")
        self.btn_worker.configure(state="normal")
        self.btn_net_stop.configure(state="disabled")

    def _refresh_master_list(self, masters: dict):
        # Remove buttons for masters that disappeared
        for mid, btn in list(self._master_btns.items()):
            if mid not in masters:
                btn.destroy()
                del self._master_btns[mid]

        # Add / update
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
            # Disconnect toggle
            disconnect_from_master()
            self.net_status_label.configure(text="Buscando mestres na rede…")
            # Reset button appearance
            if master_id in self._master_btns:
                with _disc_lock:
                    info = discovered_masters.get(master_id, {})
                label = f"{info.get('name', master_id)}  ({info.get('ip', '')})"
                self._master_btns[master_id].configure(
                    text=label, fg_color=("#3B8ED0", "#1F6AA5")
                )
        else:
            # Disconnect previous if any
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
            self.net_status_label.configure(
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

            # local tab updates
            self.status.configure(
                text="Nenhuma pasta selecionada" if pasta_atual is None else str(pasta_atual)
            )

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
                # route net messages to net log
                if text.startswith("[Rede]"):
                    self.net_log.insert("end", text + "\n")
                    self.net_log.see("end")
                else:
                    self.log.insert("end", text + "\n")
                    self.log.see("end")

            elif tipo == "net_log":
                # results processed by this node as collaborator
                self.net_log.insert("end", msg[1] + "\n")
                self.net_log.see("end")

            # network-specific
            elif tipo == "masters_updated":
                if self.role_var.get() == "worker":
                    self._refresh_master_list(msg[1])

            elif tipo == "net_workers":
                if self.role_var.get() == "master":
                    workers = msg[1]
                    self.net_mid_label.configure(
                        text=f"Colaboradores conectados: {len(workers)}"
                        + (("  —  " + ", ".join(workers)) if workers else "")
                    )

            elif tipo == "collab_status":
                _, status, name = msg
                if status == "connected":
                    self.net_status_label.configure(text=f"Colaborando com {name}")
                else:
                    if self.role_var.get() == "worker":
                        self.net_status_label.configure(text="Desconectado — aguardando mestre…")

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