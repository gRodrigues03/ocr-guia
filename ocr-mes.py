import time
import re
import threading
import queue
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


def resource_path(filename):
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / filename

# ---------------- API ----------------

def consultar_api(id_, mes):
    url = f"http://148.1.1.11:6969/nguia?id={id_}&mes={mes}"

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
regex_mes = re.compile(r"[\\/](\d{4})[\\/](\d{2})\s*-")

ocr = RapidOCR(use_cls=False)

observer = None
pasta_atual = None

fila = queue.Queue()
ui_queue = queue.Queue()

stop_event = threading.Event()


# ---------------- PATH DATE ----------------

def extrair_data_do_path(path: Path):

    match = regex_mes.search(str(path))

    if not match:
        return None

    ano = match.group(1)
    mes = match.group(2)

    dia = path.parent.name

    if not dia.isdigit():
        return None

    return f"{ano}-{mes}-{int(dia):02d}"


# ---------------- OCR ----------------

def extrair_guia(pdf_path, data_ref):

    try:

        doc = fitz.open(pdf_path)
        page = doc[0]

        rect = page.rect

        if data_ref >= "2025-08-01":
            clip = fitz.Rect(rect.width*0.1, 0, rect.width*0.9, rect.height*0.4)
        else:
            clip = fitz.Rect(rect.width*0.55, 0, rect.width, rect.height * 0.28)

        pix = fitz.Pixmap(doc, page.get_images()[0][0])

        scale = pix.width / rect.width

        x0, y0, x1, y1 = (clip * scale)

        img = np.frombuffer(
            pix.samples,
            np.uint8
        ).reshape(pix.height, pix.width)[int(y0):int(y1), int(x0):int(x1)]

        resultado, _ = ocr(img)

        texto = "\n".join([r[1] for r in resultado]) if resultado else ""

        match = regex_guia.findall(texto)

        if match:
            for m in match:
                text = consultar_api(m, data_ref)
                if text and len(text):
                    return text.split(',')[0], texto

        return None, texto

    except Exception as e:

        ui_queue.put(("log", f"Erro em {pdf_path.name}: {e}"))

    return None, None


# ---------------- PROCESSAMENTO ----------------

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


def processar_pdf(pdf):

    if not esperar_arquivo_finalizar(pdf):
        return

    data_ref = extrair_data_do_path(pdf)

    if not data_ref:
        ui_queue.put(("error", f"Data não encontrada: {pdf.name}"))
        return

    guia, texto = extrair_guia(pdf, data_ref)

    if guia:

        novo_nome = pdf.with_name(f"{guia}.pdf")

        contador = 1

        while novo_nome.exists():
            novo_nome = pdf.with_name(f"{guia} ({contador}).pdf")
            contador += 1

        pdf.rename(novo_nome)

        ui_queue.put(("renamed", pdf.name, novo_nome.name))

    else:

        novo_nome = pdf.with_name(f"NO_OCR {pdf.name}")

        contador = 1

        while novo_nome.exists():
            novo_nome = pdf.with_name(f"NO_OCR {contador} {pdf.name}")
            contador += 1

        pdf.rename(novo_nome)

        ui_queue.put(("notfound", novo_nome.name))


def worker():

    while not stop_event.is_set():

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

    handler = Handler()

    observer = Observer()
    observer.schedule(handler, str(pasta_atual), recursive=True)
    observer.start()

    ui_queue.put(("log", f"Observando: {pasta_atual}"))

    for pdf in pasta_atual.glob("*/*.pdf"):

        if (
            pdf.name
            and not pdf.name[0].isdigit()
            and not pdf.name.startswith("NO_OCR ")
        ):
            fila.put(pdf)


# ---------------- UI ----------------

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("OCR Renomeador")
        self.geometry("600x420")

        try:
            ico = resource_path("exeicon.ico")
            png = resource_path("trayicon.png")

            if ico.exists():
                self.iconbitmap(ico)
            elif png.exists():
                from PIL import ImageTk, Image
                img = ImageTk.PhotoImage(Image.open(png))
                self.iconphoto(True, img)
                self._icon = img  # keep reference
        except Exception:
            pass

        self.processed = 0
        self.errors = 0

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self.btn = ctk.CTkButton(self, text="Selecionar pasta", command=self.selecionar_pasta)
        self.btn.grid(row=0, column=0, padx=20, pady=(20,10), sticky="w")

        self.status = ctk.CTkLabel(self, text="Nenhuma pasta selecionada, selecione uma pasta para começar" if pasta_atual is None else pasta_atual)
        self.status.grid(row=1, column=0, padx=20, sticky="w")

        self.queue_label = ctk.CTkLabel(self, text="Fila: 0")
        self.queue_label.grid(row=2, column=0, padx=20, sticky="w")

        self.counter = ctk.CTkLabel(self, text="Processados: 0 | Erros: 0")
        self.counter.grid(row=3, column=0, padx=20, sticky="w")

        self.log = ctk.CTkTextbox(self)
        self.log.grid(row=5, column=0, padx=20, pady=(0,20), sticky="nsew")

        self.after(100, self.update_ui)

    def selecionar_pasta(self):

        global pasta_atual
        global fila

        pasta = filedialog.askdirectory()

        if not pasta:
            return

        pasta_atual = Path(pasta)

        fila = queue.Queue()

        iniciar_observer()

    def update_ui(self):

        while True:

            try:
                msg = ui_queue.get_nowait()
            except queue.Empty:
                break

            tipo = msg[0]

            self.status.configure(text="Nenhuma pasta selecionada, selecione uma pasta para começar" if pasta_atual is None else pasta_atual)

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

            elif tipo == "log":
                self.log.insert("end", msg[1] + "\n")

        self.queue_label.configure(text=f"Arquivos na fila: {fila.qsize()}")
        self.counter.configure(text=f"Processados: {self.processed} | Erros: {self.errors}")

        self.after(200, self.update_ui)


# ---------------- MAIN ----------------

def main():

    threads = get_threads()

    for _ in range(threads):
        threading.Thread(target=worker, daemon=True).start()

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()