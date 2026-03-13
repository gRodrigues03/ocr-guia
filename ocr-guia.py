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

import tkinter as tk
from tkinter import filedialog, ttk

import requests
import sys


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

BASE_DIR = Path(__file__).resolve().parent

regex_guia = re.compile(r"\d{5,6}", re.IGNORECASE | re.DOTALL)
regex_mes = re.compile(r"[\\/](\d{4})[\\/](\d{2})\s*-")

ocr = RapidOCR(use_cls=False)

observer = None
pasta_atual = None

fila = queue.Queue()
ui_queue = queue.Queue()

stop_event = threading.Event()


# ---------------- UI ----------------

root = tk.Tk()
root.title("Renomeador de PDFs")
root.geometry("600x420")

frame = tk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

btn_pasta = tk.Button(frame, text="Selecionar pasta")
btn_pasta.pack(anchor="w")

status = tk.Label(frame, text="Aguardando pasta...")
status.pack(anchor="w", pady=(10, 0))

fila_label = tk.Label(frame, text="Arquivos na fila: 0")
fila_label.pack(anchor="w")

contador_label = tk.Label(frame, text="Processados: 0 | Erros: 0")
contador_label.pack(anchor="w")

progress = ttk.Progressbar(frame, mode="indeterminate")
progress.pack(fill="x", pady=10)

log = tk.Text(frame, height=15)
log.pack(fill="both", expand=True)

processed_count = 0
error_count = 0


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

    ui_queue.put(("processing", pdf.name))

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

    if not pasta_atual:
        return

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


# ---------------- UI UPDATE LOOP ----------------

def atualizar_ui():

    global processed_count
    global error_count

    while True:

        try:
            msg = ui_queue.get_nowait()
        except queue.Empty:
            break

        tipo = msg[0]

        if tipo == "processing":

            status.config(text=f"Processando: {msg[1]}")
            progress.start()

        elif tipo == "renamed":

            progress.stop()

            processed_count += 1

            log.insert("end", f"{msg[1]} → {msg[2]}\n")
            log.see("end")

        elif tipo == "notfound":

            progress.stop()

            error_count += 1

            log.insert("end", f"Guia não encontrada: {msg[1]}\n")
            log.see("end")

        elif tipo == "error":

            error_count += 1

            log.insert("end", msg[1] + "\n")
            log.see("end")

        elif tipo == "log":

            log.insert("end", msg[1] + "\n")
            log.see("end")

    fila_label.config(text=f"Arquivos na fila: {fila.qsize()}")
    contador_label.config(text=f"Processados: {processed_count} | Erros: {error_count}")

    root.after(100, atualizar_ui)


# ---------------- UI ACTIONS ----------------

def selecionar_pasta():

    global pasta_atual
    global fila

    pasta = filedialog.askdirectory(parent=root)

    if not pasta:
        return

    pasta_atual = Path(pasta)

    fila = queue.Queue()

    iniciar_observer()


btn_pasta.config(command=selecionar_pasta)


# ---------------- MAIN ----------------

def main():

    threads = get_threads()

    for _ in range(threads):
        threading.Thread(target=worker, daemon=True).start()

    root.after(200, atualizar_ui)

    root.mainloop()


if __name__ == "__main__":
    main()