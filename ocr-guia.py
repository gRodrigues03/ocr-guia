import time
import re
import threading
import queue
from pathlib import Path

import fitz
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import pystray
from pystray import MenuItem as item

import tkinter as tk
from tkinter import filedialog

from datetime import datetime

import requests


trayicon = None


def consultar_api(id_, mes):
    url = f"http://148.1.1.11:6969/nguia?id={id_}&mes={mes}"

    r = requests.get(url, timeout=5)
    r.raise_for_status()  # levanta erro se não for 200

    return r.text

# ---------------- CONFIG ----------------


regex_guia = re.compile(r"\d{5,6}", re.IGNORECASE | re.DOTALL)

regex_data = re.compile(r'[\\/](\d{4})[\\/](\d{2})\s*-.*?[\\/](\d{2})')

ocr = RapidOCR(use_cls=False)

observer = None
pasta_atual = None
mes_selecionado = datetime.now().strftime("%Y-%m")
over_date = datetime.now().strftime("%Y-%m")

fila = queue.Queue()

stop_event = threading.Event()

root = tk.Tk()
root.withdraw()


# ---------------- OCR ----------------

def extrair_guia(pdf_path):

    try:

        doc = fitz.open(pdf_path)
        page = doc[0]

        rect = page.rect

        if (len(over_date) == 7 and over_date >= '2025-08') or (len(over_date) > 7 and over_date >= '2025-08-01'):
            clip = fitz.Rect(0, 0, rect.width, rect.height*0.4)
        else:
            clip = fitz.Rect(rect.width*0.5, 0, rect.width, rect.height * 0.3)

        pix = page.get_pixmap(matrix=fitz.Matrix(200 / 72, 200 / 72), clip=clip, colorspace=fitz.csGRAY)

        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)

        resultado, _ = ocr(img)

        texto = "\n".join([r[1] for r in resultado]) if resultado else ""

        match = regex_guia.findall(texto)

        if match:
            for m in match:
                text = consultar_api(m, over_date)
                if text and len(text):
                    return text.split(',')[0], texto

        return None, texto

    except Exception as e:

        print(f"Erro em {pdf_path.name}: {e}")

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

    guia, texto = extrair_guia(pdf)

    if guia:

        novo_nome = pdf.with_name(f"{guia}.pdf")

        contador = 1

        while novo_nome.exists():

            novo_nome = pdf.with_name(f"{guia} ({contador}).pdf")
            contador += 1

        pdf.rename(novo_nome)

        print(f"{pdf.name} → {novo_nome.name}")

    else:

        print(f"Guia não encontrada: {pdf.name}")
        print(texto)


def worker():
    while not stop_event.is_set():
        try:
            pdf = fila.get(timeout=0.5)
        except queue.Empty:
            continue

        processar_pdf(pdf)
        fila.task_done()
    print('worker thread stopped')

# ---------------- WATCHDOG ----------------

class Handler(FileSystemEventHandler):

    def on_created(self, event):

        if event.is_directory:
            return

        path = Path(event.src_path)

        if path.suffix.lower() == ".pdf" and not path.name[0].isdigit():

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
    observer.schedule(handler, str(pasta_atual), recursive=False)
    observer.start()

    print("Observando:", pasta_atual)

    for pdf in pasta_atual.glob("*.pdf"):

        if not pdf.name[0].isdigit():
            fila.put(pdf)


# ---------------- PASTA ----------------

def escolher_pasta():
    global over_date

    pasta = filedialog.askdirectory(parent=root)

    if pasta:
        try:
            tmp_path = str(pasta)
            for i in ('GLORIA', 'PONTE', 'GARDEL'):
                if i in tmp_path:
                    date = regex_data.search(tmp_path)
                    if date:
                        over_date = f"{date.group(1)}-{date.group(2)}-{date.group(3)}"
                    else:
                        over_date = mes_selecionado
                    break
        except:
            over_date = mes_selecionado
        return Path(pasta)

    return None


# ---------------- MENU ----------------

def escolher_mes():
    global mes_selecionado
    global over_date

    janela = tk.Toplevel(root)
    janela.title("Selecionar mês")
    janela.resizable(False, False)

    ano_var = tk.IntVar(value=int(mes_selecionado[:4]))
    mes_var = tk.IntVar(value=int(mes_selecionado[5:]))

    resultado = {"valor": None}

    tk.Label(janela, text="Ano").grid(row=0, column=0, padx=10, pady=5)
    tk.Label(janela, text="Mês").grid(row=0, column=1, padx=10, pady=5)

    tk.Spinbox(janela, from_=2000, to=2100, textvariable=ano_var, width=8)\
        .grid(row=1, column=0, padx=10)

    tk.Spinbox(janela, from_=1, to=12, textvariable=mes_var, width=5)\
        .grid(row=1, column=1, padx=10)

    def confirmar():
        resultado["valor"] = f"{ano_var.get()}-{mes_var.get():02d}"
        janela.destroy()

    tk.Button(janela, text="OK", command=confirmar)\
        .grid(row=2, column=0, columnspan=2, pady=10)

    janela.grab_set()
    janela.wait_window()

    if len(over_date) == 7:
        over_date = resultado["valor"]
    mes_selecionado = resultado["valor"]

def alterar_mes(icon, item):

    root.after(0, escolher_mes)

def alterar_pasta(icon, item):

    def selecionar():

        global pasta_atual
        global fila

        nova = escolher_pasta()

        if not nova:
            return

        pasta_atual = nova

        fila = queue.Queue()

        iniciar_observer()

    root.after(0, selecionar)


def sair(icon, item):
    stop_event.set()

    observer.stop()
    observer.join()

    trayicon.stop()

    root.destroy()


# ---------------- TRAY ----------------

def iniciar_tray():
    global trayicon
    trayicon = pystray.Icon(
        "RenomeadorPDF",
        Image.open("trayicon.png"),
        menu=pystray.Menu(
            item("Alterar pasta", alterar_pasta),
            item("Selecionar mês", alterar_mes),
            item("Sair", sair)
        )
    )

    trayicon.run()
    print('tray stopped')


# ---------------- MAIN ----------------

def main():

    global pasta_atual

    escolher_mes()
    pasta_atual = escolher_pasta()

    if not pasta_atual:
        return

    for _ in range(2):
        threading.Thread(target=worker, daemon=True).start()

    iniciar_observer()

    threading.Thread(target=iniciar_tray, daemon=True).start()

    root.mainloop()


if __name__ == "__main__":
    main()