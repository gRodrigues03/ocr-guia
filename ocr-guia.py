import time
import re
import threading
import queue
from base64 import b64decode
from io import BytesIO
from pathlib import Path

import fitz
from PIL import Image, ImageEnhance
from rapidocr_onnxruntime import RapidOCR

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import pystray
from pystray import MenuItem as item

import tkinter as tk
from tkinter import filedialog

from datetime import datetime

import requests


ICON = """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"""

def load_icon():
    return Image.open(BytesIO(b64decode(ICON)))

def consultar_api(id_, mes):
    url = f"http://148.1.1.11:6969/nguia?id={id_}&mes={mes}"

    r = requests.get(url, timeout=10)
    r.raise_for_status()  # levanta erro se não for 200

    return r.text

# ---------------- CONFIG ----------------


regex_guia = re.compile(r"\d{5,6}", re.IGNORECASE | re.DOTALL)

regex_data = re.compile(r'\\(\d{4})\\(\d{2})\s*-.*?\\(\d{2})\\')

ocr = RapidOCR()

observer = None
pasta_atual = None
mes_selecionado = datetime.now().strftime("%Y-%m")
over_date = datetime.now().strftime("%Y-%m")

fila = queue.Queue()

root = tk.Tk()
root.withdraw()


# ---------------- OCR ----------------

def extrair_guia(pdf_path):

    try:

        doc = fitz.open(pdf_path)
        page = doc[0]

        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        largura, altura = img.size

        if over_date >= '2025-08':
            img = img.crop((largura, 0, largura, int(altura * 0.4)))
        else:
            img = img.crop(((largura*0.5), 0, largura, int(altura * 0.3)))

        img = img.convert("L")
        img = ImageEnhance.Contrast(img).enhance(1.25)

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

    while True:

        pdf = fila.get()

        try:
            processar_pdf(pdf)
        except Exception as e:
            print("Erro processamento:", e)

        fila.task_done()


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
        except:
            over_date = mes_selecionado
        return Path(pasta)

    return None


# ---------------- MENU ----------------

def escolher_mes():

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

    return resultado["valor"]

def alterar_mes(icon, item):

    root.after(0, escolher_mes)

def alterar_pasta(icon, item):

    def selecionar():

        global pasta_atual

        nova = escolher_pasta()

        if not nova:
            return

        pasta_atual = nova

        fila = queue.Queue()

        iniciar_observer()

    root.after(0, selecionar)


def sair(icon, item):

    global observer

    if observer:

        observer.stop()

    icon.stop()

    root.quit()


# ---------------- TRAY ----------------

def iniciar_tray():

    icon = pystray.Icon(
        "RenomeadorPDF",
        load_icon(),
        menu=pystray.Menu(
            item("Alterar pasta", alterar_pasta),
            item("Selecionar mês", alterar_mes),
            item("Sair", sair)
        )
    )

    icon.run()


# ---------------- MAIN ----------------

def main():

    global pasta_atual

    pasta_atual = escolher_pasta()
    escolher_mes()

    if not pasta_atual:
        return

    threading.Thread(target=worker, daemon=True).start()

    iniciar_observer()

    threading.Thread(target=iniciar_tray, daemon=True).start()

    root.mainloop()


if __name__ == "__main__":
    main()