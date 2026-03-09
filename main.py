import time
import re
import threading
import queue
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

def consultar_api(id_, mes):
    url = f"http://148.1.1.38:6969/nguia?id={id_}&mes={mes}"

    r = requests.get(url, timeout=10)
    r.raise_for_status()  # levanta erro se não for 200

    return r.text

# ---------------- CONFIG ----------------




regex_guia = re.compile(r"\d{5,6}", re.IGNORECASE | re.DOTALL)

ocr = RapidOCR()

observer = None
pasta_atual = None
mes_selecionado = datetime.now().strftime("%Y-%m")

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

        img = img.crop(((largura*0.5), 0, largura, int(altura * 0.3)))

        img = img.convert("L")
        img = ImageEnhance.Contrast(img).enhance(1.25)

        resultado, _ = ocr(img)

        texto = "\n".join([r[1] for r in resultado]) if resultado else ""

        match = regex_guia.findall(texto)

        if match:
            for m in match:
                text = consultar_api(m, mes_selecionado)
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

        if event.src_path.lower().endswith(".pdf"):

            fila.put(Path(event.src_path))


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
        fila.put(pdf)


# ---------------- PASTA ----------------

def escolher_pasta():

    pasta = filedialog.askdirectory(parent=root)

    if pasta:
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
        Image.new("RGB", (64, 64), (40, 120, 255)),
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
    global mes_selecionado

    pasta_atual = escolher_pasta()
    mes = escolher_mes()
    if mes:
        mes_selecionado = mes


    if not pasta_atual:
        return

    threading.Thread(target=worker, daemon=True).start()

    iniciar_observer()

    threading.Thread(target=iniciar_tray, daemon=True).start()

    root.mainloop()


if __name__ == "__main__":
    main()