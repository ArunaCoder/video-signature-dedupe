#!/usr/bin/env python3
"""
Cross-platform video deduplication via 16-point frame color signature,
filename parsing & global hotkey (Ctrl+Shift+P / ⌘+⇧+P).
"""

import sys
import os
import re
import threading
import subprocess
import cv2
import pythoncom            # PyWin32 COM initializer
import win32gui             # PyWin32 GUI utilities
import win32com.client      # PyWin32 COM client
from pynput import keyboard
import tkinter as tk
from tkinter import messagebox

# ————— CONFIGURATION —————
GRID = 4                    # 4×4 grid → 16 sample points
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080
RECORD_FILE = 'processed_videos.txt'
# regex to parse "2025-05-07_08-22-55 1132 (channel) Title text"
NAME_PATTERN = re.compile(
    r'^(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\s+'
    r'(\d+)\s*'
    r'(?:\([^\)]*\))?\s*'
    r'(.*)$'
)
# build the GRID×GRID sample coordinates
POSITIONS = [
    (
        int((i + 1) * VIDEO_WIDTH / (GRID + 1)),
        int((j + 1) * VIDEO_HEIGHT / (GRID + 1))
    )
    for j in range(GRID)
    for i in range(GRID)
]
# ————————————————————————

def alert_info(title, msg):
    """Caixa de mensagem informativa."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, msg)
    root.destroy()

def alert_warn(title, msg):
    """Caixa de mensagem de alerta."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning(title, msg)
    root.destroy()

def parse_name_parts(filename):
    """Extrai (date_time, code, text) de acordo com o padrão."""
    base = os.path.splitext(filename)[0]
    m = NAME_PATTERN.match(base)
    if m:
        date_part = m.group(1)
        code_part = m.group(2)
        text_part = m.group(3).strip()
        return date_part, code_part, text_part
    return None, None, base

def get_selected_file():
    """
    Retorna o único arquivo selecionado no Explorer (Windows) ou Finder (macOS).
    """
    # --- Windows ---
    if sys.platform.startswith('win'):
        # inicializa o COM nesta thread
        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()
            hwnd_active = win32gui.GetForegroundWindow()

            for i in range(windows.Count):
                window = windows.Item(i)
                try:
                    if window.HWND == hwnd_active:
                        sel = window.Document.SelectedItems()
                        if sel.Count == 1:
                            return sel.Item(0).Path
                except Exception:
                    continue
        finally:
            pythoncom.CoUninitialize()
        return None

    # --- macOS ---
    elif sys.platform == 'darwin':
        applescript = '''
        tell application "Finder"
          set sel to selection as alias list
          if (count sel) > 0 then
            POSIX path of (item 1 of sel)
          end if
        end tell
        '''
        proc = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True, text=True
        )
        path = proc.stdout.strip()
        return path if path else None

    # --- Outros OS ---
    else:
        return None

def process_video(path):
    """Faz todos os checks (nome + cor) e registra o vídeo se for novo."""
    fn = os.path.basename(path)
    date_cur, code_cur, text_cur = parse_name_parts(fn)

    # carrega registros existentes
    records = []
    if os.path.exists(RECORD_FILE):
        with open(RECORD_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                rec_fn, rec_key = line.strip().split('|', 1)
                records.append((rec_fn, rec_key))

    # 1) filename exato?
    for rec_fn, _ in records:
        if rec_fn == fn:
            alert_warn("Duplicate Filename", f"'{fn}' já está registrado.")
            return

    # 2) date-time?
    if date_cur:
        for rec_fn, _ in records:
            date_rec, _, _ = parse_name_parts(rec_fn)
            if date_rec == date_cur:
                alert_warn("Duplicate Date-Time", f"Date/time '{date_cur}' já usado.")
                return

    # 3) numeric code?
    if code_cur:
        for rec_fn, _ in records:
            _, code_rec, _ = parse_name_parts(rec_fn)
            if code_rec == code_cur:
                alert_warn("Duplicate Code", f"Código '{code_cur}' já usado.")
                return

    # 4) title text?
    if text_cur:
        for rec_fn, _ in records:
            _, _, text_rec = parse_name_parts(rec_fn)
            if text_rec and text_rec.lower() == text_cur.lower():
                alert_warn("Duplicate Title", f"Título '{text_cur}' já usado.")
                return

    # 5) captura e amostra cores
    cap = cv2.VideoCapture(path)
    ok, frame = cap.read()
    cap.release()
    if not ok:
        alert_warn("Read Error", f"Não foi possível ler '{fn}'.")
        return

    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    h, w, _ = frame_rgb.shape
    colors = []
    for x, y in POSITIONS:
        if 0 <= x < w and 0 <= y < h:
            r, g, b = frame_rgb[y, x]
            colors.append(f"#{r:02X}{g:02X}{b:02X}")
        else:
            colors.append("#000000")
    key = ",".join(colors)

    # 6) cor duplicada?
    for _, rec_key in records:
        if rec_key == key:
            alert_warn("Duplicate Colors", f"Assinatura de cores já registrada no vídeo '{rec_fn}'.")
            return

    # 7) tudo ok → grava registro
    with open(RECORD_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{fn}|{key}\n")

    alert_info("Done", f"'{fn}' processado e registrado.")

def on_activate():
    """Callback do hotkey."""
    sel = get_selected_file()
    if not sel or not os.path.isfile(sel):
        alert_warn("No Selection", "Selecione exatamente um arquivo de vídeo.")
        return
    # executa processamento em thread separada
    threading.Thread(target=process_video, args=(sel,), daemon=True).start()

if __name__ == '__main__':
    # define hotkey cross-platform
    hotkey = '<cmd>+<shift>+p' if sys.platform=='darwin' else '<ctrl>+<shift>+p'
    print(f"Listening for {hotkey} … (select a video in Explorer/Finder and press it)")
    with keyboard.GlobalHotKeys({hotkey: on_activate}) as listener:
        listener.join()
