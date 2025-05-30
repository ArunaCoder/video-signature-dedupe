"""
Requirements:
    pip install opencv-python pynput pywin32   # Windows
    pip install opencv-python pynput           # macOS

On macOS you must grant Accessibility permissions
so it can detect ⌘+⇧+P and query Finder.
"""

import os
import sys
import re
import cv2
import threading
import subprocess
from pynput import keyboard
import tkinter as tk
from tkinter import messagebox

# ————— CONFIGURATION —————
GRID = 4                         # 4×4 grid → 16 points
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080
RECORD_FILE = 'processed_videos.txt'
# ————————————————————————

# build the GRID×GRID sample coordinates
POSITIONS = [
    (
        int((i + 1) * VIDEO_WIDTH / (GRID + 1)),
        int((j + 1) * VIDEO_HEIGHT / (GRID + 1))
    )
    for j in range(GRID)
    for i in range(GRID)
]

# regex to parse "2025-05-07_08-22-55 1132 (channel) Title text"
NAME_PATTERN = re.compile(
    r'^(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\s+'
    r'(\d+)\s*'
    r'(?:\([^\)]*\))?\s*'
    r'(.*)$'
)

def parse_name_parts(filename):
    """Extract (date_time, code, text) or (None, None, full-base)."""
    base = os.path.splitext(filename)[0]
    m = NAME_PATTERN.match(base)
    if m:
        date_part = m.group(1)
        code_part = m.group(2)
        text_part = m.group(3).strip()
        return date_part, code_part, text_part
    return None, None, base

def alert(title, msg):
    """Popup a simple messagebox."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning(title, msg)
    root.destroy()

def get_selected_file():
    """Return the single selected file in Explorer/Finder, or None."""
    if sys.platform.startswith('win'):
        try:
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            sel = shell.Windows().Item(0).Document.SelectedItems()
            if sel.Count == 1:
                return sel.Item(0).Path
        except Exception:
            pass
    elif sys.platform == 'darwin':
        script = (
            'tell application "Finder"\n'
            '  set sel to selection as alias list\n'
            '  if (count sel) > 0 then\n'
            '    POSIX path of (item 1 of sel)\n'
            '  end if\n'
            'end tell'
        )
        p = subprocess.run(['osascript','-e',script], capture_output=True, text=True)
        path = p.stdout.strip()
        if path:
            return path
    return None

def process_video(path):
    fn = os.path.basename(path)
    date_cur, code_cur, text_cur = parse_name_parts(fn)

    # load existing records
    records = []
    if os.path.exists(RECORD_FILE):
        with open(RECORD_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                rec_fn, rec_key = line.strip().split('|', 1)
                records.append((rec_fn, rec_key))

    # 1) exact filename?
    for rec_fn, _ in records:
        if rec_fn == fn:
            alert("Duplicate Filename", f"'{fn}' has already been recorded.")
            return

    # 2) date-time prefix?
    if date_cur:
        for rec_fn, _ in records:
            date_rec, _, _ = parse_name_parts(rec_fn)
            if date_rec == date_cur:
                alert("Duplicate Date-Time", f"Date/time '{date_cur}' already used.")
                return

    # 3) numeric code?
    if code_cur:
        for rec_fn, _ in records:
            _, code_rec, _ = parse_name_parts(rec_fn)
            if code_rec == code_cur:
                alert("Duplicate Code", f"Code '{code_cur}' already used.")
                return

    # 4) title text?
    if text_cur:
        for rec_fn, _ in records:
            _, _, text_rec = parse_name_parts(rec_fn)
            if text_rec and text_rec.lower() == text_cur.lower():
                alert("Duplicate Title", f"Title '{text_cur}' already used.")
                return

    # 5) now sample colors
    cap = cv2.VideoCapture(path)
    ok, frame = cap.read()
    cap.release()
    if not ok:
        alert("Read Error", f"Could not read first frame of '{fn}'.")
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

    # 6) color-signature duplicate?
    for _, rec_key in records:
        if rec_key == key:
            alert("Duplicate Colors", "This exact color signature was recorded.")
            return

    # 7) all checks passed → record it
    with open(RECORD_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{fn}|{key}\n")

    # final success alert
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Done", f"'{fn}' processed and recorded.")
    root.destroy()

def on_activate():
    sel = get_selected_file()
    if not sel or not os.path.isfile(sel):
        alert("No Selection", "Please select exactly one video file first.")
        return
    threading.Thread(target=process_video, args=(sel,), daemon=True).start()

if __name__ == '__main__':
    hotkey = '<cmd>+<shift>+p' if sys.platform=='darwin' else '<ctrl>+<shift>+p'
    with keyboard.GlobalHotKeys({hotkey: on_activate}) as listener:
        print(f"Listening for {hotkey} …")
        listener.join()
