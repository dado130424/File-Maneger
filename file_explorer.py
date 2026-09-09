#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Esplora File — un file explorer in stile Windows Explorer.
Realizzato con SOLA libreria standard (tkinter): nessuna dipendenza esterna.

Esegui:        python file_explorer.py
                 oppure doppio clic su build_exe.bat per creare l'.exe

Funzionalità:
  - Navigazione con cronologia (Indietro / Avanti), cartella superiore, aggiorna
  - Schede multiple come nel browser (Ctrl+T nuova, Ctrl+W chiudi, Ctrl+Tab cambia)
  - Barra degli indirizzi + barra laterale (Accesso rapido e Unità disco)
  - Ordinamento per colonna (Nome / Tipo / Dimensioni / Data)
  - Ricerca RICORSIVA: il filtro cerca anche in tutte le sottocartelle
    (in background, con colonna "Percorso" nei risultati)
  - Apri con doppio clic / Invio, menu contestuale (tasto destro)
  - Anteprima laterale: immagini (con zoom e risoluzione regolabile),
    testo e codice (decine di formati), PDF, documenti Office e
    contenuto di archivi ZIP/TAR
  - Copia, Taglia, Incolla, Rinomina, Elimina (nel Cestino su Windows),
    Proprietà, Copia percorso, Apri terminale qui
  - Menu "Crea" (tasto destro): cartella, file di testo, Word, PowerPoint, Python
  - Drag & drop interno: trascina file/cartelle su cartelle, sulla barra
    laterale o su un'altra scheda (Ctrl per copiare invece di spostare)
  - Filtro di ricerca istantaneo, file nascosti opzionali, scorciatoie da tastiera
"""

import base64
import ctypes
import csv
import datetime
import difflib
import fnmatch
import hashlib
import json
import math
import os
import queue
import re
import shutil
import string
import struct
import subprocess
import sys
import tarfile
import tempfile
import threading
import zipfile
import zlib
from html import unescape as _xml_unescape
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, font as tkfont

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"

# Etichetta della voce "mostra nel file manager di sistema" (varia per OS)
if IS_WINDOWS:
    REVEAL_LABEL = "Apri in Esplora Risorse"
elif IS_MACOS:
    REVEAL_LABEL = "Mostra nel Finder"
else:
    REVEAL_LABEL = "Mostra nel file manager"

MAX_SEARCH_RESULTS = 2000

PREVIEW_TEXT_LIMIT = 128 * 1024          # anteprima testo: primi ~128 KB
DIFF_READ_LIMIT = 2 * 1024 * 1024         # max byte letti per file nel confronto
CONTENT_SEARCH_LIMIT = 4 * 1024 * 1024    # max byte letti per file nella ricerca nel contenuto
PREVIEW_IMAGE_MAX_PIXELS = 60_000_000    # limite dimensione anteprima immagini
MAX_PDF_STREAM = 8 * 1024 * 1024         # max byte decompressi per stream PDF (anti bomba)

# --------------------------------------------------------------------------
# Icone (emoji) e descrizioni dei tipi
# --------------------------------------------------------------------------
ICON_FOLDER = "\U0001F4C1"      # 📁
ICON_DRIVE  = "\U0001F4BD"      # 💽
ICON_PC     = "\U0001F5A5"      # 🖥️
ICON_IMG    = "\U0001F5BC"      # 🖼️
ICON_AUDIO  = "\U0001F3B5"      # 🎵
ICON_VIDEO  = "\U0001F3AC"      # 🎬
ICON_ARCH   = "\U0001F4E6"      # 📦
ICON_EXE    = "\u2699"          # ⚙️
ICON_DOC    = "\U0001F4C4"      # 📄
ICON_CODE   = "\U0001F4BB"      # 💻
ICON_TEXT   = "\U0001F4DD"      # 📝
ICON_MISC   = "\U0001F4C4"      # 📄

IMAGE_EXTS   = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg",
                ".ico", ".heic", ".tiff", ".tif", ".raw", ".psd"}
AUDIO_EXTS   = {".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma", ".opus"}
VIDEO_EXTS   = {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm", ".m4v", ".mpg", ".mpeg"}
ARCHIVE_EXTS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".iso", ".cab", ".tgz"}
EXE_EXTS     = {".exe", ".msi", ".bat", ".cmd", ".com", ".ps1", ".apk", ".jar"}
CODE_EXTS    = {".py", ".js", ".ts", ".html", ".css", ".json", ".xml", ".yml",
                ".yaml", ".c", ".cpp", ".h", ".java", ".cs", ".go", ".rs", ".sql",
                ".php", ".rb", ".lua", ".pl", ".perl", ".swift", ".kt", ".kts",
                ".dart", ".r", ".m", ".scala", ".groovy", ".sh", ".vb", ".vue",
                ".jsx", ".tsx", ".less", ".scss", ".sass", ".styl",
                ".dockerfile", ".makefile", ".cmake", ".gradle"}
TEXT_EXTS    = {".txt", ".md", ".log", ".ini", ".cfg", ".csv", ".rtf", ".srt",
                ".vtt", ".eml", ".vcf", ".ics", ".nfo", ".text", ".rst",
                ".asciidoc", ".tex", ".bib", ".properties", ".toml", ".env",
                ".editorconfig", ".gitignore", ".reg", ".inf"}
DOC_EXTS     = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp"}

TYPE_LABELS = {
    ".pdf": "Documento PDF", ".doc": "Documento Word", ".docx": "Documento Word",
    ".xls": "Foglio Excel", ".xlsx": "Foglio Excel", ".ppt": "Presentazione",
    ".pptx": "Presentazione", ".odt": "Documento", ".ods": "Foglio di calcolo",
    ".odp": "Presentazione", ".png": "Immagine PNG", ".jpg": "Immagine JPEG",
    ".jpeg": "Immagine JPEG", ".gif": "Immagine GIF", ".bmp": "Immagine BMP",
    ".webp": "Immagine WebP", ".svg": "Immagine SVG", ".ico": "Icona",
    ".mp3": "Audio MP3", ".wav": "Audio WAV", ".flac": "Audio FLAC",
    ".aac": "Audio AAC", ".ogg": "Audio OGG", ".m4a": "Audio M4A",
    ".mp4": "Video MP4", ".mkv": "Video MKV", ".avi": "Video AVI",
    ".mov": "Video MOV", ".wmv": "Video WMV", ".webm": "Video WebM",
    ".zip": "Archivio ZIP", ".rar": "Archivio RAR", ".7z": "Archivio 7-Zip",
    ".tar": "Archivio TAR", ".gz": "Archivio GZIP", ".iso": "Immagine disco",
    ".exe": "Applicazione", ".msi": "Programma di installazione",
    ".bat": "File batch", ".py": "Script Python", ".js": "Script JavaScript",
    ".ts": "Script TypeScript", ".html": "Pagina HTML", ".css": "Foglio di stile",
    ".json": "File JSON", ".xml": "File XML", ".txt": "Documento di testo",
    ".md": "File Markdown", ".log": "File di log", ".csv": "File CSV",
    ".sql": "Script SQL", ".java": "File Java", ".c": "File C", ".cpp": "File C++",
    ".cs": "File C#", ".go": "File Go", ".rs": "File Rust",
}

ICON_EXTS = {
    ".png": ICON_IMG, ".jpg": ICON_IMG, ".jpeg": ICON_IMG, ".gif": ICON_IMG,
    ".bmp": ICON_IMG, ".webp": ICON_IMG, ".svg": ICON_IMG, ".ico": ICON_IMG,
    ".heic": ICON_IMG, ".tiff": ICON_IMG,
    ".mp3": ICON_AUDIO, ".wav": ICON_AUDIO, ".flac": ICON_AUDIO, ".aac": ICON_AUDIO,
    ".ogg": ICON_AUDIO, ".m4a": ICON_AUDIO, ".wma": ICON_AUDIO,
    ".mp4": ICON_VIDEO, ".mkv": ICON_VIDEO, ".avi": ICON_VIDEO, ".mov": ICON_VIDEO,
    ".wmv": ICON_VIDEO, ".flv": ICON_VIDEO, ".webm": ICON_VIDEO, ".m4v": ICON_VIDEO,
    ".zip": ICON_ARCH, ".rar": ICON_ARCH, ".7z": ICON_ARCH, ".tar": ICON_ARCH,
    ".gz": ICON_ARCH, ".bz2": ICON_ARCH, ".xz": ICON_ARCH, ".iso": ICON_ARCH,
    ".exe": ICON_EXE, ".msi": ICON_EXE, ".bat": ICON_EXE, ".cmd": ICON_EXE,
    ".com": ICON_EXE, ".ps1": ICON_EXE, ".apk": ICON_EXE, ".jar": ICON_EXE,
    ".py": ICON_CODE, ".js": ICON_CODE, ".ts": ICON_CODE, ".html": ICON_CODE,
    ".css": ICON_CODE, ".json": ICON_CODE, ".xml": ICON_CODE, ".yml": ICON_CODE,
    ".yaml": ICON_CODE, ".c": ICON_CODE, ".cpp": ICON_CODE, ".h": ICON_CODE,
    ".java": ICON_CODE, ".cs": ICON_CODE, ".go": ICON_CODE, ".rs": ICON_CODE,
    ".sql": ICON_CODE,
    ".txt": ICON_TEXT, ".md": ICON_TEXT, ".log": ICON_TEXT, ".ini": ICON_TEXT,
    ".cfg": ICON_TEXT, ".rtf": ICON_TEXT,
    ".pdf": ICON_DOC, ".doc": ICON_DOC, ".docx": ICON_DOC, ".xls": ICON_DOC,
    ".xlsx": ICON_DOC, ".ppt": ICON_DOC, ".pptx": ICON_DOC,
}

# Estensioni con anteprima di testo e tipi binari noti (nessuna anteprima)
PREVIEW_TEXT_EXTS = TEXT_EXTS | CODE_EXTS
KNOWN_BINARY_EXTS = ((IMAGE_EXTS | AUDIO_EXTS | VIDEO_EXTS | ARCHIVE_EXTS
                      | EXE_EXTS | DOC_EXTS) - {".png", ".gif", ".pdf"})
# Formati testuali "nascosti" tra i binari (EXE/immagini): il dispatch
# li mostra come testo PRIMA del fallback dei binari
TEXT_LIKE_EXTS = {".bat", ".cmd", ".ps1", ".svg"}

# Nomi leggibili dei formati riconosciuti da mutagen (nome della classe)
MEDIA_FORMAT_LABELS = {
    "MP3": "MP3", "FLAC": "FLAC", "OggVorbis": "OGG Vorbis",
    "OggOpus": "OGG Opus", "OggFLAC": "OGG FLAC", "OggSpeex": "OGG Speex",
    "MP4": "MP4 / M4A", "ASF": "WMA", "WAVE": "WAV", "AIFF": "AIFF",
    "Matroska": "MKV / WebM", "OggTheora": "OGG Theora",
    "APEv2": "APE", "Musepack": "Musepack", "WavPack": "WavPack",
    "TrueAudio": "TTA", "OptimFROG": "OptimFROG", "DSF": "DSF (DSD)",
    "DSDIFF": "DFF (DSD)",
}

# --------------------------------------------------------------------------
# Cestino di Windows via API nativa (nessuna dipendenza esterna)
# --------------------------------------------------------------------------
def send_to_recycle_bin(paths):
    """Sposta i percorsi nel Cestino usando SHFileOperationW (Windows)."""
    if not IS_WINDOWS or not paths:
        return False
    try:
        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", ctypes.c_void_p),
                ("wFunc", ctypes.c_uint),
                ("pFrom", ctypes.c_wchar_p),
                ("pTo", ctypes.c_wchar_p),
                ("fFlags", ctypes.c_ushort),
                ("fAnyOperationsAborted", ctypes.c_int),   # BOOL a 32 bit
                ("hNameMappings", ctypes.c_void_p),
                ("lpszProgressTitle", ctypes.c_wchar_p),
            ]
        FO_DELETE = 3
        FOF_ALLOWUNDO = 0x40
        FOF_NOCONFIRMATION = 0x10
        FOF_SILENT = 0x4
        src = "\0".join(str(p) for p in paths) + "\0\0"
        op = SHFILEOPSTRUCTW(None, FO_DELETE, src, None,
                             FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT,
                             0, None, None)
        sh_file_op = ctypes.windll.shell32.SHFileOperationW
        sh_file_op.argtypes = [ctypes.POINTER(SHFILEOPSTRUCTW)]
        sh_file_op.restype = ctypes.c_int
        result = sh_file_op(ctypes.byref(op))
        return result == 0
    except Exception:
        return False


def empty_recycle_bin():
    """Svuota il Cestino di Windows (SHEmptyRecycleBinW). True se riuscito."""
    if not IS_WINDOWS:
        return False
    try:
        # SHERB_NOCONFIRMATION=0x1 | SHERB_NOPROGRESSUI=0x2 | SHERB_NOSOUND=0x4:
        # la conferma viene chiesta dall'app, non dalla finestra di sistema
        she = ctypes.windll.shell32.SHEmptyRecycleBinW
        she.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint]
        she.restype = ctypes.c_long
        return she(None, None, 0x7) == 0
    except Exception:
        return False


def create_shortcut_lnk(target, lnk_path, description="", work_dir=""):
    """Crea un file .lnk verso 'target' usando IShellLinkW (solo Windows).

    Restituisce True in caso di successo; False su altri OS o in caso di
    errore (il chiamante decide come informare l'utente).
    """
    if not IS_WINDOWS:
        return False
    try:
        import ctypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", ctypes.c_uint32),
                        ("Data2", ctypes.c_uint16),
                        ("Data3", ctypes.c_uint16),
                        ("Data4", ctypes.c_ubyte * 8)]

        def _guid(s):
            d1, d2, d3, d4a, d4b = s.split("-")
            raw = bytes.fromhex(d4a + d4b)
            return GUID(int(d1, 16), int(d2, 16), int(d3, 16),
                        (ctypes.c_ubyte * 8)(*raw))

        CLSID_SHELLLINK = _guid("00021401-0000-0000-C000-000000000046")
        IID_ISHELLLINKW = _guid("000214F9-0000-0000-C000-000000000046")
        IID_IPERSISTFILE = _guid("0000010b-0000-0000-C000-000000000046")

        ole32 = ctypes.windll.ole32
        ole32.CoInitialize(None)   # COM va inizializzato sul thread corrente
        try:
            ole32.CoCreateInstance.argtypes = [
                ctypes.POINTER(GUID), ctypes.c_void_p, ctypes.c_uint32,
                ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)]
            ole32.CoCreateInstance.restype = ctypes.c_long

            shell_link = ctypes.c_void_p()
            hr = ole32.CoCreateInstance(
                ctypes.byref(CLSID_SHELLLINK), None, 1,   # CLSCTX_INPROC_SERVER
                ctypes.byref(IID_ISHELLLINKW), ctypes.byref(shell_link))
            if hr != 0:
                return False

            # vtable IShellLinkW: SetPath=20, SetDescription=7, SetWorkingDirectory=9
            vtbl = ctypes.cast(shell_link,
                               ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
            set_str = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.c_wchar_p)
            set_str(vtbl[20])(shell_link, str(target))            # SetPath
            if description:
                set_str(vtbl[7])(shell_link, str(description))    # SetDescription
            if work_dir:
                set_str(vtbl[9])(shell_link, str(work_dir))       # SetWorkingDirectory

            # IPersistFile per salvare su disco: QueryInterface (indice 0) poi Save (6)
            qi = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                    ctypes.POINTER(GUID),
                                    ctypes.POINTER(ctypes.c_void_p))
            persist = ctypes.c_void_p()
            hr = qi(vtbl[0])(shell_link, ctypes.byref(IID_IPERSISTFILE),
                             ctypes.byref(persist))
            if hr != 0:
                return False
            pvtbl = ctypes.cast(persist,
                                ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
            save = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                      ctypes.c_wchar_p, ctypes.c_int)
            hr = save(pvtbl[6])(persist, str(lnk_path), 1)        # Save

            rel = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
            rel(pvtbl[2])(persist)    # IPersistFile::Release
            rel(vtbl[2])(shell_link)  # IShellLinkW::Release
            return hr == 0
        finally:
            ole32.CoUninitialize()
    except Exception:
        return False


def format_size(num_bytes):
    """Formatta i byte in unità leggibili (KB, MB, GB...)."""
    try:
        n = float(num_bytes)
    except (TypeError, ValueError):
        return ""
    if n < 1024:
        return f"{int(n)} B"
    for unit in ("KB", "MB", "GB", "TB", "PB"):
        n /= 1024.0
        if n < 1024:
            return f"{n:.1f} {unit}"
    return f"{n:.1f} PB"


def parse_size(text):
    """Converte una stringa dimensione ('10 MB', '500 kb', '2.5 gb', '1024')
    in byte. Restituisce None se non è interpretabile."""
    try:
        s = str(text).strip().lower().replace(",", ".")
    except (TypeError, ValueError):
        return None
    if not s:
        return None
    units = {"b": 1, "kb": 1024, "mb": 1024 ** 2, "gb": 1024 ** 3,
             "tb": 1024 ** 4, "k": 1024, "m": 1024 ** 2,
             "g": 1024 ** 3, "t": 1024 ** 4}
    m = re.match(r"^([\d.]+)\s*([a-z]*)$", s)
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    unit = m.group(2)
    if unit and unit not in units:
        return None
    return int(num * units.get(unit, 1))


def format_date(ts):
    try:
        return datetime.datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M")
    except (OSError, ValueError, OverflowError):
        return ""


def format_duration(seconds):
    """Formatta i secondi in m:ss (o h:mm:ss oltre l'ora)."""
    try:
        s = int(round(seconds))
    except (TypeError, ValueError):
        return ""
    h, rem = divmod(s, 3600)
    m, sec = divmod(rem, 60)
    return f"{h}:{m:02d}:{sec:02d}" if h else f"{m}:{sec:02d}"


def natural_key(text):
    """Chiave di ordinamento naturale: 'file2' precede 'file10' (come un umano).

    Divide il testo in segmenti alfabetici e numerici; i segmenti numerici
    diventano int così '10' pesa più di '2'. L'alternanza garantita dal
    pattern (posizioni pari = stringhe, dispari = numeri) evita i confronti
    str/int incompatibili di Python 3.
    """
    return tuple(int(p) if i % 2 else p.lower()
                 for i, p in enumerate(re.split(r"(\d+)", str(text))))


def build_batch_name(model, start, index, path):
    """Nuovo nome per un file secondo il modello {n} (contatore),
    {name} (nome senza estensione) e {ext} (estensione col punto).
    Eventuali separatori di percorso nel modello vengono rimossi
    (os.path.basename) per non creare sottocartelle indesiderate."""
    base, ext = os.path.splitext(os.path.basename(path))
    new = (model.replace("{n}", str(start + index))
                .replace("{name}", base)
                .replace("{ext}", ext))
    return os.path.basename(new)


def readonly_text_widget(master, text, width=50):
    """Campo di testo selezionabile e copiabile ma NON modificabile.

    Un tk.Text con state='disabled' non permette la selezione, quindi il
    testo resta abilitato ma ogni tasto di modifica viene bloccato: si
    possono solo selezionare (mouse o Ctrl+A) e copiare (Ctrl+C o tasto
    destro). L'altezza è stimata dal contenuto così il testo è visibile.
    """
    height = max(1, math.ceil(len(text) / (width - 5)))
    txt = tk.Text(master, width=width, height=height, wrap="word",
                  font=("Segoe UI", 9), relief="flat", bd=0,
                  highlightthickness=0, padx=0, pady=0)
    txt.insert("1.0", text)
    try:
        bg = ttk.Style(master).lookup("TLabel", "background")
    except tk.TclError:
        bg = master.cget("bg")
    txt.configure(bg=bg, insertbackground=bg, cursor="arrow")

    def _on_key(event):
        ctrl = bool(event.state & 0x4)
        keysym = event.keysym
        if ctrl and keysym.lower() == "a":
            txt.tag_remove("sel", "1.0", "end")
            txt.tag_add("sel", "1.0", "end-1c")
            return "break"
        if ctrl and keysym.lower() == "c":
            # copia la selezione (o l'intero valore) e FERMA la propagazione:
            # altrimenti il Ctrl+C della finestra copierebbe tutte le proprietà
            try:
                sel = txt.get("sel.first", "sel.last")
            except tk.TclError:
                sel = ""
            if not sel:
                sel = txt.get("1.0", "end-1c")
            if sel:
                txt.clipboard_clear()
                txt.clipboard_append(sel)
            return "break"
        if ctrl:
            if keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                          "Prior", "Next"):
                return None      # navigazione con Ctrl: permessa
            return "break"      # altre combinazioni Ctrl: bloccate
        if event.char:
            return "break"      # digitazione bloccata (spazio compreso)
        return None              # frecce / Home / End / Shift+selezione

    txt.bind("<Key>", _on_key)
    return txt


# --------------------------------------------------------------------------
# Barra delle schede in stile browser (Canvas disegnato)
# --------------------------------------------------------------------------
class TabBar(tk.Frame):
    BG = "#2b2d30"
    TAB_INACTIVE = "#3a3d42"
    TAB_ACTIVE = "#f5f6f8"
    TEXT_INACTIVE = "#d0d3d8"
    TEXT_ACTIVE = "#1c1f23"
    ACCENT = "#4a90d9"
    CLOSE_HOVER = "#e81123"

    def __init__(self, master, on_select=None, on_close=None, on_new=None,
                 on_context=None):
        super().__init__(master, bg=self.BG, height=34)
        self.pack_propagate(False)
        self.on_select = on_select
        self.on_close = on_close
        self.on_new = on_new
        self.on_context = on_context

        self._tabs = []          # dict: {text, x0, x1}
        self._selected = -1
        self._scroll = 0.0
        self._max_scroll = 0.0
        self._font = ("Segoe UI", 9)
        self._measure = tkfont.Font(family="Segoe UI", size=9)

        self.canvas = tk.Canvas(self, bg=self.BG, height=30,
                                highlightthickness=0, bd=0)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Button-2>", self._on_middle_click)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<Double-1>", self._on_double_click)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Configure>", lambda e: self._redraw())

        self.btn_left = tk.Button(self, text="\u2039", bg=self.BG, fg=self.TEXT_INACTIVE,
                                  relief="flat", bd=0, width=2, activebackground="#4a4d53",
                                  command=lambda: self.scroll(-90))
        self.btn_right = tk.Button(self, text="\u203A", bg=self.BG, fg=self.TEXT_INACTIVE,
                                   relief="flat", bd=0, width=2, activebackground="#4a4d53",
                                   command=lambda: self.scroll(90))
        self.btn_plus = tk.Button(self, text="+", bg=self.BG, fg=self.TEXT_INACTIVE,
                                  relief="flat", bd=0, width=3,
                                  activebackground="#4a4d53", font=("Segoe UI", 11, "bold"),
                                  command=self._on_new)
        self.btn_plus.pack(side="right", fill="y")

    # ------------------------------------------------ API pubblica
    @property
    def selected(self):
        return self._selected

    def add_tab(self, text):
        self._tabs.append({"text": text, "x0": 0, "x1": 0})
        self._redraw()
        return len(self._tabs) - 1

    def insert_tab(self, index, text):
        """Inserisce una scheda alla posizione data (usato da 'riapri scheda chiusa')."""
        index = max(0, min(index, len(self._tabs)))
        self._tabs.insert(index, {"text": text, "x0": 0, "x1": 0})
        # se la scheda selezionata sta a destra (o coincide) del punto di
        # inserimento, sposta in avanti il suo indice per mantenerla puntata.
        if self._selected >= index:
            self._selected += 1
        self._redraw()
        return index

    def remove_tab(self, index):
        if not (0 <= index < len(self._tabs)):
            return
        self._tabs.pop(index)
        if self._selected > index:
            self._selected -= 1
        elif self._selected == index:
            self._selected = min(index, len(self._tabs) - 1)
        self._redraw()

    def set_title(self, index, text):
        if 0 <= index < len(self._tabs):
            self._tabs[index]["text"] = text
            self._redraw()

    def title_of(self, index):
        if 0 <= index < len(self._tabs):
            return self._tabs[index]["text"]
        return ""

    def select(self, index):
        if index < 0 or index >= len(self._tabs):
            return
        if index != self._selected:
            self._selected = index
            self._redraw()
            if self.on_select:
                self.on_select(index)

    def index_at(self, x, y):
        for i in range(len(self._tabs) - 1, -1, -1):
            t = self._tabs[i]
            if t["x0"] <= x <= t["x1"]:
                return i
        return -1

    def scroll(self, delta):
        self._scroll = max(0.0, min(self._scroll + delta, self._max_scroll))
        self._redraw()

    # ------------------------------------------------ disegno
    def _tab_width(self, text):
        return max(90, min(int(self._measure.measure(text)) + 48, 240))

    def _redraw(self):
        c = self.canvas
        c.delete("all")
        cw = max(c.winfo_width(), 120)
        x = 4 - self._scroll
        for i, t in enumerate(self._tabs):
            w = self._tab_width(t["text"])
            active = (i == self._selected)
            fill = self.TAB_ACTIVE if active else self.TAB_INACTIVE
            r = 7
            pts = [x, 2 + r, x + r, 2, x + w - r, 2, x + w, 2 + r,
                   x + w, 30, x, 30]
            c.create_polygon(pts, fill=fill, outline="")
            if active:
                c.create_line(x + 8, 30, x + w - 8, 30, fill=self.ACCENT, width=2)
            fg = self.TEXT_ACTIVE if active else self.TEXT_INACTIVE
            c.create_text(x + 14, 17, text=t["text"], anchor="w", fill=fg,
                          font=self._font)
            c.create_text(x + w - 13, 16, text="\u2715", fill="#8a8f96",
                          font=("Segoe UI", 8))
            t["x0"], t["x1"] = x, x + w
            x += w + 2
        total = x - 2
        self._max_scroll = max(0.0, total - cw + 8)
        if self._scroll > self._max_scroll:
            self._scroll = self._max_scroll
        if self._max_scroll > 0:
            # pack lato destro: l'ultimo va più a sinistra; così a destra del
            # canvas i bottoni compaiono nell'ordine [‹][›][+]
            self.btn_right.pack(side="right", fill="y")
            self.btn_left.pack(side="right", fill="y")
        else:
            self.btn_left.pack_forget()
            self.btn_right.pack_forget()

    # ------------------------------------------------ eventi
    def _on_new(self):
        if self.on_new:
            self.on_new()

    def _on_click(self, event):
        idx = self.index_at(event.x, event.y)
        if idx < 0:
            return
        if event.x >= self._tabs[idx]["x1"] - 20:
            if self.on_close:
                self.on_close(idx)
        else:
            self.select(idx)

    def _on_middle_click(self, event):
        idx = self.index_at(event.x, event.y)
        if idx >= 0 and self.on_close:
            self.on_close(idx)

    def _on_right_click(self, event):
        idx = self.index_at(event.x, event.y)
        if self.on_context:
            self.on_context(idx, event.x_root, event.y_root)

    def _on_double_click(self, event):
        if self.index_at(event.x, event.y) < 0:
            self._on_new()

    def _on_wheel(self, event):
        if self._max_scroll > 0:
            self.scroll(-event.delta)
        return "break"


# --------------------------------------------------------------------------
# Singola scheda: albero file + stato di navigazione
# --------------------------------------------------------------------------
class TabView:
    # etichette base delle colonne (vi si aggiunge ▲/▼ per l'ordinamento)
    HEADINGS = {"name": "Nome", "type": "Tipo", "size": "Dimensioni",
                "date": "Data modifica", "path": "Percorso"}

    def __init__(self, app):
        self.app = app
        c = app._colors
        self.frame = tk.Frame(app.pages, bg=c["page"])

        # Stato della scheda
        self.current_path = os.path.expanduser("~")
        self.history = []
        self.history_pos = -1
        self.entries = []           # (path, nome, is_dir, size, mtime, rel_path)
        self.sort_col = "name"
        self.sort_desc = False
        self.search_active = False

        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)

        columns = ("name", "type", "size", "date", "path")
        self.tree = ttk.Treeview(self.frame, columns=columns, show="headings",
                                 selectmode="extended")
        self.tree.heading("name", text="Nome", command=lambda: self._sort_by("name"))
        self.tree.heading("type", text="Tipo", command=lambda: self._sort_by("type"))
        self.tree.heading("size", text="Dimensioni", command=lambda: self._sort_by("size"))
        self.tree.heading("date", text="Data modifica", command=lambda: self._sort_by("date"))
        self.tree.heading("path", text="Percorso", command=lambda: self._sort_by("path"))
        self._update_sort_headers()
        self.tree.column("name", width=240, anchor="w", stretch=False)
        self.tree.column("type", width=130, anchor="w", stretch=False)
        self.tree.column("size", width=90, anchor="e", stretch=False)
        self.tree.column("date", width=130, anchor="w", stretch=False)
        self.tree.column("path", width=220, anchor="w", stretch=False)
        self.tree.configure(displaycolumns=("name", "type", "size", "date"))

        vsb = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # tag di stile (colori dal tema attivo)
        self.tree.tag_configure("folder", font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("odd", background=c["tree_odd"])
        self.tree.tag_configure("even", background=c["tree_even"])
        self.tree.tag_configure("dnd_target", background=c["dnd"])

        # eventi
        self.tree.bind("<Double-1>", app._on_double_click)
        self.tree.bind("<Return>", lambda e: app._open_selected())
        self.tree.bind("<BackSpace>", lambda e: app.go_up())
        self.tree.bind("<F2>", lambda e: app.rename_selected())
        self.tree.bind("<Delete>", lambda e: app.delete_selected())
        self.tree.bind("<F5>", lambda e: app.refresh())
        self.tree.bind("<Button-3>", app._show_context_menu)
        self.tree.bind("<<TreeviewSelect>>", lambda e: app._on_tree_select())
        self.tree.bind("<Control-a>", lambda e: self.tree.selection_set(self.tree.get_children()))
        self.tree.bind("<Alt-Left>", lambda e: app.go_back())
        self.tree.bind("<Alt-Right>", lambda e: app.go_forward())
        self.tree.bind("<Alt-Up>", lambda e: app.go_up())
        self.tree.bind("<Button-2>", app._on_tree_middle_click)
        # drag & drop
        self.tree.bind("<ButtonPress-1>", app._dnd_press)
        self.tree.bind("<B1-Motion>", app._dnd_motion)
        self.tree.bind("<ButtonRelease-1>", app._dnd_release)
        # tooltip con dettagli al passaggio del mouse
        self.tree.bind("<Motion>", app._on_tree_motion)
        self.tree.bind("<Leave>", app._on_tree_leave)

    # ----------------------------------------------------- colonne
    def _apply_columns(self):
        if self.search_active:
            cols = ("name", "path", "type", "size", "date") if self.app.show_details \
                else ("name", "path")
        else:
            cols = ("name", "type", "size", "date") if self.app.show_details \
                else ("name",)
        self.tree.configure(displaycolumns=cols)

    # ----------------------------------------------------- contenuto
    def refresh(self, keep_selection=True):
        """Ricarica la cartella corrente; se il filtro è attivo avvia la ricerca."""
        self.app._hide_tooltip()   # la vista cambia: via il tooltip stantio
        if self.app.filter_var.get().strip():
            self.search_active = True
            self._apply_columns()
            self.app._on_filter_changed()
            return

        self.search_active = False
        self._apply_columns()
        selected = set(self.tree.selection()) if keep_selection else set()
        try:
            items = list(os.scandir(self.current_path))
        except PermissionError:
            messagebox.showerror("Esplora File", f"Accesso negato:\n{self.current_path}")
            return
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile aprire la cartella:\n{exc}")
            return

        self.entries = []
        for it in items:
            name = it.name
            if not self.app.show_hidden and name.startswith("."):
                continue
            try:
                st = it.stat(follow_symlinks=True)
                is_dir = it.is_dir(follow_symlinks=True)
                size = st.st_size if not is_dir else 0
                mtime = st.st_mtime
            except OSError:
                is_dir = it.is_dir()
                size = 0
                mtime = 0
            self.entries.append((it.path, name, is_dir, size, mtime, ""))

        # filtro rapido per estensione (le cartelle restano sempre visibili)
        ext_filter = self.app._ext_filter
        if ext_filter:
            self.entries = [e for e in self.entries
                            if e[2] or os.path.splitext(e[1])[1].lower() == ext_filter]

        self._render_tree(selected)
        self.app._queue_auto_sizes()

    def _render_tree(self, selected=()):
        def key_func(entry):
            path, name, is_dir, size, mtime, rel = entry
            if self.sort_col == "name":
                k = (not is_dir, natural_key(name))
            elif self.sort_col == "type":
                k = (not is_dir, natural_key(self._type_label(entry)),
                     natural_key(name))
            elif self.sort_col == "size":
                # per le cartelle usa la dimensione in cache se disponibile
                # (altrimenti ordinano tutte come 0, ignorando la colonna)
                if not is_dir:
                    eff_size = size
                else:
                    eff_size = self.app.size_cache.cached_size(path) or 0
                k = (not is_dir, eff_size)
            elif self.sort_col == "path":
                k = (not is_dir, rel.lower(), name.lower())
            else:  # date
                k = (not is_dir, mtime)
            return k

        self.entries.sort(key=key_func, reverse=self.sort_desc)

        self.tree.delete(*self.tree.get_children())
        for i, entry in enumerate(self.entries):
            path, name, is_dir, size, mtime, rel = entry
            icon = ICON_FOLDER if is_dir else self._icon_for(name)
            tag = "folder" if is_dir else ""
            rowtag = (tag,) if tag else ()
            display = f"{icon} {name}"
            values = [display, self._type_label(entry),
                      self._size_cell(path, is_dir, size), format_date(mtime), rel]
            iid = self.tree.insert("", "end", iid=path, values=values, tags=rowtag)
            if iid in selected:
                self.tree.selection_add(iid)
        current = self.tree.selection()
        if current:
            self.tree.see(current[0])
        self.app._update_status()

    def _icon_for(self, name):
        ext = os.path.splitext(name)[1].lower()
        return ICON_EXTS.get(ext, ICON_MISC)

    def _size_cell(self, path, is_dir, size):
        """Valore della colonna Dimensioni: per le cartelle usa la cache."""
        if not is_dir:
            return format_size(size)
        cached = self.app.size_cache.cached_size(path)
        return format_size(cached) if cached is not None else ""

    def _type_label(self, entry):
        path, name, is_dir, size, mtime, rel = entry
        if is_dir:
            return "Cartella di file"
        ext = os.path.splitext(name)[1].lower()
        return TYPE_LABELS.get(ext, f"File {ext.upper().lstrip('.') or '(nessuna estensione)'}")

    def _sort_by(self, col):
        if self.sort_col == col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            self.sort_desc = False
        self._update_sort_headers()
        self.refresh(keep_selection=True)

    def _update_sort_headers(self):
        """Mostra una freccia (▲/▼) sull'intestazione della colonna ordinata."""
        for col, label in self.HEADINGS.items():
            if col == self.sort_col:
                arrow = " ▲" if not self.sort_desc else " ▼"
                self.tree.heading(col, text=label + arrow)
            else:
                self.tree.heading(col, text=label)


# --------------------------------------------------------------------------
# Anteprima file: immagini, testo, PDF e documenti Office (sola stdlib)
# --------------------------------------------------------------------------
def read_text_preview(path, limit=PREVIEW_TEXT_LIMIT):
    """Legge i primi 'limit' byte di un file come testo (o None se binario)."""
    try:
        with open(path, "rb") as f:
            data = f.read(limit)
    except OSError:
        return None
    if b"\x00" in data[:1024]:
        return None  # binario
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def extract_pdf_text(path, max_chars=8000):
    """Estrae testo da un PDF (best-effort) decodificando gli stream FlateDecode."""
    try:
        with open(path, "rb") as f:
            data = f.read(20 * 1024 * 1024)   # limite prudente
    except OSError:
        return ""
    texts = []
    pos = 0
    while True:
        s = data.find(b"stream", pos)
        if s < 0:
            break
        e = data.find(b"endstream", s)
        if e < 0:
            break
        header = data[max(0, s - 200):s]
        chunk = data[s + 6:e].lstrip(b"\r\n")   # lo stream PDF inizia dopo l'EOL
        if b"FlateDecode" in header:
            try:
                # decompressione CAPATA: un PDF malevolo non deve poter
                # espandere uno stream piccolo in gigabyte (decompression bomb)
                dobj = zlib.decompressobj()
                dec = dobj.decompress(chunk, MAX_PDF_STREAM)
                if not dobj.unconsumed_tail:
                    texts.append(_pdf_text_ops(dec))
                # stream più grande del limite: scartato (niente testi)
            except zlib.error:
                pass
        pos = e + 9
    text = "\n".join(t for t in texts if t)
    return text[:max_chars]


def _pdf_text_ops(content):
    """Estrae le stringhe (...) seguite da Tj / TJ da un content stream PDF."""
    out = []
    i, n = 0, len(content)
    while i < n:
        if content[i] == 40:  # '('
            j, depth = i + 1, 1
            buf = bytearray()
            while j < n and depth:
                c = content[j]
                if c == 92:  # backslash di escape
                    if j + 1 < n:
                        buf.append(content[j + 1])
                        j += 2
                    else:
                        j += 1
                elif c == 40:
                    depth += 1
                    buf.append(c)
                    j += 1
                elif c == 41:
                    depth -= 1
                    j += 1
                else:
                    buf.append(c)
                    j += 1
            k = j
            while k < n and content[k] in b" \r\n\t":
                k += 1
            # servono almeno 2 byte per i comandi Tj/TJ (in Python lo slicing
            # non lancia IndexError: restituisce una porzione più corta, ma
            # il controllo rende l'intento esplicito e difensivo)
            if k + 2 <= n and content[k:k + 2] in (b"Tj", b"TJ"):
                try:
                    s = buf.decode("latin-1")
                    if s.strip():
                        out.append(s)
                except UnicodeDecodeError:
                    pass
            i = j
        else:
            i += 1
    return " ".join(out)


# --------------------------------------------------------------------------
# Estrazione testo dai documenti Office (sola libreria standard)
# --------------------------------------------------------------------------
# <w:t>/<a:t> (con prefisso) e <t> (XLSX usa il namespace di default)
_T_TEXT_RE = re.compile(r"<(?:(?:w|a|m|t):)?t(?![\-\w:])[^>]*>(.*?)</(?:(?:w|a|m|t):)?t(?![\-\w:])>",
                        re.S)
_T_BREAK_RE = re.compile(r"</(?:w|a):p>|</w:tr>|</si>|<(?:w|a):br\s*/>", re.S)


def _ooxml_xml_text(xml_bytes):
    """Testo di una parte XML OOXML: contenuto dei tag di testo <w:t>/<a:t>/<t>.

    Ogni paragrafo (<w:p>, <a:p>, <w:tr>) e ogni stringa condivisa (<si>)
    va a capo; le entità XML, nominate e numeriche (&amp;, &#8217;, ...),
    vengono decodificate.
    """
    try:
        s = xml_bytes.decode("utf-8", errors="replace")
    except Exception:
        return ""
    out = []
    for para in _T_BREAK_RE.split(s):
        line = "".join(_T_TEXT_RE.findall(para))
        if not line.strip():
            continue
        out.append(_xml_unescape(line).strip())
    return "\n".join(out)


def extract_ooxml_text(path, max_chars=8000):
    """Estrae il testo di DOCX/PPTX/XLSX (archivi ZIP di XML OOXML).

    Legge i componenti che contengono il testo (word/document.xml con
    header/footer, ppt/slides/*.xml, xl/sharedStrings.xml) e ne ricava il
    testo leggibile. Ritorna '' se il file non è un OOXML valido.
    """
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile):
        return ""
    parts = []
    try:
        for name in zf.namelist()[:2000]:   # guardia anti zip con troppe voci
            low = name.lower()
            if not ("word/document.xml" in low
                    or low.startswith("word/header")
                    or low.startswith("word/footer")
                    or (low.startswith("ppt/slides/") and low.endswith(".xml"))
                    or low == "xl/sharedstrings.xml"):
                continue
            try:
                if zf.getinfo(name).file_size > MAX_PDF_STREAM:
                    continue   # componente enorme: scartato (anti bomba)
                parts.append(_ooxml_xml_text(zf.read(name)))
            except (KeyError, zipfile.BadZipFile, RuntimeError, zlib.error,
                    NotImplementedError, OSError):
                continue       # voce corrotta o cifrata: salta
    finally:
        zf.close()
    text = "\n\n".join(p for p in parts if p)
    return text[:max_chars]


def extract_legacy_office_text(path, max_chars=8000):
    """Estrae testo best-effort da DOC/PPT/XLS binari (formato OLE2).

    I vecchi documenti Office memorizzano il testo come sequenze UTF-16LE:
    si scansiona il file alla ricerca di tratti leggibili consecutivi.
    Non è una decodifica completa (il formato è chiuso), ma restituisce il
    testo principale dei documenti occidentali.
    """
    try:
        with open(path, "rb") as f:
            data = f.read(20 * 1024 * 1024)   # limite prudente
    except OSError:
        return ""
    # magic OLE2 (D0 CF 11 E0 A1 B1 1A E1): non è un documento binario Office
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return ""
    runs = []
    cur = []
    i, n = 0, len(data) - 1
    while i < n:
        lo, hi = data[i], data[i + 1]
        if hi == 0 and (0x20 <= lo <= 0x7E or 0xA0 <= lo <= 0xFF):
            cur.append(chr(lo))          # ASCII stampabile o Latin-1
            i += 2
        elif hi == 0 and lo in (0x0A, 0x0D):
            cur.append("\n")
            i += 2
        else:
            if cur:
                runs.append("".join(cur))
                cur = []
            i += 2
    if cur:
        runs.append("".join(cur))
    # tiene solo i tratti di almeno 3 caratteri: scarta il rumore binario
    out = [r.strip() for r in runs if len(r.strip()) >= 3]
    return "\n".join(out)[:max_chars]


# --------------------------------------------------------------------------
# Creazione di nuovi file (menu "Crea") — solo libreria standard
# --------------------------------------------------------------------------
CREATE_FILE_TYPES = {
    "txt":  (".txt",  "documento",     "File di testo"),
    "py":   (".py",   "script",        "File Python"),
    "docx": (".docx", "documento",     "Documento Word"),
    "pptx": (".pptx", "presentazione", "Presentazione PowerPoint"),
}


def create_blank_ooxml(path, kind):
    """Crea un DOCX o PPTX minimale ma valido (archivio ZIP OOXML).

    kind: 'docx' o 'pptx'. Il file si apre con Word/PowerPoint e il suo
    testo è estraibile dall'anteprima (vedi extract_ooxml_text).
    """
    if kind == "docx":
        parts = {
            "[Content_Types].xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '</Types>'),
            "_rels/.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '</Relationships>'),
            "word/document.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t>Nuovo documento</w:t></w:r></w:p>'
                '</w:body></w:document>'),
            "word/_rels/document.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
        }
    else:  # pptx
        parts = {
            "[Content_Types].xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
                '</Types>'),
            "_rels/.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                '</Relationships>'),
            "ppt/presentation.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>'
                '<p:sldSz cx="9144000" cy="6858000"/>'
                '<p:notesSz cx="6858000" cy="9144000"/>'
                '</p:presentation>'),
            "ppt/_rels/presentation.xml.rels": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
                '</Relationships>'),
            "ppt/slides/slide1.xml": (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<p:cSld><p:spTree>'
                '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                '<p:grpSpPr/>'
                '<p:sp>'
                '<p:nvSpPr><p:cNvPr id="2" name="Titolo"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm><a:off x="457200" y="411480"/><a:ext cx="8229600" cy="1600200"/></a:xfrm></p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/>'
                '<a:p><a:r><a:rPr lang="it-IT"/><a:t>Nuova presentazione</a:t></a:r></a:p>'
                '</p:txBody>'
                '</p:sp>'
                '</p:spTree></p:cSld>'
                '</p:sld>'),
        }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)


def decode_image_rgb(path):
    """Decodifica un'immagine (JPEG, BMP, TIFF, ICO...) via GDI+ su Windows.

    Ritorna (larghezza, altezza, byte_rgb) oppure None se non riesce.
    """
    if not IS_WINDOWS:
        return None
    try:
        gdiplus = ctypes.windll.gdiplus

        class GdiplusStartupInput(ctypes.Structure):
            _fields_ = [("GdiplusVersion", ctypes.c_uint32),
                        ("DebugEventCallback", ctypes.c_void_p),
                        ("SuppressBackgroundThread", ctypes.c_int),
                        ("SuppressExternalCodecs", ctypes.c_int)]

        class GpRect(ctypes.Structure):
            _fields_ = [("X", ctypes.c_int), ("Y", ctypes.c_int),
                        ("Width", ctypes.c_int), ("Height", ctypes.c_int)]

        class GpBitmapData(ctypes.Structure):
            _fields_ = [("Width", ctypes.c_uint),
                        ("Height", ctypes.c_uint),
                        ("Stride", ctypes.c_int),
                        ("PixelFormat", ctypes.c_int),
                        ("Scan0", ctypes.c_void_p),
                        ("Reserved", ctypes.c_uint)]

        gdiplus.GdiplusStartup.argtypes = [
            ctypes.POINTER(ctypes.c_size_t),
            ctypes.POINTER(GdiplusStartupInput),
            ctypes.c_void_p]
        gdiplus.GdiplusStartup.restype = ctypes.c_int
        gdiplus.GdipCreateBitmapFromFile.argtypes = [
            ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_void_p)]
        gdiplus.GdipCreateBitmapFromFile.restype = ctypes.c_int
        gdiplus.GdipGetImageWidth.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint)]
        gdiplus.GdipGetImageHeight.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(ctypes.c_uint)]
        gdiplus.GdipBitmapLockBits.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(GpRect), ctypes.c_uint,
            ctypes.c_int, ctypes.POINTER(GpBitmapData)]
        gdiplus.GdipBitmapUnlockBits.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(GpBitmapData)]
        gdiplus.GdipDisposeImage.argtypes = [ctypes.c_void_p]
        gdiplus.GdiplusShutdown.argtypes = [ctypes.c_size_t]

        token = ctypes.c_size_t()
        inp = GdiplusStartupInput(1, None, 0, 0)
        if gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(inp), None) != 0:
            return None
        try:
            bitmap = ctypes.c_void_p()
            if gdiplus.GdipCreateBitmapFromFile(
                    ctypes.c_wchar_p(path), ctypes.byref(bitmap)) != 0:
                return None
            try:
                w, h = ctypes.c_uint(), ctypes.c_uint()
                if (gdiplus.GdipGetImageWidth(bitmap, ctypes.byref(w)) != 0 or
                        gdiplus.GdipGetImageHeight(bitmap, ctypes.byref(h)) != 0):
                    return None
                # controlla la dimensione PRIMA di decodificare in memoria
                if w.value * h.value > PREVIEW_IMAGE_MAX_PIXELS:
                    return None
                rect = GpRect(0, 0, w.value, h.value)
                bd = GpBitmapData()
                PIXELFORMAT_32ARGB = 0x0026200A
                LOCKREAD = 0x1
                if gdiplus.GdipBitmapLockBits(
                        bitmap, ctypes.byref(rect), LOCKREAD,
                        PIXELFORMAT_32ARGB, ctypes.byref(bd)) != 0:
                    return None
                try:
                    # con stride negativo Scan0 punta alla riga IN BASSO;
                    # l'inizio del buffer (riga in alto) è Scan0 + (h-1)*stride
                    if bd.Stride < 0:
                        start = ctypes.c_void_p(
                            bd.Scan0 + (h.value - 1) * bd.Stride)
                    else:
                        start = ctypes.c_void_p(bd.Scan0)
                    raw = ctypes.string_at(start, abs(bd.Stride) * h.value)
                finally:
                    gdiplus.GdipBitmapUnlockBits(bitmap, ctypes.byref(bd))
                # dopo l'aggiustamento la prima riga è sempre quella in alto;
                # 32bppARGB in memoria è memorizzato come BGRA
                n = w.value * h.value
                rgb = bytearray(n * 3)
                stride = abs(bd.Stride)
                pos = 0
                for y in range(h.value):
                    row = y * stride
                    for x in range(w.value):
                        i = row + x * 4
                        rgb[pos] = raw[i + 2]      # R
                        rgb[pos + 1] = raw[i + 1]  # G
                        rgb[pos + 2] = raw[i]      # B
                        pos += 3
                return (w.value, h.value, bytes(rgb))
            finally:
                gdiplus.GdipDisposeImage(bitmap)
        finally:
            gdiplus.GdiplusShutdown(token)
    except Exception:
        return None


def photo_from_rgb(width, height, rgb):
    """Crea una tk.PhotoImage dai byte RGB SENZA file temporanei.

    Costruisce un PNG in memoria (zlib + struct) e lo passa a Tk via
    base64: niente TOCTOU, niente file orfani, niente PermissionError su
    Windows. Se Tk non accettasse il PNG (versione molto vecchia), ricade
    sul vecchio file PPM temporaneo.
    """
    try:
        png = _png_from_rgb(width, height, rgb)
        return tk.PhotoImage(data=base64.b64encode(png))
    except (tk.TclError, MemoryError):
        pass
    # fallback: file PPM temporaneo (rimosso subito dopo la creazione)
    header = f"P6\n{width} {height}\n255\n".encode("ascii")
    fd, tmp = tempfile.mkstemp(suffix=".ppm")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(header)
            f.write(rgb)
        return tk.PhotoImage(file=tmp)
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def _png_from_rgb(width, height, rgb):
    """Codifica byte RGB in un PNG (senza dipendenze esterne)."""
    def chunk(typ, data):
        c = struct.pack(">I", len(data)) + typ + data
        return c + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF)
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8-bit RGB
    raw = bytearray()
    for y in range(height):
        raw.append(0)   # filtro None per ogni riga
        raw += rgb[y * width * 3:(y + 1) * width * 3]
    idat = zlib.compress(bytes(raw), 6)
    return (sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat)
            + chunk(b"IEND", b""))


# --------------------------------------------------------------------------
# Archivi ZIP, dimensione cartelle (con cache persistente) e "Apri con…"
# --------------------------------------------------------------------------
def create_zip(paths, dest_zip):
    """Crea un archivio ZIP con i percorsi indicati (file e cartelle).

    Ritorna None in caso di successo, altrimenti il messaggio di errore.
    """
    try:
        with zipfile.ZipFile(dest_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            used = set()
            for p in paths:
                p = os.path.normpath(p)
                if os.path.isdir(p):
                    top = os.path.basename(p)
                    name, n = top, 1
                    while name in used:
                        name = f"{top} ({n})"
                        n += 1
                    used.add(name)
                    for dirpath, _dirnames, filenames in os.walk(p):
                        for fn in filenames:
                            fp = os.path.join(dirpath, fn)
                            arc = os.path.join(
                                name, os.path.relpath(fp, p)).replace("\\", "/")
                            zf.write(fp, arc)
                else:
                    name = os.path.basename(p)
                    arc, n = name, 1
                    while arc in used:
                        stem, ext = os.path.splitext(name)
                        arc = f"{stem} ({n}){ext}"
                        n += 1
                    used.add(arc)
                    zf.write(p, arc)
        return None
    except (OSError, zipfile.BadZipFile) as exc:
        return str(exc)


def extract_zip(zip_path, dest_dir):
    """Estrae uno ZIP in una cartella, ignorando i nomi non sicuri.

    La verifica di sicurezza usa os.path.realpath + os.path.commonpath:
    robusta contro percorsi assoluti, doppi separatori, '..' e casi limite
    (es. un membro chiamato esattamente come la cartella di destinazione).
    """
    try:
        os.makedirs(dest_dir, exist_ok=True)
        base = os.path.realpath(dest_dir)
        with zipfile.ZipFile(zip_path) as zf:
            for member in zf.infolist():
                name = member.filename.replace("\\", "/")
                target = os.path.realpath(os.path.join(base, name))
                try:
                    if os.path.commonpath([base, target]) != base:
                        continue   # path traversal: salta il membro
                except ValueError:
                    continue       # drive diversi (es. C: vs D:): salta
                zf.extract(member, dest_dir)
        return None
    except (OSError, zipfile.BadZipFile) as exc:
        return str(exc)


# --------------------------------------------------------------------------
# Archivi 7z / RAR / TAR: 7-Zip (se installato) o py7zr (pura Python) per
# il 7z; il RAR può essere CREATO solo da WinRAR (formato proprietario).
# --------------------------------------------------------------------------
def _seven_zip_exe():
    """Percorso di 7z.exe (7-Zip) se installato, altrimenti None."""
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            p = os.path.join(base, "7-Zip", "7z.exe")
            if os.path.isfile(p):
                return p
    return shutil.which("7z")


def _winrar_exe():
    """Percorso di Rar.exe / WinRAR.exe se installati, altrimenti None."""
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            for name in ("Rar.exe", "WinRAR.exe"):
                p = os.path.join(base, "WinRAR", name)
                if os.path.isfile(p):
                    return p
    return None


def _run_tool(cmd):
    """Esegue un tool esterno (7z/rar) e ritorna (returncode, output)."""
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              errors="replace",
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return proc.returncode, (proc.stdout or "") + (proc.stderr or "")
    except OSError as exc:
        return -1, str(exc)


def _seven_zip_list(exe, arc, limit=500):
    """Elenca il contenuto di un archivio via 7z.exe: (nome, dimensione, cartella)."""
    try:
        proc = subprocess.run([exe, "l", "-slt", arc], capture_output=True,
                              text=True, errors="replace",
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except OSError:
        return []
    out = proc.stdout or ""
    entries = []
    cur = {}
    for line in out.splitlines():
        line = line.rstrip()
        if line.startswith("Path = "):
            cur["path"] = line[7:]
        elif line.startswith("Size = "):
            try:
                cur["size"] = int(line[7:])
            except ValueError:
                cur["size"] = 0
        elif line.startswith("Folder = "):
            cur["folder"] = line[9:].strip() == "+"
        elif line.strip() == "" and "path" in cur:
            entries.append((cur["path"], cur.get("size", 0),
                            cur.get("folder", False)))
            cur = {}
            if len(entries) >= limit:
                break
    if "path" in cur:
        entries.append((cur["path"], cur.get("size", 0),
                        cur.get("folder", False)))
    return entries


def extract_tar(arc, dest):
    """Estrae un archivio TAR (anche .tar.gz/.tgz/.tar.bz2/.tar.xz)."""
    try:
        os.makedirs(dest, exist_ok=True)
        base = os.path.realpath(dest)
        with tarfile.open(arc, "r:*") as tf:
            for m in tf.getmembers():
                if m.issym() or m.islnk():
                    continue   # niente link (sicurezza)
                target = os.path.realpath(os.path.join(base, m.name))
                try:
                    if os.path.commonpath([base, target]) != base:
                        continue   # path traversal: salta il membro
                except ValueError:
                    continue
                # difesa in profondità: riscrivi il nome con il percorso
                # relativo già verificato, così l'estrazione usa esattamente
                # il target sicuro (mai il nome originale dell'archivio)
                rel = os.path.relpath(target, base)
                if not rel or rel == ".":
                    continue   # membro che punta alla destinazione stessa
                m.name = rel
                tf.extract(m, dest)
        return None
    except (OSError, tarfile.TarError) as exc:
        return str(exc)


def extract_7z(arc, dest):
    """Estrae un archivio 7z: via 7-Zip se presente, altrimenti py7zr."""
    os.makedirs(dest, exist_ok=True)
    exe = _seven_zip_exe()
    if exe:
        rc, out = _run_tool([exe, "x", arc, f"-o{dest}", "-y"])
        if rc <= 1:   # 0 = ok, 1 = avviso non fatale (file comunque estratti)
            return None
        return f"7-Zip: {out.strip()[:300]}"
    try:
        import py7zr
    except ImportError:
        return ("Per estrarre archivi 7z installa 7-Zip\n"
                "(https://www.7-zip.org/) oppure la libreria 'py7zr'\n"
                "(pip install py7zr).")
    try:
        base = os.path.realpath(dest)
        with py7zr.SevenZipFile(arc, "r") as z:
            for name in z.getnames():
                target = os.path.realpath(os.path.join(base, name))
                try:
                    if os.path.commonpath([base, target]) != base:
                        return "Archivio non sicuro (nome non valido)."
                except ValueError:
                    return "Archivio non sicuro (nome non valido)."
            z.extractall(path=dest)
        return None
    except Exception as exc:
        return str(exc)


def extract_rar(arc, dest):
    """Estrae un archivio RAR (richiede 7-Zip o WinRAR)."""
    os.makedirs(dest, exist_ok=True)
    exe = _seven_zip_exe() or _winrar_exe()
    if not exe:
        return ("Per estrarre archivi RAR installa 7-Zip\n"
                "(https://www.7-zip.org/) o WinRAR.")
    if exe.lower().endswith("7z.exe"):
        rc, out = _run_tool([exe, "x", arc, f"-o{dest}", "-y"])
    else:
        rc, out = _run_tool([exe, "x", "-y", arc, dest + os.sep])
    if rc <= 1:   # 0 = ok, 1 = avviso non fatale
        return None
    return f"{os.path.basename(exe)}: {out.strip()[:300]}"


def create_7z(paths, dest):
    """Crea un archivio 7z: via 7-Zip se presente, altrimenti py7zr."""
    exe = _seven_zip_exe()
    if exe:
        rc, out = _run_tool([exe, "a", "-t7z", "-y", dest] + list(paths))
        if rc <= 1:   # 0 = ok, 1 = avviso non fatale
            return None
        return f"7-Zip: {out.strip()[:300]}"
    try:
        import py7zr
    except ImportError:
        return ("Per creare archivi 7z installa 7-Zip\n"
                "(https://www.7-zip.org/) oppure la libreria 'py7zr'\n"
                "(pip install py7zr).")
    try:
        with py7zr.SevenZipFile(dest, "w") as z:
            for p in paths:
                p = os.path.normpath(p)
                if os.path.isdir(p):
                    z.writeall(p, arcname=os.path.basename(p))
                else:
                    z.write(p, arcname=os.path.basename(p))
        return None
    except Exception as exc:
        return str(exc)


def create_rar(paths, dest):
    """Crea un archivio RAR (richiede WinRAR: formato proprietario)."""
    exe = _winrar_exe()
    if not exe:
        return ("La creazione di archivi RAR richiede WinRAR:\n"
                "il formato RAR è proprietario e solo WinRAR può crearlo.\n"
                "Usa 7z o ZIP come alternativa.")
    rc, out = _run_tool([exe, "a", "-r", dest] + list(paths))
    if rc <= 1:   # 0 = ok, 1 = avviso non fatale
        return None
    return f"{os.path.basename(exe)}: {out.strip()[:300]}"


SECURE_DELETE_PASSES = 3   # passaggi di sovrascrittura prima della cancellazione


FILE_ATTRIBUTE_REPARSE_POINT = 0x400   # attributo Windows dei symlink/junction


def _is_link_or_junction(path):
    """True per symlink E junction point, su TUTTE le versioni di Python.

    Su Windows i junction NON vengono rilevati da os.path.islink (e
    os.path.isjunction esiste solo dalla 3.12): si controlla quindi
    l'attributo FILE_ATTRIBUTE_REPARSE_POINT via os.lstat, disponibile
    ovunque. Senza questo controllo una scansione seguirebbe i junction
    e cancellerebbe file FUORI dalla cartella selezionata.
    """
    if os.path.islink(path):
        return True
    if IS_WINDOWS:
        try:
            return bool(os.lstat(path).st_file_attributes
                        & FILE_ATTRIBUTE_REPARSE_POINT)
        except OSError:
            return False
    return False


def _is_drive_root(path):
    """True se il percorso è una radice di unità (es. C:\\, C:, D:\\)"""
    drive, tail = os.path.splitdrive(path)
    if not drive:
        return False
    root = os.path.normpath(drive + os.sep)
    return os.path.normpath(path) in (root, os.path.normpath(drive))


DESTROY_SIZE_LIMIT = 5 * 1024 ** 3         # limite distruzione: 5 GB
DESTROY_TYPING_LIMIT = 500 * 1024 * 1024   # oltre: seconda conferma con digitazione


def _is_system_path(path):
    """True se il percorso è una cartella critica di sistema o sta DENTRO
    una di esse (es. C:\\Windows\\System32).

    Il confronto usa normcase+normpath: 'C:\\Windows\\', 'c:/windows' e
    'C:\\Windows' sono tutti riconosciuti. Proteggere anche i discendenti
    è voluto: una sottocartella di Windows o di Program Files non è meno
    critica della radice.
    """
    p = os.path.normcase(os.path.normpath(path))
    for env in ("WINDIR", "SystemRoot", "ProgramFiles",
                "ProgramFiles(x86)", "ProgramData"):
        v = os.environ.get(env)
        if v:
            root = os.path.normcase(os.path.normpath(v))
            if p == root or p.startswith(root + os.sep):
                return True
    return False


def delete_tree_no_follow(root, errors):
    """Rimuove ricorsivamente una cartella senza MAI seguire symlink/junction.

    I puntatori (symlink e junction point di Windows) vengono rimossi da
    soli, mai il contenuto a cui puntano (che potrebbe stare FUORI dalla
    selezione). Usata da delete_selected al posto di shutil.rmtree, che su
    Windows segue i junction e cancellerebbe il contenuto del target.
    """
    try:
        entries = list(os.scandir(root))
    except OSError as exc:
        errors.append(f"{os.path.basename(root)}: {exc}")
        return
    for e in entries:
        path = e.path
        if e.is_symlink() or _is_link_or_junction(path):
            # rimuovi SOLO il puntatore, mai il contenuto puntato
            try:
                os.remove(path)
            except OSError:
                try:
                    os.rmdir(path)
                except OSError as exc:
                    errors.append(f"{os.path.basename(path)}: {exc}")
            continue
        try:
            if e.is_dir(follow_symlinks=False):
                delete_tree_no_follow(path, errors)
                try:
                    os.rmdir(path)
                except OSError as exc:
                    errors.append(f"{os.path.basename(path)}: {exc}")
            else:
                os.remove(path)
        except OSError as exc:
            errors.append(f"{os.path.basename(path)}: {exc}")


def _folder_size_limited(root, limit, acc):
    """Somma a acc[0] la dimensione dei file sotto root (senza seguire
    symlink/junction), fermandosi appena si supera 'limit'."""
    if acc[0] > limit:
        return
    try:
        with os.scandir(root) as it:
            for e in it:
                if acc[0] > limit:
                    break
                if e.is_symlink() or _is_link_or_junction(e.path):
                    continue
                try:
                    if e.is_dir(follow_symlinks=False):
                        _folder_size_limited(e.path, limit, acc)
                    else:
                        acc[0] += e.stat(follow_symlinks=False).st_size
                except OSError:
                    continue
    except OSError:
        pass


def _selection_size(paths, limit):
    """(dimensione_totale, oltre_limite) della selezione, senza seguire link."""
    acc = [0]
    for p in paths:
        if _is_link_or_junction(p):
            continue
        if os.path.isdir(p):
            _folder_size_limited(p, limit, acc)
        else:
            try:
                acc[0] += os.lstat(p).st_size
            except OSError:
                pass
        if acc[0] > limit:
            break
    return acc[0], acc[0] > limit


def secure_delete(paths, passes=SECURE_DELETE_PASSES, progress_cb=None):
    """Cancella DEFINITIVAMENTE sovrascrivendo i dati con byte casuali.

    Ogni file viene riscritto 'passes' volte con byte casuali
    (os.urandom, a blocchi), forzato su disco con fsync e poi cancellato.
    La scansione delle cartelle NON segue MAI symlink o junction point:
    i puntatori vengono rimossi da soli, mai il contenuto a cui puntano.
    progress_cb(fatti, totale) viene chiamata a ogni file.
    Ritorna la lista degli errori (vuota = tutto ok).

    NOTA: su SSD la sovrascrittura NON garantisce la cancellazione fisica
    (wear leveling e TRIM spostano i dati); per dati molto sensibili usa
    la funzione Secure Erase del disco o la cifratura dell'intero volume.
    """
    errors = []
    done = 0

    def _count_safe(root):
        """Conta i file reali sotto root senza mai seguire symlink/junction."""
        n = 0
        try:
            for e in os.scandir(root):
                if e.is_symlink() or _is_link_or_junction(e.path):
                    continue
                if e.is_dir(follow_symlinks=False):
                    n += _count_safe(e.path)
                else:
                    n += 1
        except OSError:
            pass
        return n

    total = 0
    for p in paths:
        if _is_link_or_junction(p) or not os.path.isdir(p):
            total += 1
        else:
            total += _count_safe(p)

    def _bump():
        nonlocal done
        done += 1
        if progress_cb:
            progress_cb(done, total)

    def _wipe_file(file_path):
        try:
            size = os.path.getsize(file_path)
            with open(file_path, "r+b") as f:
                for _ in range(passes):
                    f.seek(0)
                    remaining = size
                    while remaining > 0:
                        chunk = min(remaining, 1024 * 1024)
                        f.write(os.urandom(chunk))
                        remaining -= chunk
                    f.flush()
                    os.fsync(f.fileno())
            os.remove(file_path)
        except OSError as exc:
            errors.append(f"{os.path.basename(file_path)}: {exc}")
        _bump()

    def _remove_pointer(path):
        """Rimuove SOLO un symlink/junction, mai il contenuto puntato."""
        try:
            os.remove(path)
        except OSError:
            try:
                os.rmdir(path)
            except OSError as exc:
                errors.append(f"{os.path.basename(path)}: {exc}")

    def _destroy_tree(root):
        """Cancella ricorsivamente root senza MAI seguire symlink/junction."""
        try:
            entries = list(os.scandir(root))
        except OSError as exc:
            errors.append(f"{os.path.basename(root)}: {exc}")
            return
        for e in entries:
            path = e.path
            if e.is_symlink() or _is_link_or_junction(path):
                _remove_pointer(path)
                continue
            try:
                if e.is_dir(follow_symlinks=False):
                    _destroy_tree(path)
                    try:
                        os.rmdir(path)
                    except OSError as exc:
                        errors.append(f"{os.path.basename(path)}: {exc}")
                else:
                    _wipe_file(path)
            except OSError as exc:
                # 'path' è già assegnata qui (all'inizio del loop): l'errore
                # va REGISTRATO, non inghiottito in silenzio
                errors.append(f"{os.path.basename(path)}: {exc}")

    for p in paths:
        if _is_link_or_junction(p):
            _remove_pointer(p)
            _bump()
        elif os.path.isdir(p):
            _destroy_tree(p)
            try:
                os.rmdir(p)
            except OSError as exc:
                errors.append(f"{os.path.basename(p)} (cartella): {exc}")
        else:
            _wipe_file(p)
    return errors


def file_hashes(path, chunk=1024 * 1024, progress_cb=None):
    """Calcola gli hash MD5 e SHA-256 di un file in UNA sola lettura.

    Ritorna (md5, sha256) oppure None se il file non è leggibile.
    progress_cb(byte_letti, byte_totali) viene chiamata a ogni blocco.
    """
    md5 = hashlib.md5()
    sha256 = hashlib.sha256()
    letti = 0
    try:
        totale = os.path.getsize(path)
        with open(path, "rb") as f:
            while True:
                b = f.read(chunk)
                if not b:
                    break
                md5.update(b)
                sha256.update(b)
                letti += len(b)
                if progress_cb:
                    progress_cb(letti, totale)
        return md5.hexdigest(), sha256.hexdigest()
    except OSError:
        return None


def file_md5(path, chunk=1024 * 1024):
    """Hash MD5 di un file letto a blocchi (o None se illeggibile)."""
    res = file_hashes(path, chunk)
    return res[0] if res else None


def find_duplicate_files(root, progress_cb=None):
    """Trova gruppi di file duplicati (stesso contenuto) sotto root.

    Ritorna una lista di gruppi [(dimensione, [percorso, ...]), ...] ordinati
    per dimensione decrescente. Prima raggruppa per dimensione (veloce), poi
    conferma con l'hash MD5 dei file della stessa dimensione.
    """
    by_size = {}
    count = 0
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            p = os.path.join(dirpath, fn)
            try:
                sz = os.path.getsize(p)
            except OSError:
                continue
            by_size.setdefault(sz, []).append(p)
            count += 1
            if progress_cb and count % 100 == 0:
                progress_cb(count)
    if progress_cb:
        progress_cb(count)
    groups = []
    for sz, paths in by_size.items():
        if len(paths) < 2:
            continue
        by_hash = {}
        for p in paths:
            h = file_md5(p)
            if h:
                by_hash.setdefault(h, []).append(p)
        for h, ps in by_hash.items():
            if len(ps) > 1:
                groups.append((sz, ps))
    groups.sort(key=lambda g: -g[0])
    return groups


def count_files(path):
    """Numero di file regolari sotto path (senza seguire i symlink)."""
    if not os.path.isdir(path):
        return 1
    n = 0
    for _dirpath, _dirs, files in os.walk(path):
        n += len(files)
    return n


def copy_tree_with_progress(src, dst, progress_cb):
    """Copia ricorsivamente src in dst, chiamando progress_cb() per ogni file.

    Usa shutil.copytree (stesso comportamento della copia normale) ma con una
    copy_function che notifica ogni singolo file copiato.
    """
    def _copy2(s, d, **kwargs):
        shutil.copy2(s, d, **kwargs)
        progress_cb()
    shutil.copytree(src, dst, copy_function=_copy2)


def _split_path(path):
    """Scompone un percorso in segmenti (etichetta, percorso completo).

    Ritorna la lista ordinata dalla radice alla foglia.
    """
    path = os.path.normpath(path)
    if not path:
        return []
    parts = []
    cur = path
    while True:
        head, tail = os.path.split(cur)
        if tail:
            parts.append((tail, cur))
        if head == cur:  # radice (unita' oppure /)
            parts.append((head.rstrip("\\/") or head, head))
            break
        cur = head
        if not cur:
            break
    parts.reverse()
    return parts


def _settings_file():
    """Percorso del file JSON delle impostazioni dell'app."""
    if IS_WINDOWS and os.environ.get("APPDATA"):
        base = os.path.join(os.environ["APPDATA"], "EsploraFile")
    else:
        base = os.path.join(os.path.expanduser("~"), ".esplora_file")
    return os.path.join(base, "settings.json")


def load_settings():
    """Legge le impostazioni salvate (JSON), o {} se assenti o corrotte."""
    try:
        with open(_settings_file(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data):
    """Salva le impostazioni (JSON) in modo silenzioso."""
    try:
        os.makedirs(os.path.dirname(_settings_file()), exist_ok=True)
        with open(_settings_file(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except OSError:
        pass


class FolderSizeCache:
    """Cache persistente delle dimensioni delle cartelle (JSON su disco).

    Ogni voce: {size, count, mtime} con chiave il percorso normalizzato.
    Thread-safe: get/set/cached_size/save sono protette da un unico RLock
    (compute_folder_size gira in un thread di background).
    """

    def __init__(self, path=None):
        if path is None:
            if IS_WINDOWS and os.environ.get("APPDATA"):
                base = os.path.join(os.environ["APPDATA"], "EsploraFile")
            else:
                base = os.path.join(os.path.expanduser("~"), ".esplora_file")
            path = os.path.join(base, "folder_sizes.json")
        self.path = path
        self.data = {}
        self._lock = threading.RLock()
        self.load()

    def load(self):
        with self._lock:
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.data = data if isinstance(data, dict) else {}
            except (OSError, ValueError):
                self.data = {}

    def save(self):
        with self._lock:
            try:
                os.makedirs(os.path.dirname(self.path), exist_ok=True)
                with open(self.path, "w", encoding="utf-8") as f:
                    json.dump(self.data, f, ensure_ascii=False)
            except OSError:
                pass

    def _key(self, path):
        return os.path.normcase(os.path.normpath(path))

    def get(self, path):
        with self._lock:
            return self.data.get(self._key(path))

    def set(self, path, size, count, mtime):
        with self._lock:
            self.data[self._key(path)] = {
                "size": size, "count": count, "mtime": mtime}

    def cached_size(self, path):
        """Dimensione ricordata, o None se la cartella è cambiata (mtime)."""
        entry = self.get(path)
        if not entry:
            return None
        try:
            if abs(os.stat(path).st_mtime - entry["mtime"]) > 0.01:
                return None
        except OSError:
            return None
        return entry["size"]


def compute_folder_size(root, cache, progress_cb=None, _seen=None, _depth=0,
                        _acc=None):
    """Calcola (dimensione, numero_file) di una cartella, ricorsivamente.

    Percorre sempre l'albero (rileva ogni modifica) e aggiorna la cache con
    i valori correnti, salvandola periodicamente per riprendere in caso di
    interruzione. Evita i loop dei collegamenti simbolici.

    progress_cb(byte_totali, numero_file) viene chiamata dopo ogni cartella
    con il totale CUMULATO finora: il numero sale man mano che la scansione
    procede e arriva al risultato finale esatto.
    """
    if _depth > 900:
        return (0, 0)   # guardia anti ricorsione per alberi patologici
    if _seen is None:
        _seen = set()
    if _acc is None:
        _acc = [0, 0]   # totale cumulato condiviso (byte, file)
    real = os.path.realpath(root)
    if real in _seen:
        return (0, 0)
    _seen.add(real)
    total, count = 0, 0
    try:
        with os.scandir(root) as it:
            entries = list(it)
        st = os.stat(root)
    except OSError:
        return (0, 0)
    for e in entries:
        try:
            if e.is_dir(follow_symlinks=False):
                s, c = compute_folder_size(e.path, cache, progress_cb,
                                           _seen, _depth + 1, _acc)
                total += s
                count += c
            else:
                est = e.stat(follow_symlinks=False)
                total += est.st_size
                count += 1
                _acc[0] += est.st_size    # ogni file è contato una sola volta
                _acc[1] += 1
        except OSError:
            continue
    if progress_cb:
        progress_cb(_acc[0], _acc[1])
    # Cache DOPO il progress: se una scansione automatica viene annullata qui
    # (il callback lancia _ScanCancelled), una dimensione PARZIALE non deve
    # finire in cache come se fosse completa.
    cache.set(root, total, count, st.st_mtime)
    return (total, count)


def folder_stats(root, progress_cb=None):
    """Statistiche di una cartella: file, cartelle, byte, per estensione.

    Ritorna un dict:
      files, folders, bytes_total, per_ext {estensione: [conteggio, byte]},
      largest (percorso, dimensione), avg (byte medi per file).
    La scansione NON segue symlink/junction (sicurezza, come secure_delete).
    progress_cb(numero_file) viene chiamata ogni ~100 file.
    """
    stats = {"files": 0, "folders": 0, "bytes_total": 0,
             "per_ext": {}, "largest": (None, 0), "avg": 0.0}

    def walk(p):
        try:
            with os.scandir(p) as it:
                entries = list(it)
        except OSError:
            return
        for e in entries:
            if e.is_symlink() or _is_link_or_junction(e.path):
                continue
            try:
                if e.is_dir(follow_symlinks=False):
                    stats["folders"] += 1
                    walk(e.path)
                else:
                    st = e.stat(follow_symlinks=False)
                    size = st.st_size
                    stats["files"] += 1
                    stats["bytes_total"] += size
                    if size > stats["largest"][1]:
                        stats["largest"] = (e.path, size)
                    ext = os.path.splitext(e.name)[1].lower() or "(nessuna)"
                    cur = stats["per_ext"].setdefault(ext, [0, 0])
                    cur[0] += 1
                    cur[1] += size
                    if stats["files"] % 100 == 0 and progress_cb:
                        progress_cb(stats["files"])
            except OSError:
                continue

    walk(root)
    if stats["files"]:
        stats["avg"] = stats["bytes_total"] / stats["files"]
    return stats


# --------------------------------------------------------------------------
# Tema (chiaro / scuro): palette usata da _apply_theme
# --------------------------------------------------------------------------
THEME_LIGHT = {
    "window": "#f0f0f0",      # sfondo finestra (barre esterne)
    "panel": "#f5f6f8",       # pannelli (toolbar, anteprima, menubar)
    "page": "#ffffff",        # area dei file (albero)
    "tree_odd": "#f5f7fa",    # riga alternata dell'albero
    "tree_even": "#ffffff",
    "text": "#1c1f23",
    "muted": "#666666",
    "tooltip_bg": "#ffffe1",
    "tooltip_fg": "#1c1f23",
    "dnd": "#ffe08a",        # evidenza drag & drop
}
THEME_DARK = {
    "window": "#1e1f22",
    "panel": "#2b2d30",
    "page": "#26282b",
    "tree_odd": "#2e3136",
    "tree_even": "#26282b",
    "text": "#e6e8eb",
    "muted": "#9aa0a8",
    "tooltip_bg": "#3a3d42",
    "tooltip_fg": "#e6e8eb",
    "dnd": "#7a6310",
}


class _ScanCancelled(Exception):
    """Scansione automatica interrotta perché la vista è cambiata."""


class PreviewPanel(tk.Frame):
    """Pannello laterale di anteprima: immagini (con zoom), testo e PDF."""

    def __init__(self, master, app):
        c = app._colors
        super().__init__(master, bg=c["panel"], width=320)
        self.pack_propagate(False)
        self.app = app
        self._photo = None          # riferimento per evitare la garbage collection
        self._current_photo = None
        self._current_path = None
        self._img_w = self._img_h = 0
        self._zoom = 1.0            # scala visualizzata (1.0 = 100%)
        self._fit_zoom = 1.0
        self._build()

    def apply_theme(self, c):
        """Ricolorizza header/body/testo con la palette del tema attivo."""
        self.config(bg=c["panel"])
        self.header.config(bg=c["panel"])
        self.title_label.config(bg=c["panel"], fg=c["text"])
        self.meta_label.config(bg=c["panel"], fg=c["muted"])
        self.body.config(bg=c["page"])
        self.image_canvas.config(bg=c["page"])
        self.text_frame.config(bg=c["page"])
        self.text_widget.config(bg=c["page"], fg=c["text"])
        self.info_label.config(bg=c["page"], fg=c["muted"])
        self.footer.config(bg=c["panel"])

    def _build(self):
        c = self.app._colors
        self.header = tk.Frame(self, bg=c["panel"])
        self.header.pack(fill="x", padx=8, pady=(8, 4))
        self.title_label = tk.Label(self.header, text="Anteprima",
                                    font=("Segoe UI", 10, "bold"),
                                    bg=c["panel"], fg=c["text"], anchor="w",
                                    wraplength=300)
        self.title_label.pack(fill="x")
        self.meta_label = tk.Label(self.header, text="", font=("Segoe UI", 8),
                                   fg=c["muted"], bg=c["panel"], anchor="w",
                                   wraplength=300)
        self.meta_label.pack(fill="x")

        self.body = tk.Frame(self, bg=c["page"])
        self.body.pack(fill="both", expand=True, padx=8, pady=4)

        self.image_canvas = tk.Canvas(self.body, bg=c["page"], highlightthickness=0)
        self.image_canvas.bind("<MouseWheel>", self._on_image_wheel)
        self.image_canvas.bind("<Button-4>", self._on_image_wheel)
        self.image_canvas.bind("<Button-5>", self._on_image_wheel)
        self.image_canvas.bind("<Double-1>", self._on_image_double_click)
        self.image_canvas.bind("<Configure>", self._on_canvas_configure)
        self.text_frame = tk.Frame(self.body, bg=c["page"])
        self.text_widget = tk.Text(self.text_frame, wrap="word", font=("Consolas", 9),
                                   bg=c["page"], fg=c["text"], relief="flat",
                                   state="disabled")
        text_vsb = ttk.Scrollbar(self.text_frame, orient="vertical",
                                 command=self.text_widget.yview)
        self.text_widget.configure(yscrollcommand=text_vsb.set)
        self.text_widget.pack(side="left", fill="both", expand=True)
        text_vsb.pack(side="right", fill="y")
        self.info_label = tk.Label(self.body, text="", font=("Segoe UI", 9),
                                   fg=c["muted"], bg=c["page"], justify="left",
                                   anchor="n", wraplength=290, padx=8, pady=8)

        self.footer = tk.Frame(self, bg=c["panel"])
        self.footer.pack(fill="x", padx=8, pady=6)
        self.open_btn = ttk.Button(self.footer, text="Apri", command=self._open_current)
        self.open_btn.pack(side="left")

        self.show_placeholder()

    # ------------------------------------------------ visualizzazione
    def _show(self, widget):
        for w in (self.image_canvas, self.text_frame, self.info_label):
            w.pack_forget()
        widget.pack(fill="both", expand=True)

    def show_placeholder(self, text="Seleziona un file per vederne l'anteprima."):
        self._current_path = None
        self._photo = None
        self._current_photo = None
        self.title_label.config(text="Anteprima")
        self.meta_label.config(text="")
        self.open_btn.config(state="disabled")
        self.info_label.config(text=text)
        self._show(self.info_label)

    def show_path(self, path):
        self._current_path = path
        name = os.path.basename(path)
        self.title_label.config(text=name)
        if os.path.isdir(path):
            self._show_folder(path)
            return
        try:
            st = os.stat(path)
            meta = f"{format_size(st.st_size)}  ·  {format_date(st.st_mtime)}"
        except OSError:
            meta = ""
        self.meta_label.config(text=meta)
        ext = os.path.splitext(name)[1].lower()
        if ext in (".png", ".gif", ".ppm", ".pgm", ".pbm", ".jpg", ".jpeg",
                   ".bmp", ".tif", ".tiff", ".ico", ".webp"):
            self._show_image(path)
        elif ext == ".pdf":
            self._show_pdf(path)
        elif ext in (".docx", ".pptx", ".xlsx"):
            self._show_ooxml(path)
        elif ext in (".doc", ".ppt", ".xls"):
            self._show_legacy_office(path)
        elif ext in (".zip", ".tar", ".tgz", ".gz", ".bz2", ".xz", ".7z", ".rar"):
            self._show_archive(path)
        elif ext in AUDIO_EXTS or ext in VIDEO_EXTS:
            self._show_media(path)
        elif ext in TEXT_LIKE_EXTS or ext in PREVIEW_TEXT_EXTS or self._looks_like_text(path):
            self._show_text(path)
        elif ext in KNOWN_BINARY_EXTS:
            self._show_fallback("Nessuna anteprima disponibile per questo tipo di file.")
        else:
            self._show_fallback("Nessuna anteprima disponibile per questo tipo di file.")

    def _looks_like_text(self, path):
        return read_text_preview(path, 1024) is not None

    def _show_image(self, path):
        """Mostra l'immagine (PNG/GIF/PPM nativi; JPEG/BMP/TIFF/ICO via GDI+)."""
        photo = None
        try:
            if os.path.getsize(path) > 12 * 1024 * 1024:
                raise tk.TclError("troppo grande")
            photo = tk.PhotoImage(file=path)
        except (tk.TclError, OSError):
            photo = None
        if photo is None:
            dec = decode_image_rgb(path)
            if dec is None:
                self._show_fallback("Formato non supportato dall'anteprima\n"
                                    "(supportati: PNG, GIF, PPM, JPEG, BMP, TIFF, ICO, WebP).")
                return
            w, h, rgb = dec
            if w * h > PREVIEW_IMAGE_MAX_PIXELS:
                self._show_fallback("Immagine troppo grande per l'anteprima.")
                return
            photo = photo_from_rgb(w, h, rgb)
        self._photo = photo
        self._img_w, self._img_h = photo.width(), photo.height()
        if self._img_w < 1 or self._img_h < 1:
            self._show_fallback("Immagine con dimensioni non valide.")
            return
        if self._img_w * self._img_h > PREVIEW_IMAGE_MAX_PIXELS:
            self._show_fallback("Immagine troppo grande per l'anteprima.")
            return
        bw = max(self.image_canvas.winfo_width(), 150)
        bh = max(self.image_canvas.winfo_height(), 120)
        self._fit_zoom = min(1.0, bw / self._img_w, bh / self._img_h)
        self._zoom = self._initial_zoom()
        self._show(self.image_canvas)
        self._render_image()

    def _initial_zoom(self):
        """Zoom di apertura in base alla risoluzione scelta (Visualizza).

        Bassa: immagine adattata al pannello (tutta visibile).
        Media: 100%, dimensione reale.
        Alta: 200%, per vedere i dettagli (limitato dal limite di memoria).
        """
        res = int(self.app.preview_res.get())
        if res <= 1024:
            return self._fit_zoom
        if res <= 2048:
            return 1.0
        return 2.0

    def _render_image(self):
        """Ridisegna l'immagine al livello di zoom corrente (centrata)."""
        if self._photo is None:
            return
        if self._img_w < 1 or self._img_h < 1:
            return   # guardia: evita divisioni per zero con dimensioni vuote
        c = self.image_canvas
        c.delete("all")
        bw = max(c.winfo_width(), 150)
        bh = max(c.winfo_height(), 120)
        z = max(0.05, min(self._zoom, 64.0))
        # limite di memoria: dipende dalla risoluzione scelta nel menu
        # Visualizza > Risoluzione anteprima (default 4096 px)
        max_out = max(256, min(int(self.app.preview_res.get()) or 4096, 8192))
        if self._img_w * z > max_out or self._img_h * z > max_out:
            z = min(max_out / self._img_w, max_out / self._img_h)
        if self._img_w * self._img_h * z * z > PREVIEW_IMAGE_MAX_PIXELS:
            z = math.sqrt(PREVIEW_IMAGE_MAX_PIXELS / (self._img_w * self._img_h))
        d = 4
        n = max(1, min(int(z * d + 0.5), 8192))
        n = min(n, max(1, (max_out * d) // max(self._img_w, self._img_h)))
        photo = self._photo
        base = photo.subsample(d, d) if d > 1 else photo
        scaled = base.zoom(n, n) if n > 1 else base
        self._current_photo = scaled
        c.create_image(bw // 2, bh // 2, image=scaled)
        pct = round(n / d * 100)     # percentuale reale (n/d, dopo i limiti)
        hint = "rotella = zoom · doppio clic = 100%"
        self.meta_label.config(text=f"{self._img_w}×{self._img_h} px · {pct}% · {hint}")

    def _on_image_wheel(self, event):
        if self._photo is None:
            return
        delta = getattr(event, "delta", 0)
        if delta == 0:
            num = getattr(event, "num", 0)
            delta = 120 if num == 4 else (-120 if num == 5 else 0)
        if delta > 0:
            self._zoom = min(self._zoom * 1.25, 64.0)
        elif delta < 0:
            self._zoom = max(self._zoom / 1.25, 0.05)
        self._render_image()
        return "break"

    def _on_image_double_click(self, event=None):
        if self._photo is not None:
            self._zoom = 1.0
            self._render_image()

    def _on_canvas_configure(self, event):
        if self._photo is not None:
            self._render_image()

    def _show_text(self, path):
        text = read_text_preview(path)
        if text is None:
            self._show_fallback("File binario: nessuna anteprima di testo.")
            return
        self._photo = None
        self._current_photo = None
        self._show(self.text_frame)
        self.text_widget.config(state="normal")
        self.text_widget.delete("1.0", "end")
        self.text_widget.insert("1.0", text)
        try:
            if os.path.getsize(path) > PREVIEW_TEXT_LIMIT:
                self.text_widget.insert("end", "\n\n… (anteprima troncata)")
        except OSError:
            pass
        self.text_widget.config(state="disabled")

    def _show_pdf(self, path):
        text = extract_pdf_text(path)
        if not text.strip():
            self._show_fallback("PDF senza testo estraibile\n(scansione o protetto).")
            return
        self._show_text_parts(text)

    def _show_ooxml(self, path):
        """Mostra il testo estratto da DOCX/PPTX/XLSX (archivi ZIP di XML)."""
        text = extract_ooxml_text(path)
        if not text.strip():
            self._show_fallback("Documento Office senza testo estraibile.")
            return
        self._show_text_parts(text)

    def _show_legacy_office(self, path):
        """Mostra il testo best-effort da DOC/PPT/XLS binari (OLE2)."""
        text = extract_legacy_office_text(path)
        if not text.strip():
            self._show_fallback("Documento Office senza testo estraibile.")
            return
        self._show_text_parts(text)

    def _show_text_parts(self, text):
        """Riempie la finestra di testo dell'anteprima (PDF e documenti Office)."""
        self._photo = None
        self._current_photo = None
        self._show(self.text_frame)
        self.text_widget.config(state="normal")
        self.text_widget.delete("1.0", "end")
        self.text_widget.insert("1.0", text)
        self.text_widget.config(state="disabled")

    def _show_archive(self, path):
        """Mostra il contenuto di un archivio ZIP/TAR/7z/RAR (primi 200 elementi)."""
        entries = []
        total = 0
        ext = os.path.splitext(path)[1].lower()
        if ext == ".zip":
            try:
                with zipfile.ZipFile(path) as zf:
                    infos = zf.infolist()
            except (OSError, zipfile.BadZipFile) as exc:
                self._show_fallback(f"Archivio non leggibile:\n{exc}")
                return
            total = len(infos)
            for i in infos[:200]:
                icon = ICON_FOLDER if i.is_dir() else ICON_DOC
                entries.append(f"{icon} {i.filename}  {format_size(i.file_size)}")
        elif ext in (".tar", ".tgz", ".gz", ".bz2", ".xz"):
            try:
                with tarfile.open(path, "r:*") as tf:
                    members = tf.getmembers()
            except (OSError, tarfile.TarError) as exc:
                self._show_fallback(f"Archivio non leggibile:\n{exc}")
                return
            total = len(members)
            for m in members[:200]:
                icon = ICON_FOLDER if m.isdir() else ICON_DOC
                entries.append(f"{icon} {m.name}  {format_size(m.size)}")
        elif ext == ".7z":
            infos = self._list_7z(path)
            if infos is None:
                self._show_fallback("Archivio 7z non leggibile (serve 7-Zip o py7zr).")
                return
            total = len(infos)
            for name, size, folder in infos[:200]:
                icon = ICON_FOLDER if folder else ICON_DOC
                entries.append(f"{icon} {name}  {format_size(size)}")
        elif ext == ".rar":
            exe = _seven_zip_exe()
            if not exe:
                self._show_fallback("Anteprima RAR non disponibile: installa 7-Zip.\n"
                                    "(L'estrazione funziona anche con WinRAR.)")
                return
            infos = _seven_zip_list(exe, path)
            if not infos:
                self._show_fallback("Archivio RAR non leggibile.")
                return
            total = len(infos)
            for name, size, folder in infos[:200]:
                icon = ICON_FOLDER if folder else ICON_DOC
                entries.append(f"{icon} {name}  {format_size(size)}")
        else:
            self._show_fallback("Archivio non supportato.")
            return
        if not entries:
            self._show_fallback("Archivio vuoto.")
            return
        if total > 200:
            entries.append(f"… e altri {total - 200} elemento/i")
        self.meta_label.config(text=f"{total} elemento/i")
        self._show_text_parts("\n".join(entries))

    def _list_7z(self, path):
        """Elenca un archivio 7z: (nome, dimensione, cartella) via py7zr o 7-Zip."""
        try:
            import py7zr
        except ImportError:
            exe = _seven_zip_exe()
            return _seven_zip_list(exe, path) if exe else None
        try:
            with py7zr.SevenZipFile(path, "r") as z:
                return [(fi.filename, getattr(fi, "uncompressed", 0),
                         bool(getattr(fi, "is_directory", False)))
                        for fi in z.list()]
        except Exception:
            exe = _seven_zip_exe()
            return _seven_zip_list(exe, path) if exe else None

    def _show_media(self, path):
        """Anteprima dei metadati audio/video (via mutagen, se installato).

        mutagen è una libreria pura Python: compatibile con Nuitka senza flag
        speciali né DLL esterne. Se non è installata l'app continua a
        funzionare normalmente e mostra solo un avviso.
        """
        self._photo = None
        self._current_photo = None
        try:
            from mutagen import File as _mutagen_file
        except ImportError:
            self._show_fallback(
                "Anteprima audio/video non disponibile:\n"
                "la libreria 'mutagen' non è installata in questo Python.\n\n"
                f"Python in uso:\n{sys.executable}\n\n"
                "Per abilitarla esegui nel terminale:\n"
                "python -m pip install mutagen")
            return
        try:
            media = _mutagen_file(path, easy=True)
        except Exception:
            media = None
        if media is None or getattr(media, "info", None) is None:
            self._show_fallback(
                "Nessun metadato disponibile per questo file\n"
                "(formato non riconosciuto da mutagen).")
            return

        info = media.info
        cls = type(media).__name__.lower()
        is_video = (cls in ("matroska", "oggtheora") or "video" in cls
                    or getattr(info, "width", None)
                    or getattr(info, "height", None))
        fmt = MEDIA_FORMAT_LABELS.get(type(media).__name__, type(media).__name__)

        lines = [f"{'VIDEO' if is_video else 'AUDIO'}  ·  {fmt}", ""]

        length = getattr(info, "length", None)
        if length:
            lines.append(f"Durata:       {format_duration(length)}")
        bitrate = getattr(info, "bitrate", None)
        if bitrate:
            # nei video mutagen riporta il bitrate della traccia AUDIO
            suffix = " (audio)" if is_video else ""
            lines.append(f"Bitrate:      {bitrate / 1000:.0f} kbps{suffix}")
        sample_rate = getattr(info, "sample_rate", None)
        if sample_rate:
            lines.append(f"Frequenza:    {sample_rate / 1000:.1f} kHz")
        channels = getattr(info, "channels", None)
        if channels:
            ch = {1: "mono", 2: "stereo"}.get(channels, f"{channels} canali")
            lines.append(f"Canali:       {ch}")
        bits = getattr(info, "bits_per_sample", None)
        if bits:
            lines.append(f"Profondità:   {bits} bit")
        w, h = getattr(info, "width", None), getattr(info, "height", None)
        if w and h:
            lines.append(f"Risoluzione:  {w} × {h}")
        codec = getattr(info, "codec", None)
        if not codec and hasattr(info, "layer"):
            codec = f"MPEG layer {info.layer}"
        if codec:
            lines.append(f"Codec:        {codec}")

        tags = []
        for key, label in (("title", "Titolo"), ("artist", "Artista"),
                           ("album", "Album"), ("date", "Anno"),
                           ("genre", "Genere"), ("tracknumber", "Traccia")):
            val = media.get(key)
            if val:
                first = val[0] if isinstance(val, list) else val
                if first:
                    tags.append(f"{label}: {first}")
        if tags:
            lines.append("")
            lines.append("— Metadati —")
            lines.extend(tags)

        self._show_text_parts("\n".join(lines))

    def _show_fallback(self, msg):
        self._photo = None
        self._current_photo = None
        self.info_label.config(text=msg + "\n\nUsa il pulsante Apri per aprirlo.")
        self._show(self.info_label)

    def _show_folder(self, path):
        """Mostra il contenuto della cartella selezionata (come se fosse aperta)."""
        self._photo = None
        self._current_photo = None
        self.open_btn.config(state="normal")
        try:
            with os.scandir(path) as it:
                raw = list(it)
        except OSError:
            self._show_fallback("Accesso negato alla cartella.")
            return
        # raccoglie (nome, is_dir, stat) con try/except per elemento PRIMA di
        # ordinare: un singolo symlink rotto non deve far fallire la lista
        entries = []
        for e in raw:
            name = e.name
            if not self.app.show_hidden and name.startswith("."):
                continue
            try:
                is_dir = e.is_dir(follow_symlinks=True)
                st = e.stat(follow_symlinks=True)
            except OSError:
                is_dir = False
                st = None
            entries.append((not is_dir, name.lower(), is_dir, st, name))
        entries.sort(key=lambda t: (t[0], t[1]))
        # limite prudente: cartelle enormi non devono bloccare l'anteprima
        shown = entries[:200]
        lines = []
        for _k, _n, is_dir, st, name in shown:
            icon = ICON_FOLDER if is_dir else ICON_EXTS.get(
                os.path.splitext(name)[1].lower(), ICON_MISC)
            # nome + data (solo giorno, senza ora) su una riga unica
            data = format_date(st.st_mtime).split(" ")[0] if st is not None else ""
            lines.append(f"{icon} {name}  {data}".rstrip())
        if len(entries) > 200:
            lines.append(f"… e altri {len(entries) - 200} elemento/i")
        if not lines:
            lines.append("(cartella vuota)")
        self.meta_label.config(text=f"{len(entries)} elemento/i")
        self._show(self.text_frame)
        self.text_widget.config(state="normal")
        self.text_widget.delete("1.0", "end")
        self.text_widget.insert("1.0", "\n".join(lines))
        self.text_widget.config(state="disabled")

    def _open_current(self):
        if self._current_path and os.path.exists(self._current_path):
            self.app._open_file(self._current_path)


# --------------------------------------------------------------------------
# Applicazione principale
# --------------------------------------------------------------------------
class FileExplorer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Esplora File")
        self.geometry("1100x680")
        self.minsize(820, 480)

        # Stato globale
        self.tabs = []
        self._closed_tabs = []   # pile delle schede chiuse (index, path) per Ctrl+Maiusc+T
        self.clipboard = {"mode": None, "paths": []}   # "copy" | "cut"
        _settings = load_settings()
        self.show_hidden = bool(_settings.get("show_hidden", False))
        self.show_details = bool(_settings.get("show_details", True))
        self.drag = None
        self._search_after = None
        self._search_token = 0
        self._search_tab = None
        self._search_queue = queue.Queue()
        self._search_thread = None
        self.preview_var = tk.BooleanVar(
            value=bool(_settings.get("preview_panel", True)))   # pannello anteprima
        # ripristina la geometria della finestra salvata nell'ultima sessione
        geo = _settings.get("window_geometry")
        if isinstance(geo, str) and geo:
            try:
                self.geometry(geo)
            except tk.TclError:
                pass
        self.dark_var = tk.BooleanVar(value=bool(_settings.get("dark_theme", False)))
        self._colors = THEME_DARK if self.dark_var.get() else THEME_LIGHT
        try:
            _res = int(_settings.get("preview_resolution", 4096))
        except (TypeError, ValueError):
            _res = 4096
        # risoluzione massima di rendering dell'anteprima immagini
        self.preview_res = tk.IntVar(value=max(256, min(_res, 8192)))
        self._preview_after = None
        self.size_cache = FolderSizeCache()            # dimensioni cartelle
        self._size_token = 0                           # per scartare scan vecchie
        self._dup_token = 0                            # per scartare scan duplicati
        self._tooltip = None                           # tooltip con dettagli file
        self._tooltip_after = None                     # debounce del tooltip
        self._tooltip_row = None                       # riga sotto il mouse
        # dimensioni automatiche delle cartelle visibili (una alla volta)
        self._auto_enabled = True
        self._auto_ready = False
        self._auto_queue = []          # cartelle in attesa di scansione
        self._auto_busy = False
        self._auto_current = None      # cartella in scansione in questo momento
        self._auto_cancel = False      # interrompe la scansione se la vista cambia
        self._copy_move_busy = False   # copia/spostamento in corso (anti sovrapposizione)
        self._slideshow = None         # stato della presentazione immagini (slideshow)
        self._free_space = ""           # spazio libero sull'unità corrente (barra di stato)
        self._ext_filter = None          # filtro rapido per estensione (None = tutti)
        # preferenze della sessione precedente (colonne e ordinamento)
        _saved_cols = _settings.get("column_widths") or {}
        self._saved_columns = _saved_cols if isinstance(_saved_cols, dict) else {}
        self._saved_sort = (_settings.get("sort_col", "name"),
                            bool(_settings.get("sort_desc", False)))
        # cartelle preferite (lista di percorsi persistita in settings.json)
        _favs = _settings.get("favorites") or []
        self.favorites = [p for p in _favs if isinstance(p, str)]
        # cartelle recenti (menu Navigazione, persistite in settings.json)
        _recent = _settings.get("recent_folders") or []
        self.recent_folders = [p for p in _recent if isinstance(p, str)]

        self._setup_style()
        self._build_ui()
        self._populate_sidebar()
        self._restore_tabs(_settings)
        self._on_tab_changed()
        self._apply_theme()   # tema salvato (chiaro o scuro) su tutti i widget
        if not self.preview_var.get():
            self.paned.forget(self.preview_panel)   # anteprima disattivata all'avvio
        # attiva le dimensioni automatiche subito dopo l'avvio
        self._auto_start_id = self.after(150, self._auto_start)

    # ------------------------------------------- delega alla scheda attiva
    @property
    def active_tab(self):
        if not self.tabs:
            return None
        idx = self.tabbar.selected
        if idx < 0 or idx >= len(self.tabs):
            return self.tabs[0]
        return self.tabs[idx]

    @property
    def tree(self):
        return self.active_tab.tree

    @property
    def current_path(self):
        return self.active_tab.current_path

    @current_path.setter
    def current_path(self, value):
        self.active_tab.current_path = value

    @property
    def history(self):
        return self.active_tab.history

    @history.setter
    def history(self, value):
        self.active_tab.history = value

    @property
    def history_pos(self):
        return self.active_tab.history_pos

    @history_pos.setter
    def history_pos(self, value):
        self.active_tab.history_pos = value

    # ------------------------------------------------------------------ UI
    def _setup_style(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Toolbar.TFrame", padding=(6, 6))
        style.configure("Tool.TButton", padding=(6, 3))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))

    def _apply_theme(self):
        """Applica la palette (chiara o scura) a stili e widget principali.

        ttk permette di configurare gli stili; i widget tk (frame, albero,
        anteprima, barra laterale) vengono riconfigurati con i colori della
        palette. Chiamato all'avvio e quando si attiva/disattiva il tema.
        """
        c = self._colors
        style = ttk.Style(self)
        style.configure("Toolbar.TFrame", background=c["panel"])
        style.configure("Treeview", background=c["page"],
                        fieldbackground=c["page"], foreground=c["text"])
        style.configure("Treeview.Heading", background=c["panel"],
                        foreground=c["text"])
        style.configure("TFrame", background=c["panel"])
        style.configure("TLabel", background=c["panel"], foreground=c["text"])
        style.configure("Status.TLabel", background=c["panel"],
                        foreground=c["muted"])
        style.configure("TEntry", fieldbackground=c["page"],
                        foreground=c["text"])
        style.configure("TSeparator", background=c["panel"])
        self.configure(bg=c["window"])
        # schede esistenti
        for tab in self.tabs:
            tab.frame.config(bg=c["page"])
            tab.tree.tag_configure("odd", background=c["tree_odd"])
            tab.tree.tag_configure("even", background=c["tree_even"])
            tab.tree.tag_configure("dnd_target", background=c["dnd"])
        # barra laterale
        self.sidebar.tag_configure("dnd_target", background=c["dnd"])
        # breadcrumb (è un tk.Frame: va ricolorato a mano)
        self.breadcrumb.config(bg=c["panel"])
        # anteprima
        self.preview_panel.apply_theme(c)

    def toggle_dark_theme(self):
        """Attiva/disattiva il tema scuro e salva la preferenza."""
        self._colors = THEME_DARK if self.dark_var.get() else THEME_LIGHT
        self._apply_theme()
        s = load_settings()
        s["dark_theme"] = bool(self.dark_var.get())
        save_settings(s)
        self.refresh()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # --- Barra degli strumenti
        toolbar = ttk.Frame(self, style="Toolbar.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(5, weight=1)

        self.btn_back = ttk.Button(toolbar, text="\u2190", width=3,
                                   command=self.go_back, state="disabled")
        self.btn_forward = ttk.Button(toolbar, text="\u2192", width=3,
                                      command=self.go_forward, state="disabled")
        btn_up = ttk.Button(toolbar, text="\u2191", width=3, command=self.go_up)
        btn_refresh = ttk.Button(toolbar, text="\u21BB", width=3, command=self.refresh)
        btn_home = ttk.Button(toolbar, text="\U0001F3E0", width=3,
                              command=lambda: self.navigate_to(os.path.expanduser("~"), record=True))
        btn_new_tab = ttk.Button(toolbar, text="+", width=3, command=self.new_tab)
        self.btn_back.grid(row=0, column=0, padx=(0, 2))
        self.btn_forward.grid(row=0, column=1, padx=2)
        btn_up.grid(row=0, column=2, padx=2)
        btn_home.grid(row=0, column=3, padx=(2, 8))
        btn_new_tab.grid(row=0, column=4, padx=(2, 2))

        self.address_var = tk.StringVar()
        # Breadcrumb cliccabile (vista predefinita) + ingresso manuale del
        # percorso (tasto ✎). Il breadcrumb è un Frame di bottoni ricostruito
        # a ogni navigazione; la Entry è mostrata solo in modalità modifica.
        self.breadcrumb = tk.Frame(toolbar, bg=self._colors["panel"])
        self.breadcrumb.grid(row=0, column=5, sticky="ew", padx=(0, 2))
        self.address_entry = ttk.Entry(toolbar, textvariable=self.address_var, font=("Segoe UI", 10))
        self.address_entry.bind("<Return>", lambda e: self._go_address())
        self.address_entry.bind("<Escape>", lambda e: self._show_address_breadcrumb())
        self._addr_mode = "breadcrumb"     # "breadcrumb" | "entry"

        self.addr_btn = ttk.Button(toolbar, text="✎", width=3,
                                   command=self._toggle_address_mode)
        self.addr_btn.grid(row=0, column=6, padx=(0, 2))
        btn_refresh.grid(row=0, column=7, padx=2)

        # --- Filtro di ricerca (ricorsivo)
        ttk.Label(toolbar, text="\U0001F50D").grid(row=0, column=8, padx=(8, 2))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self._on_filter_changed)
        self.filter_entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=16,
                                      font=("Segoe UI", 10))
        self.filter_entry.grid(row=0, column=9, padx=(0, 4))
        btn_clear = ttk.Button(toolbar, text="\u2715", width=3,
                               command=lambda: self.filter_var.set(""))
        btn_clear.grid(row=0, column=10)

        # --- ricerca nel contenuto (grep) invece che nel nome
        self.search_content_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(toolbar, text="Contenuto",
                        variable=self.search_content_var,
                        command=self._on_content_toggle).grid(row=0, column=11, padx=(4, 0))

        # --- ricerca con espressione regolare (nel nome o nel contenuto)
        self.search_regex_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(toolbar, text="Regex",
                        variable=self.search_regex_var,
                        command=self._on_content_toggle).grid(row=0, column=12, padx=(4, 0))

        # --- Anteprima (pannello laterale)
        ttk.Button(toolbar, text="\U0001F441", width=3,
                   command=lambda: (self.preview_var.set(not self.preview_var.get()),
                                    self.toggle_preview())).grid(row=0, column=13, padx=(6, 0))

        # --- Filtro rapido per estensione (elenco a tendina)
        self.ext_filter_var = tk.StringVar(value="Tutti i tipi")
        self.ext_combobox = ttk.Combobox(toolbar, textvariable=self.ext_filter_var,
                                         state="readonly", width=13)
        self.ext_combobox.grid(row=0, column=14, padx=(8, 0))
        self.ext_combobox.bind("<<ComboboxSelected>>", self._on_ext_filter)

        # --- Menu
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nuova scheda", accelerator="Ctrl+T", command=self.new_tab)
        file_menu.add_command(label="Chiudi scheda", accelerator="Ctrl+W", command=self.close_tab)
        file_menu.add_command(label="Riapri scheda chiusa", accelerator="Ctrl+Maiusc+T",
                              command=self.reopen_closed_tab)
        file_menu.add_command(label="Duplica scheda", command=self.duplicate_tab)
        file_menu.add_command(label="Nuova finestra", command=self.new_window)
        file_menu.add_separator()
        file_menu.add_command(label="Nuova cartella", accelerator="Ctrl+Maiusc+N",
                              command=self.new_folder)
        file_menu.add_command(label="Crea collegamento", command=self.create_shortcut_selected)
        file_menu.add_separator()
        file_menu.add_command(label="Rinomina", accelerator="F2", command=self.rename_selected)
        file_menu.add_command(label="Elimina definitivamente", accelerator="Canc",
                              command=self.delete_selected)
        if IS_WINDOWS:
            file_menu.add_command(label="Svuota cestino",
                                  command=self.empty_recycle_bin_action)
        file_menu.add_separator()
        file_menu.add_command(label="Proprietà", accelerator="Alt+Invio", command=self.show_properties)
        file_menu.add_command(label="Esporta elenco (CSV)…", command=self.export_listing_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Copia", accelerator="Ctrl+C", command=self.copy_selected)
        edit_menu.add_command(label="Taglia", accelerator="Ctrl+X", command=self.cut_selected)
        edit_menu.add_command(label="Incolla", accelerator="Ctrl+V", command=self.paste_clipboard)
        edit_menu.add_command(label="Duplica", command=self.duplicate_selected)
        edit_menu.add_separator()
        edit_menu.add_command(label="Seleziona tutto", accelerator="Ctrl+A",
                              command=lambda: self.tree.selection_set(self.tree.get_children()))
        edit_menu.add_command(label="Seleziona stesso tipo", accelerator="Ctrl+Maiusc+A",
                              command=self.select_same_extension)
        edit_menu.add_command(label="Inverti selezione", accelerator="Ctrl+I",
                              command=self.invert_selection)
        edit_menu.add_command(label="Seleziona per modello…", accelerator="Ctrl+Maiusc+M",
                              command=self.select_by_pattern)
        edit_menu.add_command(label="Seleziona per dimensione…",
                              command=self.select_by_size)
        edit_menu.add_command(label="Seleziona per data…",
                              command=self.select_by_age)
        edit_menu.add_command(label="Seleziona per regex…",
                              command=self.select_by_regex)
        menubar.add_cascade(label="Modifica", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Aggiorna", accelerator="F5", command=self.refresh)
        self.show_hidden_var = tk.BooleanVar(value=self.show_hidden)
        view_menu.add_checkbutton(label="Mostra file nascosti", variable=self.show_hidden_var,
                                  command=lambda: (setattr(self, "show_hidden", self.show_hidden_var.get()),
                                                   self.refresh()))
        self.details_var = tk.BooleanVar(value=self.show_details)
        view_menu.add_checkbutton(label="Visualizza dettagli", variable=self.details_var,
                                  command=self.toggle_details)
        view_menu.add_separator()
        view_menu.add_checkbutton(label="Tema scuro", variable=self.dark_var,
                                  command=self.toggle_dark_theme)
        view_menu.add_separator()
        view_menu.add_checkbutton(label="Pannello anteprima", variable=self.preview_var,
                                  command=self.toggle_preview)
        res_menu = tk.Menu(view_menu, tearoff=0)
        res_menu.add_radiobutton(label="Bassa — adatta (1024 px)", variable=self.preview_res,
                                 value=1024, command=self.set_preview_resolution)
        res_menu.add_radiobutton(label="Media — 100% (2048 px)", variable=self.preview_res,
                                 value=2048, command=self.set_preview_resolution)
        res_menu.add_radiobutton(label="Alta — 200% (4096 px)", variable=self.preview_res,
                                 value=4096, command=self.set_preview_resolution)
        view_menu.add_cascade(label="Risoluzione anteprima", menu=res_menu)
        view_menu.add_separator()
        view_menu.add_command(label="Presentazione (slideshow)",
                              command=self.start_slideshow)
        menubar.add_cascade(label="Visualizza", menu=view_menu)

        nav_menu = tk.Menu(menubar, tearoff=0)
        nav_menu.add_command(label="Nuova scheda", accelerator="Ctrl+T", command=self.new_tab)
        nav_menu.add_command(label="Scheda successiva", accelerator="Ctrl+Tab",
                             command=lambda: self._next_tab(1))
        nav_menu.add_command(label="Scheda precedente", accelerator="Ctrl+Maiusc+Tab",
                             command=lambda: self._next_tab(-1))
        nav_menu.add_separator()
        nav_menu.add_command(label="Indietro", accelerator="Alt+\u2190", command=self.go_back)
        nav_menu.add_command(label="Avanti", accelerator="Alt+\u2192", command=self.go_forward)
        nav_menu.add_command(label="Cartella superiore", accelerator="Alt+\u2191", command=self.go_up)
        nav_menu.add_separator()
        nav_menu.add_command(label=REVEAL_LABEL, command=self.reveal_in_file_manager)
        nav_menu.add_command(label="Confronta cartelle…", command=self.compare_folders)
        self._recent_menu = tk.Menu(nav_menu, tearoff=0)
        nav_menu.add_cascade(label="Cartelle recenti", menu=self._recent_menu)
        nav_menu.config(postcommand=self._refresh_recent_menu)
        menubar.add_cascade(label="Navigazione", menu=nav_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Informazioni", command=self.show_about)
        menubar.add_cascade(label="Aiuto", menu=help_menu)
        self.config(menu=menubar)

        # --- Area principale: barra laterale + schede (stile browser)
        self.paned = ttk.Panedwindow(self, orient="horizontal")
        self.paned.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))

        side_frame = ttk.Frame(self.paned)
        self.sidebar = ttk.Treeview(side_frame, show="tree", selectmode="browse")
        sb_side = ttk.Scrollbar(side_frame, orient="vertical", command=self.sidebar.yview)
        self.sidebar.configure(yscrollcommand=sb_side.set)
        self.sidebar.pack(side="left", fill="both", expand=True)
        sb_side.pack(side="right", fill="y")
        self.sidebar.bind("<<TreeviewSelect>>", self._on_sidebar_select)
        self.sidebar.bind("<Button-3>", self._on_sidebar_menu)
        self.sidebar.tag_configure("dnd_target", background="#ffe08a")
        self.paned.add(side_frame, weight=0)

        tab_frame = ttk.Frame(self.paned)
        tab_frame.columnconfigure(0, weight=1)
        tab_frame.rowconfigure(1, weight=1)

        self.tabbar = TabBar(tab_frame,
                             on_select=self._on_tab_changed,
                             on_close=self.close_tab,
                             on_new=self.new_tab,
                             on_context=self._on_tab_context)
        self.tabbar.grid(row=0, column=0, sticky="ew")

        self.pages = tk.Frame(tab_frame, bg="#ffffff")
        self.pages.grid(row=1, column=0, sticky="nsew")
        self.pages.columnconfigure(0, weight=1)
        self.pages.rowconfigure(0, weight=1)
        self.paned.add(tab_frame, weight=1)

        # --- Pannello di anteprima (a destra)
        self.preview_panel = PreviewPanel(self.paned, self)
        self.paned.add(self.preview_panel, weight=0)

        # --- Barra di stato
        status = ttk.Frame(self, relief="sunken")
        status.grid(row=2, column=0, sticky="ew")
        self.status_left = ttk.Label(status, text="", style="Status.TLabel", padding=(6, 2))
        self.status_left.pack(side="left")
        self.status_right = ttk.Label(status, text="", style="Status.TLabel", padding=(6, 2))
        self.status_right.pack(side="right")
        # barra di avanzamento per copia/spostamento (nascosta finché non serve)
        self.progress_bar = ttk.Progressbar(status, mode="determinate", length=220)

        # scorciatoie globali per le schede
        self.bind("<Control-t>", lambda e: (self.new_tab(), "break")[1])
        self.bind("<Control-w>", lambda e: (self.close_tab(), "break")[1])
        self.bind("<Control-Shift-T>", lambda e: (self.reopen_closed_tab(), "break")[1])
        self.bind("<Control-Tab>", lambda e: (self._next_tab(1), "break")[1])
        self.bind("<Control-Shift-Tab>", lambda e: (self._next_tab(-1), "break")[1])
        self.bind("<Control-i>", lambda e: (self.invert_selection(), "break")[1])
        self.bind("<Control-Shift-A>", lambda e: (self.select_same_extension(), "break")[1])
        self.bind("<Control-Shift-M>", lambda e: (self.select_by_pattern(), "break")[1])
        self.bind("<Control-f>", lambda e: (self._focus_search(), "break")[1])
        self.bind("<Control-l>", lambda e: (self._focus_address(), "break")[1])
        self.bind("<Alt-p>", lambda e: (self.preview_var.set(not self.preview_var.get()),
                                         self.toggle_preview(), "break")[2])
        # tasti laterali del mouse da gaming: XButton1 (posteriore) = Indietro,
        # XButton2 (anteriore) = Avanti, come le frecce in alto a sinistra.
        # Su Windows la rotella genera MouseWheel, quindi i pulsanti 4 e 5
        # sono proprio XButton1/XButton2 (la documentazione Tk ammette solo
        # bottoni 1-5). bind_all: funzionano ovunque (albero, sidebar, ...)
        if IS_WINDOWS:
            self.bind_all("<Button-4>", lambda e: (self.go_back(), "break")[1])
            self.bind_all("<Button-5>", lambda e: (self.go_forward(), "break")[1])

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------- schede
    def _restore_tabs(self, settings):
        """Riapre le schede dell'ultima sessione (o una sola sulla home)."""
        raw = settings.get("tabs") or []
        if isinstance(raw, list):
            seen = set()
            for p in raw:
                if not isinstance(p, str):
                    continue
                target = os.path.expanduser(p)
                if not os.path.isdir(target):
                    continue
                key = os.path.normcase(os.path.normpath(target))
                if key in seen:
                    continue
                seen.add(key)
                self.new_tab(path=target)
        if not self.tabs:
            self.new_tab()
        idx = settings.get("active_tab_index")
        if isinstance(idx, int) and 0 <= idx < len(self.tabs):
            self.tabbar.select(idx)

    def new_tab(self, path=None, index=None):
        """Crea una scheda; se `index` è dato la inserisce in quella posizione
        (usato da 'riapri scheda chiusa' per ripristinare l'ordine originale)."""
        tab = TabView(self)
        tab.frame.grid(row=0, column=0, sticky="nsew")
        target = os.path.normpath(os.path.expanduser(
            str(path if path is not None else os.path.expanduser("~"))))
        if not os.path.isdir(target):
            target = os.path.expanduser("~")
        tab.current_path = target
        tab.history = [target]
        tab.history_pos = 0
        # applica le preferenze salvate: larghezze colonne e ordinamento
        for col, w in self._saved_columns.items():
            if col in ("name", "type", "size", "date", "path") and isinstance(w, (int, float)):
                try:
                    tab.tree.column(col, width=int(w))
                except tk.TclError:
                    pass
        if self._saved_sort[0] in ("name", "type", "size", "date", "path"):
            tab.sort_col = self._saved_sort[0]
        tab.sort_desc = bool(self._saved_sort[1])
        tab._update_sort_headers()
        # mantiene sincronizzate la lista dei TabView e quella della TabBar
        if index is None:
            self.tabs.append(tab)
            idx = self.tabbar.add_tab("Nuova scheda")
        else:
            index = max(0, min(int(index), len(self.tabs)))
            self.tabs.insert(index, tab)
            idx = self.tabbar.insert_tab(index, "Nuova scheda")
        self.tabbar.select(idx)          # attiva la scheda (aggiorna la UI)
        return tab

    def duplicate_tab(self):
        self.new_tab(path=self.current_path)

    def empty_recycle_bin_action(self):
        """Svuota il Cestino (con conferma esplicita). Solo Windows."""
        if not IS_WINDOWS:
            messagebox.showinfo("Esplora File",
                                "Il Cestino è disponibile solo su Windows.")
            return
        if not messagebox.askyesno("Svuota cestino",
                                   "Svuotare completamente il Cestino?\n\n"
                                   "L'operazione è irreversibile.", parent=self):
            return
        if empty_recycle_bin():
            self.status_left.config(text="Cestino svuotato.")
        else:
            messagebox.showerror("Esplora File",
                                 "Impossibile svuotare il Cestino.")

    def new_window(self, path=None):
        """Apre una seconda finestra dell'app sulla cartella indicata (o quella
        corrente). Funziona anche da eseguibile compilato (Nuitka/PyInstaller)."""
        target = os.path.normpath(path or self.current_path)
        if getattr(sys, "frozen", False):
            # eseguibile autonomo: rilancia se stesso passando la cartella
            cmd = [sys.executable, target]
        else:
            cmd = [sys.executable, os.path.abspath(__file__), target]
        try:
            subprocess.Popen(cmd)
        except OSError as exc:
            messagebox.showerror("Esplora File",
                                 f"Impossibile aprire una nuova finestra:\n{exc}")

    def close_tab(self, index=None):
        if len(self.tabs) <= 1:
            return  # mantieni almeno una scheda
        if index is None:
            index = self.tabbar.selected
        if index < 0 or index >= len(self.tabs):
            return
        tab = self.tabs.pop(index)
        path = tab.current_path
        self.tabbar.remove_tab(index)
        tab.frame.destroy()
        # ricorda la scheda chiusa per 'riapri scheda chiusa' (Ctrl+Maiusc+T)
        self._closed_tabs.append((index, path))
        if len(self._closed_tabs) > 15:
            self._closed_tabs.pop(0)
        self._on_tab_changed()

    def reopen_closed_tab(self):
        """Riapre l'ultima scheda chiusa, nella posizione originale se possibile."""
        while self._closed_tabs:
            index, path = self._closed_tabs.pop()
            if os.path.isdir(path):
                # se nel frattempo le schede sono diminuite, limita l'indice
                self.new_tab(path=path, index=max(0, min(index, len(self.tabs))))
                return
            # la cartella non esiste più: salta e riprova con la precedente

    def close_other_tabs(self, keep_idx):
        for i in range(len(self.tabs) - 1, -1, -1):
            if i != keep_idx:
                tab = self.tabs.pop(i)
                self.tabbar.remove_tab(i)
                tab.frame.destroy()
        self.tabbar.select(0)
        self._on_tab_changed()

    def _next_tab(self, delta):
        if len(self.tabs) < 2:
            return
        nxt = (self.tabbar.selected + delta) % len(self.tabs)
        self.tabbar.select(nxt)

    def _on_tab_context(self, index, x_root, y_root):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Nuova scheda", command=self.new_tab)
        if 0 <= index < len(self.tabs):
            menu.add_separator()
            menu.add_command(label="Duplica scheda",
                             command=lambda: self.new_tab(path=self.tabs[index].current_path))
            menu.add_command(label="Chiudi scheda", command=lambda: self.close_tab(index))
            menu.add_command(label="Riapri scheda chiusa", command=self.reopen_closed_tab)
            menu.add_command(label="Chiudi le altre",
                             command=lambda: self.close_other_tabs(index))
        menu.tk_popup(x_root, y_root)
        menu.grab_release()

    # ------------------------------------------------------- breadcrumb
    def _update_breadcrumb(self, path=None):
        """Ricostruisce i segmenti cliccabili del percorso nella barra indirizzi."""
        if path is None:
            path = self.current_path
        for w in self.breadcrumb.winfo_children():
            w.destroy()
        segs = _split_path(path) or [(path, path)]
        for i, (label, full) in enumerate(segs):
            if i > 0:
                ttk.Label(self.breadcrumb, text="›").pack(side="left", padx=3)
            ttk.Button(self.breadcrumb, text=label,
                       command=lambda p=full: self.navigate_to(p, record=True)
                       ).pack(side="left")

    def _toggle_address_mode(self):
        """Alterna tra breadcrumb cliccabile e ingresso manuale del percorso."""
        if self._addr_mode == "breadcrumb":
            self._addr_mode = "entry"
            self.breadcrumb.grid_remove()
            self.address_entry.grid(row=0, column=5, sticky="ew", padx=(0, 2))
            self.addr_btn.config(text="Vai", command=self._go_address)
            self.address_entry.focus_set()
            self.address_entry.selection_range(0, "end")
            self.address_entry.icursor("end")
        else:
            self._show_address_breadcrumb()

    def _show_address_breadcrumb(self):
        """Torna alla vista breadcrumb (nascondendo l'ingresso testo)."""
        self._addr_mode = "breadcrumb"
        self.address_entry.grid_remove()
        self.breadcrumb.grid()
        self.addr_btn.config(text="✎", command=self._toggle_address_mode)
        self._update_breadcrumb()

    def _go_address(self):
        """Naviga al percorso digitato e torna alla vista breadcrumb."""
        self.navigate_to(self.address_var.get(), record=True)
        self._show_address_breadcrumb()

    def _focus_search(self):
        """Ctrl+F: porta il focus sul campo di ricerca (selezionandone il testo)."""
        self.filter_entry.focus_set()
        self.filter_entry.selection_range(0, "end")

    def _focus_address(self):
        """Ctrl+L: porta il focus sulla barra indirizzi (in modalità modifica)."""
        if self._addr_mode != "entry":
            self._toggle_address_mode()
        else:
            self.address_entry.focus_set()
            self.address_entry.selection_range(0, "end")

    def _update_tab_title(self):
        tab = self.active_tab
        if tab is None:
            return
        name = os.path.basename(self.current_path.rstrip("\\/")) or self.current_path
        self.tabbar.set_title(self.tabbar.selected, name)

    def _on_tab_changed(self, index=None):
        if not self.tabs:
            return
        tab = self.active_tab
        if tab is None:
            return
        tab.frame.tkraise()
        self._hide_tooltip()
        self.address_var.set(self.current_path)
        self._update_breadcrumb()
        self._update_free_space()
        self._update_nav_buttons()
        self._update_tab_title()
        self._reset_ext_filter()
        self.refresh(keep_selection=False)
        self._populate_ext_filter()
        self._schedule_preview()

    # ------------------------------------------------------- ricerca ricorsiva
    def _on_filter_changed(self, *args):
        if self._search_after is not None:
            self.after_cancel(self._search_after)
            self._search_after = None
        if not self.filter_var.get().strip():
            # annulla la ricerca in corso e torna alla lista normale
            self._search_token += 1
            tab = self.active_tab
            if tab is not None:
                tab.search_active = False
                tab._apply_columns()
                tab.refresh(keep_selection=False)
            return
        self._search_after = self.after(300, self._start_search)

    def _on_content_toggle(self):
        """Rilancia la ricerca (se il filtro non è vuoto) con la nuova modalità."""
        if self.filter_var.get().strip():
            self._on_filter_changed()

    def _on_ext_filter(self, event=None):
        """Applica il filtro rapido per estensione scelto nel menu a tendina."""
        val = self.ext_filter_var.get().strip().lower()
        self._ext_filter = None if val in ("", "tutti i tipi") else val
        self.refresh()

    def _reset_ext_filter(self):
        """Azzera il filtro per estensione (a ogni navigazione, come Explorer)."""
        self._ext_filter = None
        self.ext_filter_var.set("Tutti i tipi")

    def _populate_ext_filter(self):
        """Riempie il menu a tendina con le estensioni presenti nella cartella."""
        tab = self.active_tab
        exts = set()
        if tab is not None:
            for e in tab.entries:
                if not e[2]:
                    exts.add(os.path.splitext(e[1])[1].lower())
        self.ext_combobox["values"] = ["Tutti i tipi"] + sorted(x for x in exts if x)

    def _start_search(self):
        self._search_after = None
        tab = self.active_tab
        if tab is None:
            return
        token = self._search_token = self._search_token + 1
        self._search_tab = tab
        raw = self.filter_var.get().strip()
        regex_mode = self.search_regex_var.get()
        pattern = None
        if regex_mode:
            # compila la regex nel thread principale: un pattern non valido
            # viene segnalato subito senza avviare inutilmente la ricerca
            try:
                pattern = re.compile(raw, re.IGNORECASE)
            except re.error as exc:
                self.status_left.config(text=f"Regex non valida: {exc}")
                return
        tab.search_active = True
        tab.entries = []
        tab.tree.delete(*tab.tree.get_children())
        tab._apply_columns()
        content_mode = self.search_content_var.get()
        self.status_left.config(text="Ricerca nel contenuto…" if content_mode
                               else "Ricerca in corso\u2026")
        root = tab.current_path
        filt = raw.lower()
        show_hidden = self.show_hidden
        q = self._search_queue

        def match(s):
            """True se 's' soddisfa il filtro (regex oppure sottostringa)."""
            if pattern is not None:
                return pattern.search(s) is not None
            return filt in s.lower()

        def worker():
            results = []
            truncated = False
            try:
                for dirpath, dirnames, filenames in os.walk(root):
                    if not show_hidden:
                        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
                    rel = os.path.relpath(dirpath, root)
                    rel = "" if rel == "." else rel
                    if content_mode:
                        # grep: cerca il testo nel CONTENUTO dei file (solo file,
                        # saltando i binari e leggendo al massimo CONTENT_SEARCH_LIMIT byte)
                        for name in filenames:
                            if len(results) >= MAX_SEARCH_RESULTS:
                                truncated = True
                                break
                            p = os.path.join(dirpath, name)
                            text = read_text_preview(p, CONTENT_SEARCH_LIMIT)
                            if text is not None and match(text):
                                try:
                                    st = os.stat(p)
                                    results.append((p, name, False, st.st_size,
                                                    st.st_mtime, rel))
                                except OSError:
                                    results.append((p, name, False, 0, 0, rel))
                        if truncated:
                            break
                    else:
                        for name in filenames:
                            if len(results) >= MAX_SEARCH_RESULTS:
                                truncated = True
                                break
                            if match(name):
                                p = os.path.join(dirpath, name)
                                try:
                                    st = os.stat(p)
                                    results.append((p, name, False, st.st_size,
                                                    st.st_mtime, rel))
                                except OSError:
                                    results.append((p, name, False, 0, 0, rel))
                        for name in dirnames:
                            if len(results) >= MAX_SEARCH_RESULTS:
                                truncated = True
                                break
                            if match(name):
                                p = os.path.join(dirpath, name)
                                try:
                                    st = os.stat(p)
                                    results.append((p, name, True, 0, st.st_mtime, rel))
                                except OSError:
                                    results.append((p, name, True, 0, 0, rel))
                        if truncated:
                            break
            except Exception:
                pass
            q.put((token, results, truncated))

        self._search_thread = threading.Thread(target=worker, daemon=True)
        self._search_thread.start()
        self._poll_search()

    def _poll_search(self):
        # Svuota la coda: un risultato obsoleto (token diverso) va scartato,
        # ma NON deve impedire di leggere il risultato della ricerca corrente
        # ancora in coda (prima il codice tornava subito senza riprogrammare
        # il poll, lasciando il risultato vero mai visualizzato).
        while True:
            try:
                token, results, truncated = self._search_queue.get_nowait()
            except queue.Empty:
                break
            if token != self._search_token:
                continue  # risultato obsoleto: scartalo
            # usa la scheda per cui è partita la ricerca, non quella attiva ora
            tab = self._search_tab
            if tab is None or tab is not self.active_tab or not tab.search_active:
                continue
            tab.entries = results
            tab._render_tree()
            n = len(results)
            msg = f"{n} risultato/i in tutte le sottocartelle"
            if truncated:
                msg += f"  (limite {MAX_SEARCH_RESULTS} raggiunto)"
            self.status_left.config(text=msg)
        # continua a controllare finché il thread è vivo o ci sono risultati in coda
        if (self._search_thread is not None and self._search_thread.is_alive()) \
                or not self._search_queue.empty():
            self.after(80, self._poll_search)

    # ------------------------------------------------------- barra laterale
    def _populate_sidebar(self):
        self.sidebar.delete(*self.sidebar.get_children())
        self._sidebar_map = {}

        # --- Preferiti (cartelle scelte dall'utente) ---
        fav_root = self.sidebar.insert("", "end", text="\u2605 Preferiti", open=True)
        for p in self.favorites:
            if os.path.isdir(p):
                iid = f"fav:{os.path.normcase(p)}"
                self.sidebar.insert(fav_root, "end", iid=iid,
                                    text=f"\u2b50 {os.path.basename(p) or p}")
                self._sidebar_map[iid] = p

        home = Path.home()
        quick = [
            ("Desktop", home / "Desktop"),
            ("Documenti", home / "Documents"),
            ("Download", home / "Downloads"),
            ("Immagini", home / "Pictures"),
            ("Musica", home / "Music"),
            ("Video", home / "Videos"),
        ]
        quick = [(nome, p) for nome, p in quick if p.exists()]

        quick_root = self.sidebar.insert("", "end", text="\U00002601 Accesso rapido", open=True)
        for nome, p in quick:
            iid = f"quick:{nome}"
            self.sidebar.insert(quick_root, "end", iid=iid, text=f"{ICON_FOLDER} {nome}")
            self._sidebar_map[iid] = str(p)

        drives = [f"{c}:\\" for c in string.ascii_uppercase if os.path.exists(f"{c}:\\")]
        pc_root = self.sidebar.insert("", "end", text=f"{ICON_PC} Questo PC", open=True)
        for d in drives:
            iid = f"drive:{d[:1]}"
            self.sidebar.insert(pc_root, "end", iid=iid, text=f"{ICON_DRIVE} {d}")
            self._sidebar_map[iid] = d

    def _on_sidebar_select(self, event):
        sel = self.sidebar.selection()
        if sel and sel[0] in self._sidebar_map:
            self.navigate_to(self._sidebar_map[sel[0]], record=True)

    def _on_sidebar_menu(self, event):
        """Menu contestuale sulla barra laterale (rimuovi dai Preferiti)."""
        iid = self.sidebar.identify_row(event.y)
        if not iid or not iid.startswith("fav:"):
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Rimuovi dai Preferiti",
                         command=lambda: self.toggle_favorite(self._sidebar_map.get(iid)))
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    def _fav_label(self, path):
        """Etichetta del menu contestuale per aggiungere/rimuovere un Preferito."""
        norm = os.path.normcase(os.path.normpath(path))
        return ("Rimuovi dai Preferiti"
                if any(os.path.normcase(p) == norm for p in self.favorites)
                else "Aggiungi ai Preferiti")

    def toggle_favorite(self, path=None):
        """Aggiunge o rimuove una cartella dai Preferiti (persistente)."""
        if path is None:
            path = self.current_path
        path = os.path.normpath(str(path))
        if not os.path.isdir(path):
            messagebox.showerror("Esplora File", "Seleziona una cartella valida.")
            return
        norm = os.path.normcase(path)
        if any(os.path.normcase(p) == norm for p in self.favorites):
            self.favorites = [p for p in self.favorites if os.path.normcase(p) != norm]
        else:
            self.favorites.append(path)
        s = load_settings()
        s["favorites"] = self.favorites
        save_settings(s)
        self._populate_sidebar()

    # ------------------------------------------------------- navigazione
    def navigate_to(self, path, record=True):
        path = os.path.normpath(os.path.expandvars(os.path.expanduser(str(path).strip())))
        if not os.path.isdir(path):
            if os.path.exists(path):
                messagebox.showerror("Esplora File", "Il percorso non è una cartella.")
            else:
                messagebox.showerror("Esplora File", f"Percorso non trovato:\n{path}")
            self.address_var.set(self.current_path)
            self._update_breadcrumb()
            return
        if record:
            # evita duplicati in cronologia quando si torna nella stessa cartella
            if os.path.normcase(path) == os.path.normcase(self.current_path):
                self.address_var.set(path)
                self._update_breadcrumb()
                self.refresh(keep_selection=False)
                self._update_nav_buttons()
                return
            if self.history and self.history_pos < len(self.history) - 1:
                self.history = self.history[: self.history_pos + 1]
            self.history.append(path)
            self.history_pos = len(self.history) - 1
            self._add_recent_folder(path)
        self.current_path = path
        self.address_var.set(path)
        self._update_breadcrumb()
        self._update_free_space()
        self._update_tab_title()
        self._reset_ext_filter()
        self.refresh(keep_selection=False)
        self._populate_ext_filter()
        self._update_nav_buttons()

    def _add_recent_folder(self, path):
        """Aggiunge path in cima alle cartelle recenti (max 15, senza duplicati)."""
        norm = os.path.normcase(os.path.normpath(path))
        self.recent_folders = [p for p in self.recent_folders
                               if os.path.normcase(os.path.normpath(p)) != norm]
        self.recent_folders.insert(0, path)
        self.recent_folders = self.recent_folders[:15]
        s = load_settings()
        s["recent_folders"] = self.recent_folders
        save_settings(s)

    def _refresh_recent_menu(self):
        """Ricostruisce il sottomenu 'Cartelle recenti' all'apertura."""
        m = self._recent_menu
        m.delete(0, "end")
        if not self.recent_folders:
            m.add_command(label="Nessuna cartella recente", state="disabled")
            return
        for p in self.recent_folders:
            m.add_command(label=p,
                          command=lambda path=p: self.navigate_to(path, record=True))
        m.add_separator()
        m.add_command(label="Svuota elenco", command=self._clear_recent_folders)

    def _clear_recent_folders(self):
        self.recent_folders = []
        s = load_settings()
        s["recent_folders"] = []
        save_settings(s)

    def go_back(self):
        if self.history_pos > 0:
            self.history_pos -= 1
            self.navigate_to(self.history[self.history_pos], record=False)

    def go_forward(self):
        if self.history_pos < len(self.history) - 1:
            self.history_pos += 1
            self.navigate_to(self.history[self.history_pos], record=False)

    def go_up(self):
        parent = os.path.dirname(self.current_path.rstrip("\\/"))
        if parent and os.path.isdir(parent):
            self.navigate_to(parent, record=True)

    def _update_nav_buttons(self):
        self.btn_back.config(state="normal" if self.history_pos > 0 else "disabled")
        self.btn_forward.config(state="normal" if self.history_pos < len(self.history) - 1 else "disabled")

    def refresh(self, keep_selection=True):
        """Ricarica la scheda attiva (o avvia la ricerca se il filtro è attivo)."""
        if self.active_tab is not None:
            self.active_tab.refresh(keep_selection)

    def toggle_details(self):
        self.show_details = self.details_var.get()
        for tab in self.tabs:
            tab._apply_columns()

    # ------------------------------------------------------- anteprima
    def toggle_preview(self):
        """Mostra / nasconde il pannello di anteprima a destra."""
        if self.preview_var.get():
            if str(self.preview_panel) not in self.paned.panes():
                self.paned.add(self.preview_panel, weight=0)
            self._schedule_preview()
        else:
            if str(self.preview_panel) in self.paned.panes():
                self.paned.forget(self.preview_panel)
            if self._preview_after is not None:
                self.after_cancel(self._preview_after)
                self._preview_after = None

    def _on_tree_select(self):
        self._update_status()
        self._schedule_preview()

    # ------------------------------------------------------- tooltip dettagli
    def _on_tree_motion(self, event):
        """Al passaggio del mouse su una riga avvia il debounce del tooltip."""
        if self.drag:
            self._hide_tooltip()   # niente tooltip mentre trascini
            return
        try:
            row = self.tree.identify_row(event.y)
        except tk.TclError:
            return
        if not row:
            self._hide_tooltip()
            return
        if row != self._tooltip_row:
            # riga diversa: chiudi l'eventuale tooltip e riavvia il debounce
            if self._tooltip is not None:
                self._hide_tooltip()
            self._tooltip_row = row
            if self._tooltip_after is not None:
                self.after_cancel(self._tooltip_after)
            self._tooltip_after = self.after(450, self._show_tooltip)

    def _on_tree_leave(self, event):
        self._hide_tooltip()

    def _show_tooltip(self):
        """Mostra un riquadro con i dettagli del file sotto il mouse."""
        self._tooltip_after = None
        path = self._tooltip_row
        if not path:
            return
        try:
            if not self.tree.exists(path):
                return
        except tk.TclError:
            return
        try:
            st = os.stat(path)
            is_dir = os.path.isdir(path)
        except OSError:
            return
        name = os.path.basename(path)
        if is_dir:
            tipo = "Cartella di file"
            size_txt = "—"
        else:
            entry = (path, name, False, st.st_size, st.st_mtime, "")
            tipo = self.active_tab._type_label(entry)
            size_txt = format_size(st.st_size)
        text = (f"{name}\n"
                f"Tipo: {tipo}\n"
                f"Dimensione: {size_txt}\n"
                f"Modificato: {format_date(st.st_mtime)}\n"
                f"Percorso: {path}")
        tip = tk.Toplevel(self)
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        c = self._colors
        tk.Label(tip, text=text, justify="left", bg=c["tooltip_bg"], fg=c["tooltip_fg"],
                 relief="solid", borderwidth=1, font=("Segoe UI", 9),
                 padx=8, pady=6).pack()
        self._tooltip = tip
        try:
            x_root = self.tree.winfo_pointerx()
            y_root = self.tree.winfo_pointery()
        except tk.TclError:
            x_root = y_root = 0
        tip.update_idletasks()
        w, h = tip.winfo_reqwidth(), tip.winfo_reqheight()
        x = min(x_root + 14, self.winfo_screenwidth() - w - 8)
        y = min(y_root + 18, self.winfo_screenheight() - h - 8)
        tip.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _hide_tooltip(self):
        if self._tooltip_after is not None:
            self.after_cancel(self._tooltip_after)
            self._tooltip_after = None
        if self._tooltip is not None:
            try:
                self._tooltip.destroy()
            except tk.TclError:
                pass
            self._tooltip = None
        self._tooltip_row = None

    def _schedule_preview(self):
        """Ricarica l'anteprima con un piccolo ritardo (debounce)."""
        if not self.preview_var.get():
            return
        if self._preview_after is not None:
            self.after_cancel(self._preview_after)
        self._preview_after = self.after(250, self._load_preview)

    def _load_preview(self):
        self._preview_after = None
        if not self.preview_var.get():
            return
        tab = self.active_tab
        if tab is None:
            return
        try:
            if not tab.frame.winfo_exists():
                return
            sel = tab.tree.selection()
        except tk.TclError:
            return
        if len(sel) == 1:
            # se il file è già mostrato non ricaricare: mantiene lo zoom
            if self.preview_panel._current_path == sel[0]:
                return
            self.preview_panel.show_path(sel[0])
        else:
            self.preview_panel.show_placeholder()

    # ------------------------------------------------------- apertura
    def _on_double_click(self, event):
        # Ctrl+doppio clic su una cartella: aprila in una NUOVA scheda
        # (stesso comportamento dei browser). 0x0004 = modificatore Control.
        if (event.state & 0x0004) and len(self.tree.selection()) == 1:
            path = self.tree.selection()[0]
            if os.path.isdir(path):
                self.new_tab(path=path)
                return
        self._open_selected()

    def _on_tree_middle_click(self, event):
        """Clic centrale (rotella) su una cartella: aprila in una nuova scheda."""
        row = self.tree.identify_row(event.y)
        if row and os.path.isdir(row):
            self.new_tab(path=row)

    def _open_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        path = sel[0]
        if os.path.isdir(path):
            self.navigate_to(path, record=True)
        elif os.path.isfile(path):
            self._open_file(path)

    def _open_file(self, path):
        try:
            if IS_WINDOWS:
                os.startfile(path)
            else:
                opener = "xdg-open" if sys.platform.startswith("linux") else "open"
                subprocess.Popen([opener, path])
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile aprire il file:\n{exc}")

    def open_all_selected(self):
        """Apre con il programma predefinito tutti i FILE selezionati (salta le
        cartelle). Con più di 15 file chiede conferma per evitare di aprire
        decine di finestre per errore."""
        files = [p for p in self.tree.selection() if os.path.isfile(p)]
        if not files:
            return
        if len(files) > 15 and not messagebox.askyesno(
                "Apri tutti", f"Aprire {len(files)} file contemporaneamente?",
                parent=self):
            return
        for p in files:
            self._open_file(p)

    # ------------------------------------------------------- selezione
    def invert_selection(self):
        """Inverte la selezione nella scheda attiva (seleziona il complemento)."""
        current = set(self.tree.selection())
        all_items = set(self.tree.get_children())
        # selection_set accetta una sequenza di iid, non un set
        self.tree.selection_set(list(all_items - current))

    def select_by_size(self):
        """Seleziona tutti i FILE con dimensione >= alla soglia indicata
        (es. '10 MB', '500 KB', '2 GB' o byte)."""
        tab = self.active_tab
        if tab is None:
            return
        txt = simpledialog.askstring(
            "Seleziona per dimensione",
            "Dimensione minima (es. '10 MB', '500 KB', '2 GB' o byte):",
            parent=self)
        if not txt:
            return
        limit = parse_size(txt)
        if limit is None:
            messagebox.showerror(
                "Esplora File",
                "Dimensione non riconosciuta. Usa es. '10 MB', '500 KB', '2 GB'.")
            return
        items = [e[0] for e in tab.entries if not e[2] and e[3] >= limit]
        self.tree.selection_set(items)

    def select_by_age(self):
        """Seleziona i FILE modificati negli ultimi N giorni (utile per trovare
        i file recenti)."""
        tab = self.active_tab
        if tab is None:
            return
        txt = simpledialog.askstring(
            "Seleziona per data",
            "Seleziona i file modificati negli ultimi N giorni:",
            parent=self)
        if not txt:
            return
        try:
            days = float(txt.strip().replace(",", "."))
        except ValueError:
            messagebox.showerror("Esplora File", "Numero di giorni non valido.")
            return
        if days < 0:
            messagebox.showerror("Esplora File",
                                 "Il numero di giorni non può essere negativo.")
            return
        cutoff = datetime.datetime.now().timestamp() - days * 86400
        items = [e[0] for e in tab.entries if not e[2] and e[4] >= cutoff]
        self.tree.selection_set(items)

    def select_by_regex(self):
        """Seleziona gli elementi il cui NOME corrisponde a una regex
        (es. '^report_\\d{4}'). Completa la famiglia 'seleziona per…'."""
        tab = self.active_tab
        if tab is None:
            return
        txt = simpledialog.askstring(
            "Seleziona per regex",
            "Espressione regolare sul nome (es. '^report_\\d{4}'):",
            parent=self)
        if not txt:
            return
        try:
            pattern = re.compile(txt.strip(), re.IGNORECASE)
        except re.error as exc:
            messagebox.showerror("Esplora File", f"Regex non valida: {exc}")
            return
        items = [e[0] for e in tab.entries if pattern.search(e[1])]
        self.tree.selection_set(items)

    def select_same_extension(self):
        """Seleziona tutti gli elementi con la stessa estensione del primo
        elemento selezionato (per le cartelle: tutte le cartelle)."""
        sel = self.tree.selection()
        if not sel:
            return
        ext = os.path.splitext(sel[0])[1].lower()
        matching = [c for c in self.tree.get_children()
                    if os.path.splitext(c)[1].lower() == ext]
        if matching:
            self.tree.selection_set(matching)

    def select_by_pattern(self):
        """Seleziona gli elementi il cui nome corrisponde a un modello
        wildcard (es. *.txt, report*, *2024*)."""
        pattern = simpledialog.askstring(
            "Seleziona per modello",
            "Modello (es. *.txt, report*, *2024*):", parent=self)
        if not pattern:
            return
        pattern = pattern.strip()
        if not pattern:
            return
        matching = [c for c in self.tree.get_children()
                    if fnmatch.fnmatch(os.path.basename(c), pattern)]
        if matching:
            self.tree.selection_set(matching)
        else:
            messagebox.showinfo("Esplora File",
                                "Nessun elemento corrisponde al modello.")

    # ------------------------------------------------------- menu contestuale
    def set_preview_resolution(self):
        """Salva la risoluzione scelta e riapplica lo zoom all'immagine aperta."""
        settings = load_settings()
        settings["preview_resolution"] = int(self.preview_res.get())
        save_settings(settings)
        panel = getattr(self, "preview_panel", None)
        if panel is not None and panel._photo is not None:
            panel._zoom = panel._initial_zoom()
            panel._render_image()

    def _show_context_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        menu = tk.Menu(self, tearoff=0)
        sel = self.tree.selection()

        # sottomenu "Crea": file e cartelle nuovi
        create_menu = tk.Menu(menu, tearoff=0)
        create_menu.add_command(label="Cartella", command=self.new_folder)
        create_menu.add_command(label="File di testo (.txt)",
                                command=lambda: self.create_new_file("txt"))
        create_menu.add_command(label="Documento Word (.docx)",
                                command=lambda: self.create_new_file("docx"))
        create_menu.add_command(label="Presentazione PowerPoint (.pptx)",
                                command=lambda: self.create_new_file("pptx"))
        create_menu.add_command(label="File Python (.py)",
                                command=lambda: self.create_new_file("py"))

        # voce terminale: nella cartella cliccata se è un'unica cartella, altrimenti qui
        if row and len(sel) == 1 and os.path.isdir(sel[0]):
            term_target, term_label = sel[0], "Apri terminale nella cartella"
        else:
            term_target, term_label = self.current_path, "Apri terminale qui"
        menu.add_command(label=term_label,
                         command=lambda p=term_target: self.open_terminal_here(p))
        menu.add_separator()

        if sel:
            menu.add_command(label="Apri", command=self._open_selected)
            if len(sel) > 1 and any(os.path.isfile(p) for p in sel):
                menu.add_command(label="Apri tutti i selezionati",
                                 command=self.open_all_selected)
            if len(sel) == 1 and os.path.isdir(sel[0]):
                menu.add_command(label="Apri in nuova scheda",
                                 command=lambda p=sel[0]: self.new_tab(path=p))
            if len(sel) == 1 and not os.path.isdir(sel[0]) and IS_WINDOWS:
                menu.add_command(label="Apri con…", command=self.open_with)
            if len(sel) == 2 and all(os.path.isfile(p) for p in sel):
                menu.add_command(label="Confronta file", command=self.compare_selected)
            menu.add_separator()
            menu.add_command(label="Taglia", accelerator="Ctrl+X", command=self.cut_selected)
            menu.add_command(label="Copia", accelerator="Ctrl+C", command=self.copy_selected)
            menu.add_command(label="Copia percorso", command=self.copy_path)
            menu.add_command(label="Copia nome", command=self.copy_names)
            menu.add_command(label="Duplica", command=self.duplicate_selected)
            if len(sel) == 1 and IS_WINDOWS:
                menu.add_command(label="Crea collegamento",
                                 command=self.create_shortcut_selected)
            if any(os.path.isfile(p) for p in sel):
                menu.add_command(label="Calcola hash (MD5/SHA-256)",
                                 command=self.hash_selected)
            menu.add_command(label=REVEAL_LABEL,
                             command=lambda p=sel[0]: self.reveal_in_file_manager(p))
            menu.add_separator()
            zip_menu = tk.Menu(menu, tearoff=0)
            zip_menu.add_command(label="ZIP",
                                 command=lambda: self.compress_selected("zip"))
            zip_menu.add_command(label="7z",
                                 command=lambda: self.compress_selected("7z"))
            zip_menu.add_command(label="RAR (richiede WinRAR)",
                                 command=lambda: self.compress_selected("rar"))
            menu.add_cascade(label="Comprimi in…", menu=zip_menu)
            if len(sel) == 1 and sel[0].lower().endswith(
                    (".zip", ".7z", ".rar", ".tar", ".tgz", ".gz", ".bz2", ".xz")):
                menu.add_command(label="Estrai qui", command=self.extract_archive)
            if len(sel) == 1 and os.path.isdir(sel[0]):
                menu.add_command(label="Calcola dimensione", command=self.calc_folder_size)
                menu.add_command(label="Statistiche cartella", command=self.folder_statistics)
                menu.add_command(label="Cerca duplicati", command=self.find_duplicates)
                menu.add_command(label=self._fav_label(sel[0]),
                                 command=lambda p=sel[0]: self.toggle_favorite(p))
                menu.add_command(label="Presentazione (slideshow)",
                                 command=lambda p=sel[0]: self.start_slideshow(p))
            menu.add_separator()
            menu.add_command(label="Rinomina", accelerator="F2", command=self.rename_selected)
            if len(sel) >= 2 and any(os.path.isfile(p) for p in sel):
                menu.add_command(label="Rinomina in batch", command=self.batch_rename)
                menu.add_command(label="Sostituisci nei nomi…",
                                 command=self.replace_in_names)
            menu.add_command(label="Elimina definitivamente", accelerator="Canc",
                             command=self.delete_selected)
            menu.add_command(label="Distruggi (sovrascrivi)",
                             command=self.destroy_selected)
            if IS_WINDOWS:
                menu.add_command(label="Sposta nel cestino",
                                 command=self.move_to_recycle_bin)
            menu.add_command(label="Proprietà", command=self.show_properties)
            menu.add_separator()
            menu.add_cascade(label="Crea", menu=create_menu)
        else:
            menu.add_command(label="Incolla", accelerator="Ctrl+V", command=self.paste_clipboard)
            menu.add_command(label=REVEAL_LABEL, command=self.reveal_in_file_manager)
            menu.add_command(label=self._fav_label(self.current_path),
                             command=self.toggle_favorite)
            menu.add_command(label="Presentazione (slideshow)",
                             command=self.start_slideshow)
            menu.add_separator()
            menu.add_cascade(label="Crea", menu=create_menu)
            menu.add_command(label="Aggiorna", accelerator="F5", command=self.refresh)
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    # ------------------------------------------------------- terminale
    def open_terminal_here(self, path=None):
        """Apre un terminale nella cartella indicata (o in quella corrente)."""
        if path is None:
            path = self.current_path
        if not os.path.isdir(path):
            messagebox.showerror("Esplora File", "Cartella non valida.")
            return
        try:
            if IS_WINDOWS:
                # Windows Terminal se disponibile, altrimenti cmd
                wt = shutil.which("wt.exe")
                if wt:
                    subprocess.Popen([wt, "-d", path])
                else:
                    subprocess.Popen(["cmd", "/k"], cwd=path,
                                     creationflags=subprocess.CREATE_NEW_CONSOLE)
            else:
                for term, args in (
                        ("x-terminal-emulator", ["--working-directory", path]),
                        ("gnome-terminal", ["--working-directory", path]),
                        ("konsole", ["--workdir", path]),
                        ("xfce4-terminal", ["--working-directory", path]),
                        ("xterm", ["-e", f"cd '{path}' && exec $SHELL"])):
                    if shutil.which(term):
                        subprocess.Popen([term] + args)
                        return
                raise OSError("Nessun emulatore di terminale trovato")
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile aprire il terminale:\n{exc}")

    # ------------------------------------------------------- file manager di sistema
    def reveal_in_file_manager(self, path=None):
        """Apre il file manager di sistema mostrando il percorso indicato
        (o la cartella corrente). Su Windows il file/cartella risulta
        selezionato; su macOS viene rivelato nel Finder; su Linux apre la
        cartella contenitrice (il reveal per singolo file non è standard)."""
        if path is None:
            path = self.current_path
        path = os.path.normpath(str(path))
        if not os.path.exists(path):
            messagebox.showerror("Esplora File", "Percorso non trovato.")
            return
        try:
            if IS_WINDOWS:
                # /select, evidenzia l'elemento anche quando è un file
                subprocess.Popen(
                    ["explorer", "/select,", os.path.normpath(path)],
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            elif IS_MACOS:
                subprocess.Popen(["open", "-R", path])
            else:
                # Linux: apri la cartella (o la cartella del file selezionato)
                target = path if os.path.isdir(path) else os.path.dirname(path)
                opener = shutil.which("xdg-open") or shutil.which("gio")
                if opener:
                    subprocess.Popen([opener, target])
                else:
                    raise OSError("Nessun gestore di file trovato (xdg-open/gio)")
        except OSError as exc:
            messagebox.showerror("Esplora File",
                                 f"Impossibile aprire il file manager:\n{exc}")

    # ------------------------------------------------------- clipboard
    def copy_selected(self):
        paths = list(self.tree.selection())
        if paths:
            self.clipboard = {"mode": "copy", "paths": paths}
            self._update_status()

    def cut_selected(self):
        paths = list(self.tree.selection())
        if paths:
            self.clipboard = {"mode": "cut", "paths": paths}
            self._update_status()

    def copy_path(self):
        """Copia i percorsi COMPLETI degli elementi selezionati, uno per riga
        (prima copiava solo il primo: ora è coerente con 'Copia nome')."""
        sel = self.tree.selection()
        if not sel:
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(str(p) for p in sel))

    def copy_names(self):
        """Copia i NOMI (senza percorso) degli elementi selezionati, uno per riga
        (comodo per elenchi, log o report)."""
        sel = self.tree.selection()
        if not sel:
            return
        names = "\n".join(os.path.basename(p) for p in sel)
        self.clipboard_clear()
        self.clipboard_append(names)

    def export_listing_csv(self):
        """Esporta l'elenco visualizzato (cartella corrente o risultati di ricerca)
        in un file CSV con Nome, Tipo, Dimensione, Data modifica e Percorso.
        Il BOM (utf-8-sig) fa sì che Excel apra correttamente gli accenti."""
        tab = self.active_tab
        if tab is None or not tab.entries:
            messagebox.showinfo("Esporta elenco", "Nessun elemento da esportare.")
            return
        base = os.path.basename(tab.current_path.rstrip("/\\")) or "file"
        dest = filedialog.asksaveasfilename(
            title="Esporta elenco (CSV)",
            defaultextension=".csv",
            filetypes=[("File CSV", "*.csv"), ("Tutti i file", "*.*")],
            initialfile=f"elenco_{base}.csv",
            parent=self)
        if not dest:
            return
        try:
            with open(dest, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["Nome", "Tipo", "Dimensione (byte)",
                            "Data modifica", "Percorso"])
                for entry in tab.entries:
                    path, name, is_dir, size, mtime, rel = entry
                    tipo = "Cartella" if is_dir else tab._type_label(entry)
                    w.writerow([name, tipo, size, format_date(mtime), path])
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Esportazione non riuscita:\n{exc}")
            return
        self.status_left.config(text=f"Elenco esportato in {os.path.basename(dest)}")

    def paste_clipboard(self):
        if self._copy_move_busy or not self.clipboard["paths"]:
            return
        paths = list(self.clipboard["paths"])
        mode = self.clipboard["mode"]
        if mode == "cut":
            self.clipboard = {"mode": None, "paths": []}
        dest_dir = self.current_path
        self._copy_move(paths, dest_dir, mode,
                        on_finished=lambda: self.refresh(keep_selection=False))
        self._update_status()

    @staticmethod
    def _unique_path(path):
        if not os.path.exists(path):
            return path
        parent, name = os.path.split(path)
        stem, ext = os.path.splitext(name)
        i = 1
        while True:
            new_name = f"{stem} (copia {i}){ext}" if i > 1 else f"{stem} (copia){ext}"
            candidate = os.path.join(parent, new_name)
            if not os.path.exists(candidate):
                return candidate
            i += 1

    # ------------------------------------------------------- operazioni file
    def new_folder(self):
        name = simpledialog.askstring("Nuova cartella", "Nome della cartella:",
                                      parent=self, initialvalue="Nuova cartella")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        path = self._unique_path(os.path.join(self.current_path, name))
        try:
            os.mkdir(path)
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile creare la cartella:\n{exc}")
            return
        self.refresh()
        self.tree.selection_set(path)
        self.tree.see(path)

    def _valid_filename(self, name):
        """True se 'name' è un nome di file/cartella accettabile."""
        if not name or name in (".", ".."):
            return False
        if any(c in name for c in '<>:"/\\|?*'):
            return False
        if name.endswith((" ", ".")):
            return False
        if IS_WINDOWS:
            stem = name.split(".")[0].rstrip().upper()
            if stem in {"CON", "PRN", "AUX", "NUL"} or (
                    stem[:3] in {"COM", "LPT"} and stem[3:].isdigit()):
                return False
        return True

    def create_new_file(self, kind, name=None):
        """Crea un nuovo file (txt, py, docx, pptx) chiedendone il nome.

        Se 'name' è fornito (test) non apre alcun dialogo. L'estensione
        viene aggiunta automaticamente se manca; in caso di conflitto il
        nome viene reso unico (come per le cartelle).
        """
        spec = CREATE_FILE_TYPES.get(kind)
        if spec is None:
            return
        ext, default_name, label = spec
        if name is None:
            name = simpledialog.askstring("Crea file",
                                          f"Nuovo {label.lower()}:",
                                          parent=self,
                                          initialvalue=default_name)
            if not name:
                return
        name = name.strip()
        if not name:
            return
        if os.path.splitext(name)[1].lower() != ext:
            name += ext
        if not self._valid_filename(name):
            messagebox.showerror("Esplora File",
                                 "Nome non valido (caratteri proibiti o nome riservato).")
            return
        path = self._unique_path(os.path.join(self.current_path, name))
        try:
            if kind in ("docx", "pptx"):
                create_blank_ooxml(path, kind)
            else:
                content = "# Nuovo script Python\n" if kind == "py" else ""
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile creare il file:\n{exc}")
            return
        self.refresh()
        self.tree.selection_set(path)
        self.tree.see(path)

    def rename_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        path = sel[0]
        old_name = os.path.basename(path)
        new_name = simpledialog.askstring("Rinomina", "Nuovo nome:", parent=self,
                                          initialvalue=old_name)
        if not new_name or new_name.strip() == old_name:
            return
        new_name = new_name.strip()
        if not new_name:
            return
        target = os.path.join(os.path.dirname(path), new_name)
        if os.path.exists(target):
            messagebox.showerror("Esplora File", "Esiste già un elemento con questo nome.")
            return
        try:
            os.rename(path, target)
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile rinominare:\n{exc}")
            return
        self.refresh()
        if os.path.exists(target):
            self.tree.selection_set(target)
            self.tree.see(target)

    # ------------------------------------------------------- rinomina in batch
    def batch_rename(self):
        """Rinomina in batch i file selezionati con un modello a contatore.

        Il modello supporta {n} (contatore), {name} (nome senza estensione)
        e {ext} (estensione). L'anteprima si aggiorna in tempo reale.
        """
        sel = sorted(p for p in self.tree.selection() if os.path.isfile(p))
        if not sel:
            messagebox.showinfo("Esplora File",
                                "Seleziona i file da rinominare in batch.")
            return
        win = tk.Toplevel(self)
        win.title("Rinomina in batch")
        win.transient(self)
        win.geometry("560x440")

        ttk.Label(win, text="Modello: {n} = contatore, {name} = nome, "
                            "{ext} = estensione",
                  font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(10, 2))
        model_var = tk.StringVar(value="foto_{n}{ext}")
        ttk.Entry(win, textvariable=model_var).pack(fill="x", padx=12)

        rowf = ttk.Frame(win)
        rowf.pack(fill="x", padx=12, pady=(6, 4))
        ttk.Label(rowf, text="Numero iniziale:").pack(side="left")
        start_var = tk.IntVar(value=1)
        ttk.Spinbox(rowf, from_=0, to=999999, textvariable=start_var,
                    width=8).pack(side="left", padx=(6, 0))

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=(4, 6))
        tree = ttk.Treeview(frame, columns=("old", "new"), show="headings")
        tree.heading("old", text="Nome attuale")
        tree.heading("new", text="Nuovo nome")
        tree.column("old", width=260, anchor="w", stretch=True)
        tree.column("new", width=260, anchor="w", stretch=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _start_value():
            # lo Spinbox può contenere testo non numerico digitato a mano
            try:
                return int(start_var.get())
            except (tk.TclError, ValueError):
                return 1

        def _refresh_preview():
            tree.delete(*tree.get_children())
            model = model_var.get() or "{name}{ext}"
            start = _start_value()
            for i, p in enumerate(sel):
                tree.insert("", "end", iid=p, values=(
                    os.path.basename(p), build_batch_name(model, start, i, p)))

        model_var.trace_add("write", lambda *a: _refresh_preview())
        start_var.trace_add("write", lambda *a: _refresh_preview())
        _refresh_preview()

        def _apply():
            model = model_var.get() or "{name}{ext}"
            start = _start_value()
            errors = []
            renamed = 0
            for i, p in enumerate(sel):
                new = build_batch_name(model, start, i, p)
                if not new or new == os.path.basename(p):
                    continue
                target = os.path.join(os.path.dirname(p), new)
                if os.path.exists(target) \
                        and os.path.normcase(target) != os.path.normcase(p):
                    errors.append(f"{os.path.basename(p)}: destinazione già esistente")
                    continue
                try:
                    os.rename(p, target)
                    renamed += 1
                except OSError as exc:
                    errors.append(f"{os.path.basename(p)}: {exc}")
            if errors:
                messagebox.showerror("Esplora File",
                                     "Alcuni file non sono stati rinominati:\n\n"
                                     + "\n".join(errors))
            win.destroy()
            if renamed:
                self.refresh()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Button(btns, text="Applica", command=_apply).pack(side="right")
        ttk.Button(btns, text="Annulla", command=win.destroy).pack(side="right",
                                                                   padx=(0, 6))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def replace_in_names(self):
        """Rinomina la selezione sostituendo una sottostringa nei nomi
        ('trova' → 'sostituisci'), con anteprima dal vivo prima di applicare."""
        sel = sorted(p for p in self.tree.selection() if os.path.isfile(p))
        if not sel:
            messagebox.showinfo("Esplora File",
                                "Seleziona i file da rinominare.")
            return
        win = tk.Toplevel(self)
        win.title("Sostituisci nei nomi")
        win.transient(self)
        win.geometry("560x440")

        ttk.Label(win, text="Trova:").pack(anchor="w", padx=12, pady=(10, 2))
        find_var = tk.StringVar()
        ttk.Entry(win, textvariable=find_var).pack(fill="x", padx=12)
        ttk.Label(win, text="Sostituisci con:").pack(anchor="w", padx=12, pady=(6, 2))
        repl_var = tk.StringVar()
        ttk.Entry(win, textvariable=repl_var).pack(fill="x", padx=12)

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=(8, 6))
        tree = ttk.Treeview(frame, columns=("old", "new"), show="headings")
        tree.heading("old", text="Nome attuale")
        tree.heading("new", text="Nuovo nome")
        tree.column("old", width=260, anchor="w", stretch=True)
        tree.column("new", width=260, anchor="w", stretch=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _refresh_preview():
            tree.delete(*tree.get_children())
            find = find_var.get()
            for p in sel:
                base = os.path.basename(p)
                new = base.replace(find, repl_var.get()) if find else base
                tree.insert("", "end", iid=p, values=(base, new))

        find_var.trace_add("write", lambda *a: _refresh_preview())
        repl_var.trace_add("write", lambda *a: _refresh_preview())
        _refresh_preview()

        def _apply():
            find = find_var.get()
            repl = repl_var.get()
            if not find:
                messagebox.showinfo("Esplora File", "Indica il testo da cercare.")
                return
            renamed, errors = self._do_replace_rename(sel, find, repl)
            if errors:
                messagebox.showerror("Esplora File",
                                     "Alcuni file non sono stati rinominati:\n\n"
                                     + "\n".join(errors))
            win.destroy()
            if renamed:
                self.refresh()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Button(btns, text="Applica", command=_apply).pack(side="right")
        ttk.Button(btns, text="Annulla", command=win.destroy).pack(side="right",
                                                                   padx=(0, 6))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def _do_replace_rename(self, sel, find, repl):
        """Esegue la sostituzione 'find' → 'repl' nei nomi selezionati.
        Restituisce (numero rinominati, lista errori)."""
        errors = []
        renamed = 0
        for p in sel:
            base = os.path.basename(p)
            new = base.replace(find, repl)
            if not new or new == base or new in (".", ".."):
                continue
            target = os.path.join(os.path.dirname(p), new)
            if os.path.exists(target) \
                    and os.path.normcase(target) != os.path.normcase(p):
                errors.append(f"{base}: destinazione già esistente")
                continue
            try:
                os.rename(p, target)
                renamed += 1
            except OSError as exc:
                errors.append(f"{base}: {exc}")
        return renamed, errors

    def delete_selected(self):
        """Elimina DEFINITIVAMENTE la selezione (senza passare dal Cestino)."""
        sel = list(self.tree.selection())
        if not sel:
            return
        if any(_is_drive_root(p) for p in sel):
            messagebox.showerror(
                "Esplora File",
                "Non è possibile eliminare un'intera unità (es. C:\\):\n"
                "seleziona cartelle o file specifici.")
            return
        protette = [p for p in sel if _is_system_path(p)]
        if protette:
            messagebox.showerror(
                "Esplora File",
                "Non è possibile eliminare cartelle di sistema:\n\n"
                + "\n".join(protette)
                + "\n\nSono cartelle protette da Windows.")
            return
        n = len(sel)
        label = "questo elemento" if n == 1 else f"questi {n} elementi"
        if not messagebox.askyesno(
                "Elimina definitivamente",
                f"Eliminare DEFINITIVAMENTE {label}?\n\n"
                "Non finirà nel Cestino: l'operazione è irreversibile.",
                icon="warning", parent=self):
            return
        errors = []
        for p in sel:
            try:
                if _is_link_or_junction(p):
                    # rimuovi SOLO il puntatore, mai il contenuto puntato
                    # (potrebbe stare fuori dalla selezione)
                    try:
                        os.remove(p)
                    except OSError:
                        os.rmdir(p)
                elif os.path.isdir(p):
                    # scansione senza-follow: i junction interni vengono
                    # rimossi da soli (shutil.rmtree li seguirebbe!)
                    delete_tree_no_follow(p, errors)
                    try:
                        os.rmdir(p)
                    except OSError as exc:
                        errors.append(f"{os.path.basename(p)}: {exc}")
                else:
                    os.remove(p)
            except OSError as exc:
                errors.append(f"{os.path.basename(p)}: {exc}")
        if errors:
            messagebox.showerror("Esplora File",
                                 "Alcuni elementi non sono stati eliminati:\n\n"
                                 + "\n".join(errors))
        self.refresh()

    def move_to_recycle_bin(self):
        """Sposta la selezione nel Cestino di Windows (senza conferma)."""
        sel = list(self.tree.selection())
        if not sel:
            return
        if not IS_WINDOWS:
            messagebox.showinfo("Esplora File",
                                "Il Cestino è disponibile solo su Windows.")
            return
        if any(_is_drive_root(p) for p in sel):
            messagebox.showerror("Esplora File",
                                 "Non è possibile spostare un'intera unità nel Cestino.")
            return
        if send_to_recycle_bin(sel):
            self.refresh()
        else:
            messagebox.showerror("Esplora File",
                                 "Impossibile spostare gli elementi nel Cestino.")

    def destroy_selected(self):
        """Distrugge DEFINITIVAMENTE la selezione: sovrascrive i dati con
        byte casuali (SECURE_DELETE_PASSES passaggi) e poi cancella.

        Protezioni: rifiuta radici di unità e cartelle di sistema, blocca
        le selezioni oltre 5 GB e richiede una seconda conferma (digitando
        'elimina' quando la selezione supera 500 MB).
        """
        sel = list(self.tree.selection())
        if not sel:
            return
        # 1) radici di unità
        if any(_is_drive_root(p) for p in sel):
            messagebox.showerror(
                "Esplora File",
                "Non è possibile distruggere un'intera unità (es. C:\\):\n"
                "seleziona cartelle o file specifici.")
            return
        # 2) cartelle critiche di sistema
        protette = [p for p in sel if _is_system_path(p)]
        if protette:
            messagebox.showerror(
                "Esplora File",
                "Non è possibile distruggere cartelle di sistema:\n\n"
                + "\n".join(protette)
                + "\n\nSono cartelle protette da Windows.")
            return
        # 3) limite di dimensione (5 GB): rifiuto con misurazione anticipata
        self.status_left.config(text="Verifica dimensioni…")
        self.update_idletasks()
        totale, oltre = _selection_size(sel, DESTROY_SIZE_LIMIT)
        self.status_left.config(text="")
        if oltre:
            messagebox.showerror(
                "Esplora File",
                f"Selezione troppo grande ({format_size(totale)}).\n\n"
                f"Per sicurezza la distruzione è limitata a "
                f"{format_size(DESTROY_SIZE_LIMIT)}.\n"
                "Elimina una parte dei file alla volta.")
            return
        # 4) prima conferma
        n = len(sel)
        label = "questo elemento" if n == 1 else f"questi {n} elementi"
        if not messagebox.askyesno(
                "Distruggi file",
                f"Distruggere DEFINITIVAMENTE {label}?\n\n"
                f"Il contenuto verrà sovrascritto {SECURE_DELETE_PASSES} volte "
                "con dati casuali e poi cancellato.\n"
                "NON sarà possibile recuperarlo in alcun modo.\n\n"
                "(Nota: su SSD la sovrascrittura non garantisce la\n"
                "cancellazione fisica dei dati.)",
                icon="warning", parent=self):
            return
        # 5) seconda conferma: oltre 500 MB si deve digitare 'elimina'
        if totale > DESTROY_TYPING_LIMIT:
            if not self._confirm_destroy_typing(totale):
                return
        else:
            if not messagebox.askokcancel(
                    "Ultima conferma",
                    f"Confermi la distruzione IRREVERSIBILE di {label} "
                    f"({format_size(totale)})?\n\n"
                    "I dati saranno sovrascritti e cancellati per sempre.",
                    icon="warning", parent=self):
                return
        self.status_left.config(text="Distruzione in corso…")

        def worker(q=None):
            return secure_delete(
                sel, passes=SECURE_DELETE_PASSES,
                progress_cb=lambda d, t: (q.put(("progress", (d, t)))
                                           if q is not None else None))

        def on_progress(value):
            d, t = value
            self.status_left.config(text=f"Sovrascrittura file {d}/{t}…")

        def on_done(kind, value):
            self.status_left.config(text="")
            if kind != "ok":
                messagebox.showerror("Esplora File", f"Errore: {value}")
                return
            if value:
                messagebox.showerror(
                    "Esplora File",
                    "Alcuni file non sono stati distrutti:\n\n"
                    + "\n".join(value))
            self.refresh()

        self._run_in_thread(worker, on_done, on_progress=on_progress)

    def _confirm_destroy_typing(self, totale):
        """Seconda conferma per distruzioni oltre 500 MB: bisogna digitare
        'elimina'. Ritorna True se l'utente conferma."""
        win = tk.Toplevel(self)
        win.title("Conferma distruzione")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        ttk.Label(win, text="ATTENZIONE: operazione IRREVERSIBILE",
                  font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, padx=16, pady=(14, 2), sticky="w")
        ttk.Label(win, text=("Stai per distruggere definitivamente "
                             f"{format_size(totale)} di dati.\n"
                             "Per confermare, digita 'elimina' nella casella:"),
                  wraplength=360, justify="left").grid(
            row=1, column=0, columnspan=2, padx=16, pady=4, sticky="w")

        entry = ttk.Entry(win, width=30)
        entry.grid(row=2, column=0, columnspan=2, padx=16, pady=6, sticky="we")

        err = ttk.Label(win, text="", foreground="#c00000")
        err.grid(row=3, column=0, columnspan=2, padx=16, pady=0, sticky="w")

        result = {"ok": False}

        def _confirm(_e=None):
            if entry.get().strip().lower() == "elimina":
                result["ok"] = True
                win.destroy()
            else:
                err.config(text="Testo non corretto: devi scrivere 'elimina'.")

        def _cancel():
            win.destroy()

        ttk.Button(win, text="Conferma", command=_confirm).grid(
            row=4, column=0, sticky="e", padx=(16, 6), pady=12)
        ttk.Button(win, text="Annulla", command=_cancel).grid(
            row=4, column=1, sticky="w", padx=(6, 16), pady=12)
        entry.bind("<Return>", _confirm)
        win.bind("<Escape>", lambda e: _cancel())

        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")
        entry.focus_set()
        self.wait_window(win)
        return result["ok"]

    # ------------------------------------------------------- drag & drop
    def _dnd_press(self, event):
        # non toccare la selezione qui: la gestisce il binding di classe della
        # Treeview (così Ctrl+clic e selezione multipla continuano a funzionare)
        tree = event.widget
        # Il click su un'intestazione o sul separatore NON è un drag & drop:
        # è il ridimensionamento nativo delle colonne di ttk. Se qui partisse
        # il drag, i "break" di _dnd_motion/_dnd_release bloccherebbero i
        # binding di classe (resize.drag / resize.release) e la larghezza
        # tornerebbe a quella originale al rilascio del click.
        try:
            region = tree.identify_region(event.x, event.y)
        except tk.TclError:
            region = ""
        if region in ("heading", "separator", "nothing"):
            self.drag = None
            return
        self.drag = {"active": False, "start": (event.x_root, event.y_root),
                     "paths": [], "src_tree": tree, "src_tab": self.tab_of(tree),
                     "highlight": None}

    def _dnd_motion(self, event):
        if not self.drag:
            return
        try:
            d = self.drag
            if not d["active"]:
                dx = event.x_root - d["start"][0]
                dy = event.y_root - d["start"][1]
                if max(abs(dx), abs(dy)) < 5:
                    return
                paths = list(d["src_tree"].selection())
                if not paths:
                    self.drag = None
                    return
                d["active"] = True
                d["paths"] = paths
                d["src_tree"].config(cursor="hand2")
            self._dnd_update_highlight(event)
        except tk.TclError:
            self.drag = None   # scheda chiusa durante il drag: ripulisci
        return "break"

    def _dnd_update_highlight(self, event):
        self._dnd_clear_highlight()
        target = self._dnd_target(event)
        if not target:
            return
        kind, ref, path = target
        if kind == "tree_row":
            widget, iid = ref
            tags = tuple(widget.item(iid, "tags"))
            if "dnd_target" not in tags:
                widget.item(iid, tags=tags + ("dnd_target",))
                self.drag["highlight"] = (widget, iid, tags)
        elif kind == "sidebar":
            iid = ref
            tags = tuple(self.sidebar.item(iid, "tags"))
            self.sidebar.item(iid, tags=tags + ("dnd_target",))
            self.drag["highlight"] = (self.sidebar, iid, tags)

    def _dnd_clear_highlight(self):
        if not self.drag:
            return
        h = self.drag.get("highlight")
        if h:
            widget, iid, tags = h
            try:
                widget.item(iid, tags=tags)
            except tk.TclError:
                pass
            self.drag["highlight"] = None

    def _dnd_target(self, event):
        """Ritorna (kind, ref, percorso) oppure None."""
        w = self.winfo_containing(event.x_root, event.y_root)
        if w is None:
            return None
        # scheda: rilascio sull'intestazione di una scheda
        if w is self.tabbar.canvas:
            lx = event.x_root - self.tabbar.canvas.winfo_rootx()
            ly = event.y_root - self.tabbar.canvas.winfo_rooty()
            idx = self.tabbar.index_at(lx, ly)
            if 0 <= idx < len(self.tabs):
                return ("tab", idx, self.tabs[idx].current_path)
            return None
        # albero di una scheda
        for tab in self.tabs:
            if w is tab.tree:
                lx = event.x_root - tab.tree.winfo_rootx()
                ly = event.y_root - tab.tree.winfo_rooty()
                row = tab.tree.identify_row(ly)
                if row and row not in self.drag["paths"]:
                    if os.path.isdir(row):
                        return ("tree_row", (tab.tree, row), row)
                    return None  # file: nessun bersaglio
                if row in self.drag["paths"]:
                    return None  # sopra se stesso
                return ("tree_area", tab, tab.current_path)
        # barra laterale
        if w is self.sidebar:
            ly = event.y_root - self.sidebar.winfo_rooty()
            iid = self.sidebar.identify_row(ly)
            if iid and iid in self._sidebar_map:
                return ("sidebar", iid, self._sidebar_map[iid])
            return None
        return None

    def _dnd_release(self, event):
        if not self.drag:
            return
        d = self.drag
        try:
            self._dnd_clear_highlight()
            if not d["active"]:
                self.drag = None
                return
            # calcola il bersaglio PRIMA di azzerare self.drag (serve a _dnd_target)
            target = self._dnd_target(event)
            self.drag = None
            # ripristina il cursore PRIMA dell'operazione (potrebbe richiedere
            # tempo se copia molti file)
            try:
                d["src_tree"].config(cursor="")
            except tk.TclError:
                pass
            if target and target[2]:
                copy = bool(event.state & 0x4)  # Ctrl premuto = copia
                self._do_dnd(d["paths"], target[2], copy, d["src_tab"])
        except tk.TclError:
            self.drag = None
        finally:
            # il cursore va ripristinato SEMPRE, anche se un'eccezione ha
            # interrotto il flusso qui sopra (es. finestra chiusa durante
            # il rilascio): altrimenti resterebbe bloccato su hand2
            try:
                d["src_tree"].config(cursor="")
            except (tk.TclError, KeyError, AttributeError):
                pass
        return "break"

    def _do_dnd(self, paths, dest, copy, src_tab):
        dest = os.path.normpath(dest)
        mode = "copy" if copy else "move"

        def on_finished():
            if src_tab is not None and src_tab.frame.winfo_exists():
                src_tab.refresh(keep_selection=False)
            dest_tab = self.tab_of_path(dest)
            if dest_tab is not None:
                dest_tab.refresh(keep_selection=False)

        self._copy_move(paths, dest, mode, on_finished=on_finished)

    @staticmethod
    def _is_within(parent, child):
        """True se child è dentro (o uguale a) parent."""
        parent = os.path.normcase(os.path.abspath(parent))
        child = os.path.normcase(os.path.abspath(child))
        return child == parent or child.startswith(parent + os.sep)

    def tab_of(self, tree):
        for tab in self.tabs:
            if tab.tree is tree:
                return tab
        return None

    def tab_of_path(self, path):
        norm = os.path.normcase(os.path.normpath(path))
        for tab in self.tabs:
            if os.path.normcase(tab.current_path) == norm:
                return tab
        return None

    # ------------------------------------------------------- ZIP e dimensioni
    def _run_in_thread(self, worker, on_done, on_progress=None):
        """Esegue worker(q) in un thread e richiama on_done(kind, value) nel main.

        Il worker può pubblicare progressi con q.put(("progress", valore)):
        vengono consegnati a on_progress finché non arriva il risultato.
        """
        q = queue.Queue()

        def _worker():
            try:
                q.put(("ok", worker(q)))
            except Exception as exc:
                q.put(("err", str(exc)))

        threading.Thread(target=_worker, daemon=True).start()

        def _poll():
            while True:
                try:
                    kind, value = q.get_nowait()
                except queue.Empty:
                    try:
                        self.after(60, _poll)
                    except tk.TclError:
                        pass   # finestra distrutta durante l'attesa
                    return
                if kind == "progress":
                    if on_progress:
                        try:
                            on_progress(value)
                        except Exception:
                            pass   # un progresso non deve bloccare il risultato
                    continue
                on_done(kind, value)
                return

        self.after(60, _poll)

    # ------------------------------------------------------- copia/sposta
    def _progress_show(self, maximum=100):
        """Mostra la barra di avanzamento determinata con il massimo indicato."""
        self.progress_bar.config(mode="determinate", maximum=maximum, value=0)
        self.progress_bar.pack(side="left", padx=(10, 0))

    def _progress_set(self, value):
        self.progress_bar.config(value=value)

    def _progress_hide(self):
        try:
            self.progress_bar.stop()   # ferma l'animazione (no-op se determinata)
        except tk.TclError:
            pass
        self.progress_bar.pack_forget()

    def _progress_indeterminate(self):
        """Mostra la barra in modalità indeterminata (animata) per le
        operazioni lunghe di cui non si conosce il totale a priori."""
        self.progress_bar.config(mode="indeterminate", maximum=100)
        self.progress_bar.pack(side="left", padx=(10, 0))
        self.progress_bar.start(20)

    def _copy_move(self, paths, dest, mode, on_finished=None):
        """Copia ('copy') o sposta ('move') i percorsi in dest, in background.

        Mostra una barra di avanzamento determinata (numero di file) e, al
        termine, richiama on_finished() sul thread principale per aggiornare
        le viste coinvolte.
        """
        if self._copy_move_busy:
            return
        dest = os.path.normpath(dest)
        plan = []                  # (src, target, is_dir, nfile)
        static_errors = []
        for src in paths:
            src = os.path.normpath(src)
            if os.path.abspath(src) == os.path.abspath(dest):
                continue
            if self._is_within(src, dest):
                static_errors.append(
                    f"{os.path.basename(src)}: destinazione all'interno della sorgente")
                continue
            name = os.path.basename(src)
            target = os.path.join(dest, name)
            if os.path.normcase(os.path.abspath(target)) == os.path.normcase(os.path.abspath(src)):
                continue  # già nella cartella di destinazione
            target = self._unique_path(target)
            try:
                nfiles = count_files(src)
            except OSError:
                nfiles = 1
            plan.append((src, target, os.path.isdir(src), nfiles))
        if not plan:
            if static_errors:
                messagebox.showerror("Esplora File", "\n".join(static_errors))
            return

        total = max(1, sum(p[3] for p in plan))
        is_copy = (mode == "copy")
        label = "Copia" if is_copy else "Spostamento"
        self._copy_move_busy = True
        self._progress_show(total)
        self.status_left.config(text=f"{label} in corso… (0/{total})")

        def worker(q):
            errs = list(static_errors)
            done = [0]

            def bump(n=1):
                done[0] += n
                q.put(("progress", done[0]))

            for src, target, is_dir, nfiles in plan:
                try:
                    if not is_copy:
                        shutil.move(src, target)
                        bump(nfiles)
                    elif is_dir:
                        copy_tree_with_progress(src, target, bump)
                    else:
                        shutil.copy2(src, target)
                        bump(1)
                except (OSError, shutil.Error) as exc:
                    errs.append(f"{os.path.basename(src)}: {exc}")
            return errs

        def on_progress(done):
            done = min(done, total)
            self.status_left.config(text=f"{label} in corso… ({done}/{total})")
            self._progress_set(done)

        def on_done(kind, value):
            self._copy_move_busy = False
            self._progress_hide()
            if kind != "ok":
                self.status_left.config(text=f"{label} non riuscito.")
                messagebox.showerror("Esplora File", f"{label} non riuscita:\n{value}")
                return
            errs = value or []
            if on_finished is not None:
                try:
                    on_finished()
                except Exception:
                    pass
            if errs:
                messagebox.showerror(
                    "Esplora File",
                    "Alcune operazioni non sono riuscite:\n\n" + "\n".join(errs))
            else:
                self.status_left.config(text=f"{label} completata ({total} file).")
            self._update_status()

        self._run_in_thread(worker, on_done, on_progress=on_progress)

    def duplicate_selected(self):
        """Duplica la selezione nella stessa cartella ('nome (copia).ext').

        Come copia/sposta gira in background con barra di avanzamento, così
        duplicare cartelle grandi non blocca l'interfaccia.
        """
        sel = list(self.tree.selection())
        if not sel or self._copy_move_busy:
            return
        plan = []
        for src in sel:
            src = os.path.normpath(src)
            target = self._unique_path(src)
            try:
                nfiles = count_files(src)
            except OSError:
                nfiles = 1
            plan.append((src, target, os.path.isdir(src), nfiles))
        if not plan:
            return
        total = max(1, sum(p[3] for p in plan))
        self._copy_move_busy = True
        self._progress_show(total)
        self.status_left.config(text=f"Duplicazione in corso… (0/{total})")

        def worker(q):
            errs = []
            done = [0]

            def bump(n=1):
                done[0] += n
                q.put(("progress", done[0]))

            for src, target, is_dir, nfiles in plan:
                try:
                    if is_dir:
                        copy_tree_with_progress(src, target, bump)
                    else:
                        shutil.copy2(src, target)
                        bump(1)
                except (OSError, shutil.Error) as exc:
                    errs.append(f"{os.path.basename(src)}: {exc}")
            return errs

        def on_progress(done):
            done = min(done, total)
            self.status_left.config(text=f"Duplicazione in corso… ({done}/{total})")
            self._progress_set(done)

        def on_done(kind, value):
            self._copy_move_busy = False
            self._progress_hide()
            if kind != "ok":
                self.status_left.config(text="Duplicazione non riuscita.")
                messagebox.showerror("Esplora File", f"Duplicazione non riuscita:\n{value}")
                return
            errs = value or []
            if errs:
                messagebox.showerror("Esplora File",
                                     "Alcune duplicazioni non sono riuscite:\n\n"
                                     + "\n".join(errs))
            else:
                self.status_left.config(text=f"Duplicazione completata ({total} file).")
            self.refresh()

        self._run_in_thread(worker, on_done, on_progress=on_progress)

    # ------------------------------------------------------- collegamenti (.lnk)
    def create_shortcut_selected(self):
        """Crea un collegamento (.lnk) al file/cartella selezionato."""
        if not IS_WINDOWS:
            messagebox.showinfo("Esplora File",
                                "I collegamenti .lnk sono disponibili solo su Windows.")
            return
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showinfo("Esplora File",
                                "Seleziona un singolo file o cartella.")
            return
        target = sel[0]
        parent = os.path.dirname(target) or self.current_path
        name = os.path.basename(os.path.normpath(target))
        lnk = self._unique_path(os.path.join(parent, f"{name}.lnk"))
        if create_shortcut_lnk(target, lnk, work_dir=parent):
            self.status_left.config(text=f"Collegamento creato: {os.path.basename(lnk)}")
            self.refresh()
        else:
            messagebox.showerror("Esplora File", "Impossibile creare il collegamento.")

    def compress_selected(self, fmt):
        """Comprime la selezione in ZIP/7z/RAR (in background)."""
        sel = list(self.tree.selection())
        if not sel:
            return
        base_dir = os.path.dirname(sel[0]) or self.current_path
        if len(sel) == 1:
            name = os.path.basename(os.path.normpath(sel[0]))
        else:
            name = os.path.basename(os.path.normpath(self.current_path)) or "Archivio"
        dest = self._unique_path(os.path.join(base_dir, f"{name}.{fmt}"))
        label = fmt.upper()
        self.status_left.config(text=f"Compressione {label} in corso…")
        self._progress_indeterminate()

        def worker(q=None):
            if fmt == "zip":
                return create_zip(sel, dest)
            if fmt == "7z":
                return create_7z(sel, dest)
            return create_rar(sel, dest)

        def on_done(kind, value):
            self._progress_hide()
            if kind == "ok" and value is None:
                self.status_left.config(text=f"Creato {os.path.basename(dest)}")
                self.refresh()
            else:
                messagebox.showerror("Esplora File",
                                     f"Compressione {label} non riuscita:\n{value}")
                self.status_left.config(text="")

        self._run_in_thread(worker, on_done)

    def extract_archive(self):
        """Estrae l'archivio selezionato (ZIP/TAR/7z/RAR, in background)."""
        sel = self.tree.selection()
        if len(sel) != 1:
            return
        zpath = sel[0]
        # toglie TUTTE le estensioni di archivio: "file.tar.gz" -> "file"
        dest_base = zpath
        for aext in (".tar.gz", ".tar.bz2", ".tar.xz", ".tgz", ".zip", ".7z",
                     ".rar", ".tar", ".gz", ".bz2", ".xz"):
            if dest_base.lower().endswith(aext):
                dest_base = dest_base[:-len(aext)]
                break
        dest = self._unique_path(dest_base)
        ext = os.path.splitext(zpath)[1].lower()
        fmt = "ZIP" if ext == ".zip" else ext.lstrip(".").upper()
        self.status_left.config(text=f"Estrazione {fmt} in corso…")
        self._progress_indeterminate()

        def worker(q=None):
            if ext == ".zip":
                return extract_zip(zpath, dest)
            if ext in (".tar", ".tgz", ".gz", ".bz2", ".xz"):
                return extract_tar(zpath, dest)
            if ext == ".7z":
                return extract_7z(zpath, dest)
            if ext == ".rar":
                return extract_rar(zpath, dest)
            return f"Formato non supportato: {ext}"

        def on_done(kind, value):
            self._progress_hide()
            if kind == "ok" and value is None:
                self.status_left.config(text=f"Estratto in {os.path.basename(dest)}")
                self.refresh()
            else:
                messagebox.showerror("Esplora File",
                                     f"Estrazione non riuscita:\n{value}")
                self.status_left.config(text="")

        self._run_in_thread(worker, on_done)

    def open_with(self):
        """Apre la finestra 'Apri con…' di Windows per il file selezionato."""
        sel = self.tree.selection()
        if len(sel) != 1 or os.path.isdir(sel[0]):
            return
        if not IS_WINDOWS:
            messagebox.showinfo("Esplora File",
                                "Apri con… è disponibile solo su Windows.")
            return
        try:
            # rundll32 shell32.dll,OpenAs_RunDLL apre la finestra "Apri con…"
            # anche per i file SENZA associazioni: os.startfile(path, "openas")
            # fallisce con WinError 1155 in quel caso
            subprocess.Popen(
                ["rundll32", "shell32.dll,OpenAs_RunDLL", sel[0]],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except OSError as exc:
            messagebox.showerror("Esplora File", f"Impossibile aprire con…:\n{exc}")

    # ------------------------------------------------------- confronto file
    def compare_selected(self):
        """Confronta due file di testo selezionati (diff unificato via difflib)."""
        sel = [p for p in self.tree.selection() if os.path.isfile(p)]
        if len(sel) != 2:
            messagebox.showinfo("Esplora File",
                                "Seleziona esattamente due file da confrontare.")
            return
        a, b = sel[0], sel[1]
        text_a = read_text_preview(a, DIFF_READ_LIMIT)
        text_b = read_text_preview(b, DIFF_READ_LIMIT)
        if text_a is None or text_b is None:
            messagebox.showerror("Esplora File",
                                 "Uno dei file è binario: "
                                 "impossibile confrontarlo come testo.")
            return
        self._show_diff_window(a, b, text_a, text_b)

    def _show_diff_window(self, path_a, path_b, text_a, text_b):
        """Finestra con il diff unificato colorato (rosso=rimosso, verde=aggiunto)."""
        diff = list(difflib.unified_diff(
            text_a.splitlines(), text_b.splitlines(),
            fromfile=os.path.basename(path_a),
            tofile=os.path.basename(path_b), lineterm=""))
        win = tk.Toplevel(self)
        win.title(f"Confronto — {os.path.basename(path_a)} ↔ {os.path.basename(path_b)}")
        win.geometry("900x560")
        win.transient(self)
        txt = tk.Text(win, wrap="none", font=("Consolas", 9), state="normal")
        vsb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        hsb = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        txt.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 4))
        vsb.grid(row=0, column=1, sticky="ns", pady=(8, 4))
        hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0))
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        txt.tag_configure("add", background="#c8f7c8")
        txt.tag_configure("del", background="#f7c8c8")
        txt.tag_configure("hunk", foreground="#888888")
        txt.tag_configure("meta", foreground="#3355aa", font=("Consolas", 9, "bold"))
        txt.insert("end", ("\n".join(diff) or "(nessuna differenza)") + "\n")
        idx = "1.0"
        for line in diff:
            if line.startswith("---") or line.startswith("+++"):
                txt.tag_add("meta", idx, f"{idx} lineend")
            elif line.startswith("@@"):
                txt.tag_add("hunk", idx, f"{idx} lineend")
            elif line.startswith("+"):
                txt.tag_add("add", idx, f"{idx} lineend")
            elif line.startswith("-"):
                txt.tag_add("del", idx, f"{idx} lineend")
            idx = txt.index(f"{idx} +1 line")
        txt.config(state="disabled")
        ttk.Button(win, text="Chiudi", command=win.destroy).grid(
            row=2, column=0, columnspan=2, sticky="e", padx=8, pady=(0, 8))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def compare_folders(self):
        """Confronta la cartella corrente con un'altra scelta dall'utente:
        mostra i nomi presenti solo nell'una, solo nell'altra e in comune."""
        a = self.current_path
        b = filedialog.askdirectory(title="Confronta con la cartella…", parent=self)
        if not b:
            return
        if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
            messagebox.showinfo("Esplora File", "Le due cartelle coincidono.")
            return
        try:
            set_a = set(os.listdir(a))
            set_b = set(os.listdir(b))
        except OSError as exc:
            messagebox.showerror("Esplora File",
                                 f"Impossibile leggere le cartelle:\n{exc}")
            return
        self._show_folders_diff(a, b, sorted(set_a - set_b),
                                sorted(set_b - set_a), sorted(set_a & set_b))

    def _show_folders_diff(self, a, b, only_a, only_b, common):
        """Finestra con l'esito del confronto tra due cartelle."""
        win = tk.Toplevel(self)
        win.title("Confronto cartelle")
        win.geometry("760x520")
        win.transient(self)
        txt = tk.Text(win, wrap="none", font=("Consolas", 9), state="normal")
        vsb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        hsb = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        txt.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 4))
        vsb.grid(row=0, column=1, sticky="ns", pady=(8, 4))
        hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0))
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        txt.tag_configure("head", foreground="#3355aa", font=("Consolas", 10, "bold"))
        txt.tag_configure("onlyA", foreground="#aa3355")
        txt.tag_configure("onlyB", foreground="#33803a")
        txt.tag_configure("common", foreground="#888888")
        txt.insert("end", f"Solo in {a}  ({len(only_a)})\n", "head")
        for n in only_a:
            txt.insert("end", f"  {n}\n", "onlyA")
        txt.insert("end", f"\nSolo in {b}  ({len(only_b)})\n", "head")
        for n in only_b:
            txt.insert("end", f"  {n}\n", "onlyB")
        txt.insert("end", f"\nIn comune  ({len(common)})\n", "head")
        for n in common:
            txt.insert("end", f"  {n}\n", "common")
        txt.config(state="disabled")
        ttk.Button(win, text="Chiudi", command=win.destroy).grid(
            row=2, column=0, columnspan=2, sticky="e", padx=8, pady=(0, 8))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def calc_folder_size(self):
        """Calcola in background la dimensione della cartella selezionata,
        mostrando il numero che sale mentre la scansione procede."""
        sel = self.tree.selection()
        if len(sel) != 1 or not os.path.isdir(sel[0]):
            return
        path = sel[0]
        token = self._size_token = self._size_token + 1
        name = os.path.basename(path)
        self.status_left.config(text=f"Calcolo dimensione di {name}…")
        cache = self.size_cache

        def worker(q):
            saved = [0]

            def progress(acc_size, acc_count):
                saved[0] += 1
                if saved[0] % 50 == 0:
                    cache.save()   # riprende in caso di interruzione
                # aggiorna la UI al massimo ogni 15 cartelle
                if saved[0] % 15 == 0:
                    q.put(("progress", (acc_size, acc_count)))

            size, count = compute_folder_size(path, cache, progress)
            cache.save()
            q.put(("progress", (size, count)))   # valore finale, poi il risultato
            return (size, count)

        def on_progress(value):
            if token != self._size_token:
                return
            acc_size, acc_count = value
            self.status_left.config(
                text=f"Calcolo dimensione di {name}… {format_size(acc_size)} · {acc_count} file")

        def on_done(kind, value):
            if kind != "ok":
                self.status_left.config(text="Calcolo dimensione non riuscito.")
                return
            if token != self._size_token:
                return
            size, count = value
            self.status_left.config(
                text=f"Dimensione {name}: {format_size(size)} · {count} file")
            # aggiorna la riga se visibile
            if self.tree.exists(path):
                vals = list(self.tree.item(path, "values"))
                vals[2] = format_size(size)
                self.tree.item(path, values=vals)
            self._refresh_cached_sizes()

        self._run_in_thread(worker, on_done, on_progress)

    def _refresh_cached_sizes(self):
        """Aggiorna la colonna Dimensioni della scheda attiva (se visibile)."""
        if self.filter_var.get().strip():
            return   # durante una ricerca le colonne sono diverse
        tab = self.active_tab
        if tab is not None and tab.frame.winfo_exists():
            tab.refresh(keep_selection=True)

    # ------------------------------------------------------- duplicati
    def find_duplicates(self):
        """Cerca in background i file duplicati nella cartella selezionata."""
        sel = self.tree.selection()
        if len(sel) != 1 or not os.path.isdir(sel[0]):
            return
        path = sel[0]
        token = self._dup_token = self._dup_token + 1
        name = os.path.basename(path) or path
        self.status_left.config(text=f"Cerca duplicati in {name}…")
        self._progress_indeterminate()

        def worker(q):
            def progress(count):
                if count % 200 == 0:
                    q.put(("progress", count))
            return find_duplicate_files(path, progress)

        def on_progress(count):
            if token != self._dup_token:
                return
            self.status_left.config(
                text=f"Cerca duplicati in {name}… {count} file analizzati")

        def on_done(kind, value):
            self._progress_hide()
            if kind != "ok":
                self.status_left.config(text="Ricerca duplicati non riuscita.")
                return
            if token != self._dup_token:
                return
            self.status_left.config(text="Ricerca duplicati completata.")
            self._show_duplicates(value)

        self._run_in_thread(worker, on_done, on_progress)

    # ------------------------------------------------------- statistiche
    def folder_statistics(self):
        """Calcola in background le statistiche della cartella selezionata
        (file, cartelle, byte totali, per estensione) e le mostra."""
        sel = self.tree.selection()
        if len(sel) != 1 or not os.path.isdir(sel[0]):
            return
        path = sel[0]
        name = os.path.basename(path) or path
        self.status_left.config(text=f"Calcolo statistiche di {name}…")
        self._progress_indeterminate()

        def worker(q):
            return folder_stats(path)

        def on_done(kind, value):
            self._progress_hide()
            if kind != "ok":
                self.status_left.config(text="Calcolo statistiche non riuscito.")
                return
            self.status_left.config(text=f"Statistiche di {name} calcolate.")
            self._show_folder_stats(value, name)

        self._run_in_thread(worker, on_done)

    def _show_folder_stats(self, stats, name):
        """Finestra con le statistiche della cartella e la tabella per tipo."""
        win = tk.Toplevel(self)
        win.title("Statistiche cartella")
        win.geometry("640x440")
        win.transient(self)

        ttk.Label(win, text=f"Statistiche — {name}",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        riepilogo = (f"{stats['files']} file · {stats['folders']} cartelle · "
                     f"{format_size(stats['bytes_total'])} in totale")
        if stats["files"]:
            riepilogo += (f"\nDimensione media: {format_size(stats['avg'])} per file · "
                          f"più grande: {os.path.basename(stats['largest'][0])} "
                          f"({format_size(stats['largest'][1])})")
        ttk.Label(win, text=riepilogo, font=("Segoe UI", 9)).pack(
            anchor="w", padx=12, pady=(0, 8))

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        tree = ttk.Treeview(frame, columns=("ext", "count", "size"), show="headings")
        tree.heading("ext", text="Tipo")
        tree.heading("count", text="File")
        tree.heading("size", text="Dimensione")
        tree.column("ext", width=280, anchor="w", stretch=True)
        tree.column("count", width=80, anchor="e", stretch=False)
        tree.column("size", width=120, anchor="e", stretch=False)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        if not stats["per_ext"]:
            ttk.Label(win, text="Cartella vuota.", font=("Segoe UI", 9)).pack(
                anchor="w", padx=12)
        for ext, (cnt, sz) in sorted(
                stats["per_ext"].items(), key=lambda kv: kv[1][1], reverse=True):
            tree.insert("", "end", values=(ext, cnt, format_size(sz)))
        ttk.Button(win, text="Chiudi", command=win.destroy).pack(
            anchor="e", padx=12, pady=(0, 10))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def _show_duplicates(self, groups):
        """Finestra con i gruppi di file duplicati (doppio clic per aprire)."""
        n_groups = len(groups)
        n_files = sum(len(ps) for _sz, ps in groups)
        win = tk.Toplevel(self)
        win.title("File duplicati")
        win.geometry("720x460")
        win.transient(self)
        summary = ttk.Label(win, text=f"File duplicati: {n_files} file in {n_groups} gruppi",
                            font=("Segoe UI", 10, "bold"))
        summary.pack(anchor="w", padx=12, pady=(10, 4))
        ttk.Label(win, text="Ogni gruppo contiene file con lo stesso contenuto "
                            "(doppio clic per aprire).",
                  font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 8))
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        tree = ttk.Treeview(frame, columns=("size", "path"), show="headings",
                            selectmode="extended")
        tree.heading("size", text="Dimensione")
        tree.heading("path", text="Percorso")
        tree.column("size", width=100, anchor="e", stretch=False)
        tree.column("path", width=580, anchor="w", stretch=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        if not groups:
            ttk.Label(win, text="Nessun file duplicato trovato.",
                      font=("Segoe UI", 9)).pack(anchor="w", padx=12)
        for sz, paths in groups:
            for p in paths:
                tree.insert("", "end", iid=p, values=(format_size(sz), p))
        tree.bind("<Double-1>", lambda e: self._open_dup(tree))
        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 10))
        if IS_WINDOWS:
            ttk.Button(btns, text="Sposta nel cestino",
                       command=lambda: self._delete_duplicates(win, tree, summary, True)
                       ).pack(side="left")
        ttk.Button(btns, text="Elimina definitivamente",
                   command=lambda: self._delete_duplicates(win, tree, summary, False)
                   ).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="Chiudi", command=win.destroy).pack(side="right")
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def _open_dup(self, tree):
        sel = tree.selection()
        if len(sel) == 1:
            self._open_file(sel[0])

    def _delete_duplicates(self, win, tree, summary, to_recycle):
        """Elimina i file duplicati selezionati (nel Cestino o definitivamente)."""
        sel = list(tree.selection())
        if not sel:
            return
        n = len(sel)
        label = "questo file" if n == 1 else f"questi {n} file"
        if to_recycle:
            if not messagebox.askyesno("Sposta nel cestino",
                                       f"Spostare {label} nel Cestino?",
                                       icon="question", parent=win):
                return
            if not send_to_recycle_bin(sel):
                messagebox.showerror("Esplora File",
                                     "Impossibile spostare i file nel Cestino.")
                return
        else:
            if not messagebox.askyesno("Elimina definitivamente",
                                       f"Eliminare DEFINITIVAMENTE {label}?\n\n"
                                       "Non finirà nel Cestino: l'operazione è irreversibile.",
                                       icon="warning", parent=win):
                return
            errors = []
            for p in sel:
                try:
                    os.remove(p)
                except OSError as exc:
                    errors.append(f"{os.path.basename(p)}: {exc}")
            if errors:
                messagebox.showerror("Esplora File",
                                     "Alcuni file non sono stati eliminati:\n\n"
                                     + "\n".join(errors))
        for p in sel:
            if tree.exists(p):
                tree.delete(p)
        summary.config(text=f"File duplicati: {len(tree.get_children())} file rimasti")
        self.refresh()

    # ------------------------------------------------------- presentazione (slideshow)
    def start_slideshow(self, folder=None):
        """Avvia una presentazione a tutto schermo delle immagini della cartella.

        Se folder è None usa la cartella corrente (o quella selezionata se è
        un'unica cartella). Frecce/click per avanzare, Spazio per mettere in
        pausa, Esc per chiudere.
        """
        if folder is None:
            folder = self.current_path
            sel = self.tree.selection()
            if len(sel) == 1 and os.path.isdir(sel[0]):
                folder = sel[0]
        exts = (".png", ".gif", ".ppm", ".pgm", ".pbm", ".jpg", ".jpeg",
                ".bmp", ".tif", ".tiff", ".ico", ".webp")
        images = []
        try:
            for name in sorted(os.listdir(folder)):
                if name.lower().endswith(exts):
                    p = os.path.join(folder, name)
                    if os.path.isfile(p):
                        images.append(p)
        except OSError:
            images = []
        if not images:
            messagebox.showinfo("Esplora File", "Nessuna immagine in questa cartella.")
            return
        self._stop_slideshow()   # chiude una eventuale presentazione già aperta
        win = tk.Toplevel(self)
        win.title("Slideshow")
        win.configure(bg="black")
        try:
            win.attributes("-fullscreen", True)
        except tk.TclError:
            win.geometry(f"{win.winfo_screenwidth()}x{win.winfo_screenheight()}")
        label = tk.Label(win, bg="black")
        label.pack(fill="both", expand=True)
        caption = tk.Label(win, bg="black", fg="white", font=("Segoe UI", 12))
        caption.pack(side="bottom", pady=10)
        self._slideshow = {"images": images, "idx": 0, "win": win,
                           "label": label, "caption": caption,
                           "photo": None, "after": None, "playing": True}
        win.bind("<Escape>", lambda e: self._stop_slideshow())
        win.bind("<Left>", lambda e: self._slide_show(-1))
        win.bind("<Right>", lambda e: self._slide_show(1))
        win.bind("<space>", lambda e: self._slide_toggle())
        win.bind("<Button-1>", lambda e: self._slide_show(1))
        self._slide_render()
        self._slide_schedule()

    def _slide_show(self, delta):
        """Avanza/retrocede di 'delta' immagini e riavvia il timer."""
        s = self._slideshow
        if s is None:
            return
        s["idx"] = (s["idx"] + delta) % len(s["images"])
        self._slide_render()
        self._slide_schedule()

    def _slide_render(self):
        """Carica e mostra l'immagine corrente (adattandola allo schermo)."""
        s = self._slideshow
        if s is None:
            return
        path = s["images"][s["idx"]]
        photo = None
        try:
            if os.path.getsize(path) <= 12 * 1024 * 1024:
                photo = tk.PhotoImage(file=path)
        except (tk.TclError, OSError):
            photo = None
        if photo is None:
            dec = decode_image_rgb(path)
            if dec is not None:
                w, h, rgb = dec
                if w * h <= PREVIEW_IMAGE_MAX_PIXELS:
                    try:
                        photo = photo_from_rgb(w, h, rgb)
                    except (tk.TclError, MemoryError):
                        photo = None
        if photo is not None:
            sw = max(100, s["win"].winfo_screenwidth() - 40)
            sh = max(100, s["win"].winfo_screenheight() - 80)
            if photo.width() > sw or photo.height() > sh:
                sx = max(1, math.ceil(photo.width() / sw))
                sy = max(1, math.ceil(photo.height() / sh))
                step = max(sx, sy)
                photo = photo.subsample(step, step)
            s["photo"] = photo
            s["label"].config(image=photo, text="")
        else:
            s["photo"] = None
            s["label"].config(image="", text="Formato non supportato")
        s["caption"].config(
            text=f"{s['idx'] + 1}/{len(s['images'])} — {os.path.basename(path)}")

    def _slide_advance(self):
        """Callback del timer di avanzamento automatico."""
        s = self._slideshow
        if s is None:
            return
        s["after"] = None   # il timer è già scattato: azzera il riferimento
        self._slide_show(1)

    def _slide_schedule(self):
        """Riprogramma l'avanzamento automatico (3 s) se in riproduzione."""
        s = self._slideshow
        if s is None or not s.get("playing"):
            return
        if s["after"] is not None:
            try:
                self.after_cancel(s["after"])
            except tk.TclError:
                pass
        s["after"] = self.after(3000, self._slide_advance)

    def _slide_toggle(self):
        """Mette in pausa / riprende l'avanzamento automatico."""
        s = self._slideshow
        if s is None:
            return
        s["playing"] = not s["playing"]
        if s["playing"]:
            self._slide_schedule()
        elif s["after"] is not None:
            try:
                self.after_cancel(s["after"])
            except tk.TclError:
                pass
            s["after"] = None

    def _stop_slideshow(self):
        """Chiude la presentazione e cancella l'eventuale timer."""
        s = self._slideshow
        if s is not None:
            if s["after"] is not None:
                try:
                    self.after_cancel(s["after"])
                except tk.TclError:
                    pass
            if s["win"] is not None and s["win"].winfo_exists():
                s["win"].destroy()
        self._slideshow = None

    # ------------------------------------------- dimensioni automatiche
    def _auto_start(self):
        """Attiva l'auto-size dopo l'avvio e accoda le cartelle visibili."""
        if not self._auto_enabled:
            return
        self._auto_ready = True
        self._queue_auto_sizes()

    def _queue_auto_sizes(self):
        """Ricostruisce la coda con le cartelle visibili senza dimensione.

        Le scansioni delle viste precedenti vengono scartate (o interrotte),
        così le dimensioni delle cartelle che stai guardando compaiono subito
        senza aspettare cartelle enormi rimaste in coda (es. AppData).
        """
        if not (self._auto_ready and self._auto_enabled):
            return
        if self.filter_var.get().strip():
            return   # durante la ricerca le colonne sono diverse
        tab = self.active_tab
        if tab is None or tab.search_active:
            return
        wanted = []
        seen = set()
        for entry in tab.entries:
            path, _name, is_dir, _size, _mtime, _rel = entry
            if not is_dir:
                continue
            key = os.path.normcase(path)
            if key in seen:
                continue
            seen.add(key)
            if self.size_cache.cached_size(path) is not None:
                continue   # già noto e valido
            wanted.append(path)
            if len(wanted) >= 50:
                break   # limite prudente: non accumulare scansioni
        self._auto_queue = wanted
        # se la scansione in corso non riguarda più la vista, interrompila
        if self._auto_busy and self._auto_current:
            self._auto_cancel = (os.path.normcase(self._auto_current)
                                 not in seen)
        self._auto_pump()

    def _auto_pump(self):
        """Avvia una scansione automatica se non ce n'è già una in corso."""
        if self._auto_busy or not self._auto_queue or not self._auto_ready:
            return
        path = self._auto_queue.pop(0)
        self._auto_busy = True
        self._auto_current = path
        self._auto_cancel = False
        cache = self.size_cache

        def worker(q=None):
            saved = [0]

            def progress(acc_size, acc_count):
                saved[0] += 1
                if saved[0] % 50 == 0:
                    cache.save()   # riprende in caso di interruzione
                if self._auto_cancel:
                    raise _ScanCancelled()

            try:
                size, count = compute_folder_size(path, cache, progress)
            except _ScanCancelled:
                cache.save()
                return None   # vista cambiata: cartella non più visibile
            cache.save()
            return (path, size, count)

        def on_done(kind, value):
            self._auto_busy = False
            self._auto_current = None
            if kind == "ok" and value is not None:
                done_path, size, _count = value
                self._apply_auto_size(done_path, size)
            self._auto_pump()   # passa alla cartella successiva

        self._run_in_thread(worker, on_done)

    def _apply_auto_size(self, path, size):
        """Aggiorna la cella Dimensioni della cartella (se ancora visibile)."""
        tab = self.active_tab
        if tab is None or tab.search_active:
            return
        try:
            if not tab.tree.exists(path):
                return   # non è più visibile: il valore resta in cache
            vals = list(tab.tree.item(path, "values"))
            vals[2] = format_size(size)
            tab.tree.item(path, values=vals)
        except tk.TclError:
            pass

    # ------------------------------------------------------- hash multipli
    def hash_selected(self):
        """Calcola MD5 e SHA-256 dei file selezionati (in background) e
        li mostra in una tabella copiabile."""
        files = [p for p in self.tree.selection() if os.path.isfile(p)]
        if not files:
            messagebox.showinfo("Esplora File", "Seleziona uno o più file.")
            return
        total = len(files)
        self._progress_show(total)
        self.status_left.config(text=f"Calcolo hash… (0/{total})")

        def worker(q):
            results = []
            for i, p in enumerate(files):
                h = file_hashes(p)
                results.append((p, h[0] if h else None, h[1] if h else None))
                q.put(("progress", i + 1))
            return results

        def on_progress(done):
            done = min(done, total)
            self.status_left.config(text=f"Calcolo hash… ({done}/{total})")
            self._progress_set(done)

        def on_done(kind, value):
            self._progress_hide()
            if kind != "ok":
                self.status_left.config(text="Calcolo hash non riuscito.")
                return
            self.status_left.config(text=f"Hash calcolati per {total} file.")
            self._show_hashes(value)

        self._run_in_thread(worker, on_done, on_progress=on_progress)

    def _show_hashes(self, results):
        """Finestra con la tabella file / MD5 / SHA-256."""
        win = tk.Toplevel(self)
        win.title("Hash file")
        win.geometry("860x420")
        win.transient(self)
        ttk.Label(win, text=f"Hash MD5 e SHA-256 di {len(results)} file",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        tree = ttk.Treeview(frame, columns=("file", "md5", "sha"), show="headings")
        tree.heading("file", text="File")
        tree.heading("md5", text="MD5")
        tree.heading("sha", text="SHA-256")
        tree.column("file", width=260, anchor="w", stretch=True)
        tree.column("md5", width=250, anchor="w", stretch=False)
        tree.column("sha", width=280, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        for p, md5, sha in results:
            tree.insert("", "end", values=(os.path.basename(p), md5 or "—", sha or "—"))
        ttk.Button(win, text="Chiudi", command=win.destroy).pack(
            anchor="e", padx=12, pady=(0, 10))
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    # ------------------------------------------------------- proprietà
    def _show_multi_properties(self, sel):
        """Proprietà aggregate di più elementi selezionati."""
        files = folders = 0
        total = 0
        has_uncached = False
        for p in sel:
            try:
                if os.path.isdir(p):
                    folders += 1
                    cached = self.size_cache.cached_size(p)
                    if cached is not None:
                        total += cached
                    else:
                        has_uncached = True
                elif os.path.isfile(p):
                    files += 1
                    total += os.path.getsize(p)
            except OSError:
                pass
        size_str = format_size(total)
        if has_uncached:
            size_str += "  (cartelle senza dimensione escluse)"
        win = tk.Toplevel(self)
        win.title("Proprietà")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text=f"Proprietà — {len(sel)} elementi",
                  font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 8))
        rows = [
            ("Elementi:", str(len(sel))),
            ("File:", str(files)),
            ("Cartelle:", str(folders)),
            ("Dimensione totale:", size_str),
        ]
        for i, (k, v) in enumerate(rows, start=1):
            ttk.Label(win, text=k, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, sticky="ne", padx=(14, 6), pady=2)
            ttk.Label(win, text=v).grid(
                row=i, column=1, sticky="w", padx=(0, 14), pady=2)
        ttk.Button(win, text="Chiudi", command=win.destroy).grid(
            row=len(rows) + 1, column=1, sticky="e", padx=14, pady=12)
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

    def show_properties(self):
        sel = self.tree.selection()
        if not sel:
            return
        if len(sel) > 1:
            self._show_multi_properties(list(sel))
            return
        path = sel[0]
        try:
            # follow_symlinks=True (default): un symlink rotto produce un
            # OSError qui, gestito dal messaggio qui sotto (niente crash)
            st = os.stat(path, follow_symlinks=True)
        except OSError as exc:
            messagebox.showerror("Esplora File", str(exc))
            return

        win = tk.Toplevel(self)
        win.title("Proprietà")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        title = "Proprietà — " + os.path.basename(path)
        ttk.Label(win, text=title, font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 8))

        is_dir = os.path.isdir(path)
        entry = (path, os.path.basename(path), False, st.st_size, st.st_mtime, "")
        # per le cartelle usa la dimensione in cache (come la colonna
        # dell'albero): "—" solo se non è mai stata calcolata
        cached = self.size_cache.cached_size(path) if is_dir else None
        size_str = format_size(cached) if cached is not None else (
            "—" if is_dir else format_size(st.st_size))
        rows = [
            ("Nome:", os.path.basename(path)),
            ("Tipo:", "Cartella di file" if is_dir else self.active_tab._type_label(entry)),
            ("Posizione:", os.path.dirname(path)),
            ("Dimensione:", size_str),
            ("Creato:", format_date(st.st_ctime)),
            ("Ultima modifica:", format_date(st.st_mtime)),
            ("Ultimo accesso:", format_date(st.st_atime)),
        ]
        if not is_dir:
            rows.append(("MD5:", "Calcolo in corso…"))
            rows.append(("SHA-256:", "Calcolo in corso…"))

        # I valori sono campi selezionabili e copiabili (non modificabili)
        value_widgets = []
        md5_widget = sha_widget = None
        for i, (k, v) in enumerate(rows, start=1):
            ttk.Label(win, text=k, font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, sticky="ne", padx=(14, 6), pady=2)
            w = readonly_text_widget(win, v)
            w.grid(row=i, column=1, sticky="w", padx=(0, 14), pady=2)
            self._bind_copy_menu(w)
            value_widgets.append((k, w))
            if k == "MD5:":
                md5_widget = w
            elif k == "SHA-256:":
                sha_widget = w

        last_row = len(rows) + 1
        ttk.Button(win, text="Copia tutto",
                   command=lambda: self._copy_properties(win, value_widgets,
                                                         os.path.basename(path))).grid(
            row=last_row, column=0, sticky="w", padx=14, pady=12)
        ttk.Button(win, text="Chiudi", command=win.destroy).grid(
            row=last_row, column=1, sticky="e", padx=14, pady=12)

        # Ctrl+C fuori dai campi di testo copia tutte le proprietà
        win.bind("<Control-c>",
                 lambda e: self._copy_properties(win, value_widgets,
                                                 os.path.basename(path)))

        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")

        # Hash MD5/SHA-256 calcolati in background: la finestra resta fluida
        if not is_dir:
            self._compute_property_hashes(win, path, md5_widget, sha_widget)

    # ------------------------------------------------ proprietà: copia e hash
    def _bind_copy_menu(self, widget):
        """Menu contestuale 'Copia' per i valori selezionabili."""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Copia",
                         command=lambda: self._copy_widget_value(widget))
        widget.bind("<Button-3>",
                    lambda e: menu.tk_popup(e.x_root, e.y_root))

    def _copy_widget_value(self, widget):
        """Copia la selezione di un valore (o l'intero valore) negli appunti."""
        try:
            text = widget.get("sel.first", "sel.last")
        except tk.TclError:
            text = widget.get("1.0", "end-1c")
        widget.clipboard_clear()
        widget.clipboard_append(text)

    def _copy_properties(self, win, value_widgets, name):
        """Copia tutte le proprietà (con i valori correnti) negli appunti."""
        lines = [f"Proprietà — {name}", ""]
        for label, w in value_widgets:
            lines.append(f"{label} {w.get('1.0', 'end-1c')}")
        win.clipboard_clear()
        win.clipboard_append("\n".join(lines))

    def _compute_property_hashes(self, win, path, md5_widget, sha_widget):
        """Calcola MD5 e SHA-256 in background aggiornando le righe della finestra."""

        def _set(widget, text):
            try:
                if not win.winfo_exists():
                    return
                widget.configure(state="normal")
                widget.delete("1.0", "end")
                widget.insert("1.0", text)
            except tk.TclError:
                pass   # finestra chiusa durante il calcolo

        def _progress(value):
            if not win.winfo_exists():
                return
            letti, totale = value
            if totale:
                _set(md5_widget, f"Calcolo in corso… ({letti * 100 // totale}%)")
            else:
                _set(md5_widget, "Calcolo in corso…")

        def _done(kind, value):
            if not win.winfo_exists():
                return
            if kind == "ok" and value is not None:
                md5hex, sha256hex = value
                _set(md5_widget, md5hex)
                _set(sha_widget, sha256hex)
            else:
                _set(md5_widget, "Non disponibile")
                _set(sha_widget, "Non disponibile")

        def _worker(q):
            return file_hashes(path,
                               progress_cb=lambda l, t: q.put(("progress", (l, t))))

        self._run_in_thread(_worker, _done, on_progress=_progress)

    def show_about(self):
        messagebox.showinfo(
            "Esplora File",
            "Esplora File 1.41\n\n"
            "Un file explorer in stile Windows Explorer.\n"
            "Realizzato in Python con la sola libreria standard (tkinter).\n\n"
            "Schede (stile browser):\n"
            "  Ctrl+T                       Nuova scheda\n"
            "  Ctrl+W                       Chiudi scheda\n"
            "  Ctrl+Maiusc+T                Riapri scheda chiusa\n"
            "  Ctrl+Tab / Ctrl+Maiusc+Tab   Scheda successiva / precedente\n"
            "  Clic centrale / × su scheda   Chiudi\n"
            "  Doppio clic su spazio vuoto   Nuova scheda\n\n"
            "Ricerca:\n"
            "  Il filtro cerca in TUTTE le sottocartelle (in background),\n"
            "  mostrando il percorso di ogni risultato. Attiva 'Contenuto'\n"
            "  per cercare il testo DENTRO i file (grep).\n\n"
            "Archivi e dimensioni:\n"
            "  Comprimi in ZIP/7z/RAR ed Estrai ZIP, 7z, RAR e TAR dal menu\n"
            "  contestuale (in background); le dimensioni delle cartelle\n"
            "  vengono ricordate tra le sessioni.\n\n"
            "Novità in 1.41:\n"
            "  Copia percorso: ora copia i percorsi di TUTTI gli elementi\n"
            "  selezionati (uno per riga), non più solo il primo.\n\n"
            "Novità in 1.40:\n"
            "  Seleziona per regex (menu Modifica): seleziona gli elementi\n"
            "  il cui nome corrisponde a un'espressione regolare.\n\n"
            "Novità in 1.39:\n"
            "  Filtro per estensione: un menu a tendina nella barra mostra\n"
            "  solo i file con l'estensione scelta (le cartelle restano).\n\n"
            "Novità in 1.38:\n"
            "  Svuota cestino (menu File, Windows): svuota completamente il\n"
            "  Cestino con conferma esplicita.\n\n"
            "Novità in 1.37:\n"
            "  Sostituisci nei nomi: rinomina la selezione sostituendo una\n"
            "  sottostringa nei nomi ('trova' → 'sostituisci') con anteprima.\n\n"
            "Novità in 1.36:\n"
            "  Nuova finestra (menu File): apre una seconda istanza dell'app\n"
            "  sulla cartella corrente (funziona anche da eseguibile).\n\n"
            "Novità in 1.35:\n"
            "  Confronta cartelle (menu Navigazione): mostra i nomi presenti\n"
            "  solo in una cartella, solo nell'altra e in comune.\n\n"
            "Novità in 1.34:\n"
            "  Seleziona per data (menu Modifica): seleziona i file\n"
            "  modificati negli ultimi N giorni.\n\n"
            "Novità in 1.33:\n"
            "  Spazio libero: la barra di stato mostra lo spazio libero e\n"
            "  totale dell'unità della cartella corrente.\n\n"
            "Novità in 1.32:\n"
            "  Seleziona per dimensione (menu Modifica): seleziona i file\n"
            "  più grandi di una soglia ('10 MB', '500 KB', '2 GB', ...).\n\n"
            "Novità in 1.31:\n"
            "  Apri tutti i selezionati: dal menu contestuale apre con il\n"
            "  programma predefinito tutti i file selezionati (con conferma\n"
            "  oltre i 15).\n\n"
            "Novità in 1.30:\n"
            "  Dimensione della selezione: la barra di stato mostra la\n"
            "  dimensione totale degli elementi selezionati.\n\n"
            "Novità in 1.29:\n"
            "  Indicatore di ordinamento: una freccia (▲/▼) sull'intestazione\n"
            "  della colonna mostra direzione e colonna attualmente ordinate.\n\n"
            "Novità in 1.28:\n"
            "  Esporta elenco (CSV): dal menu File salva l'elenco visualizzato\n"
            "  (cartella o risultati di ricerca) in CSV, compatibile con Excel.\n\n"
            "Novità in 1.27:\n"
            "  Copia nome: dal menu contestuale copia i soli nomi degli\n"
            "  elementi selezionati (uno per riga), senza il percorso.\n\n"
            "Novità in 1.26:\n"
            "  Ricerca con espressioni regolari: la casella 'Regex' accanto a\n"
            "  'Contenuto' interpreta il filtro come regex (es. \\d{4}) nel\n"
            "  nome o nel contenuto; un pattern non valido viene segnalato.\n\n"
            "Novità in 1.25:\n"
            "  Ordinamento naturale: i nomi vengono ordinati come un umano\n"
            "  ('file2' prima di 'file10') nelle colonne Nome e Tipo.\n\n"
            "Novità in 1.24:\n"
            "  Riapri scheda chiusa: Ctrl+Maiusc+T (o menu File / menu\n"
            "  contestuale delle schede) riapre l'ultima scheda chiusa,\n"
            "  ripristinandone il percorso e la posizione.\n\n"
            "Novità in 1.23:\n"
            "  Calcola hash: MD5 e SHA-256 di più file selezionati, in\n"
            "  background, mostrati in una tabella (menu contestuale).\n\n"
            "Novità in 1.22:\n"
            "  Clic centrale (rotella) su una cartella per aprirla in una\n"
            "  nuova scheda, come nei browser.\n\n"
            "Novità in 1.21:\n"
            "  Scorciatoie di focus: Ctrl+F sulla ricerca, Ctrl+L sulla\n"
            "  barra indirizzi (come nei browser).\n\n"
            "Novità in 1.20:\n"
            "  Seleziona per modello (Ctrl+Maiusc+M): scegli un pattern\n"
            "  wildcard (es. *.txt) per selezionare gli elementi corrispondenti.\n\n"
            "Novità in 1.19:\n"
            "  Cartelle recenti: il menu Navigazione ricorda le ultime\n"
            "  cartelle visitate (persistite tra le sessioni).\n\n"
            "Novità in 1.18:\n"
            "  Crea collegamento (.lnk) al file o alla cartella selezionata\n"
            "  (menu File o menu contestuale, solo Windows).\n\n"
            "Novità in 1.17:\n"
            "  Proprietà di più elementi selezionati: numero di file,\n"
            "  cartelle e dimensione totale aggregata.\n\n"
            "Novità in 1.16:\n"
            "  Barra di avanzamento animata (indeterminata) anche per\n"
            "  compressione, estrazione, ricerca duplicati e statistiche.\n\n"
            "Novità in 1.15:\n"
            "  Selezione avanzata: 'Seleziona stesso tipo' (Ctrl+Maiusc+A)\n"
            "  e 'Inverti selezione' (Ctrl+I) nel menu Modifica.\n\n"
            "Novità in 1.14:\n"
            "  Apri in nuova scheda: Ctrl+doppio clic su una cartella (o\n"
            "  voce nel menu contestuale) la apre in una scheda separata.\n\n"
            "Novità in 1.13:\n"
            "  Duplica file/cartelle nella stessa cartella ('nome (copia).ext')\n"
            "  dal menu Modifica o dal menu contestuale, in background con\n"
            "  barra di avanzamento.\n\n"
            "Novità in 1.12:\n"
            "  Presentazione (slideshow) a tutto schermo delle immagini:\n"
            "  frecce/click per avanzare, Spazio per mettere in pausa,\n"
            "  Esc per chiudere (Visualizza > Presentazione o menu\n"
            "  contestuale).\n\n"
            "Novità in 1.11:\n"
            "  Ricerca nel contenuto (grep): spunta 'Contenuto' accanto al\n"
            "  filtro per trovare i file che CONTENGONO il testo cercato.\n"
            "  Preferiti nella barra laterale, confronto di due file (diff\n"
            "  colorato), eliminazione dei duplicati selezionati, \"Apri in\n"
            "  Esplora Risorse\" e rinomina in batch (modello {n}/{name}/{ext}).\n\n"
            "Novità in 1.10:\n"
            "  Estrazione di archivi 7z e RAR (via 7-Zip o py7zr) e\n"
            "  compressione in 7z; per creare RAR serve WinRAR (formato\n"
            "  proprietario). Elimina ora cancella DEFINITIVAMENTE (con\n"
            "  conferma); per il Cestino usa 'Sposta nel cestino'.\n"
            "  'Distruggi (sovrascrivi)' cancella dopo aver sovrascritto il\n"
            "  contenuto con byte casuali (3 passaggi, fsync).\n\n"
            "Novità in 1.9:\n"
            "  Proprietà con hash MD5 e SHA-256 (calcolati in background)\n"
            "  e valori selezionabili e copiabili (Ctrl+C, tasto destro\n"
            "  o \"Copia tutto\"); anteprima dei metadati audio/video\n"
            "  (mutagen, libreria opzionale compatibile con Nuitka).\n\n"
            "Novità in 1.8:\n"
            "  Risoluzione dell'anteprima immagini (Visualizza): Bassa apre\n"
            "  adattata, Media al 100%, Alta al 200% (dettaglio);\n"
            "  anteprima del contenuto degli archivi ZIP/TAR e di molti\n"
            "  formati testuali nuovi. Menu Crea (1.7), Office (1.6).\n\n"
            "Scorciatoie:\n"
            "  Invio / doppio clic  Apri\n"
            "  Alt+\u2190 / Alt+\u2192   Indietro / Avanti\n"
            "  Alt+\u2191            Cartella superiore\n"
            "  F2                 Rinomina\n"
            "  Canc               Elimina\n"
            "  Ctrl+C / X / V     Copia / Taglia / Incolla\n"
            "  Ctrl+A             Seleziona tutto\n"
            "  Ctrl+F             Cerca (focus sul filtro)\n"
            "  Ctrl+L             Barra indirizzi\n"
            "  F5                 Aggiorna\n\n"
            "Drag & drop:\n"
            "  Trascina file/cartelle su una cartella, sulla barra laterale\n"
            "  o su un'altra scheda per spostarli. Tieni premuto Ctrl\n"
            "  mentre rilasci per copiare invece di spostare.")

    # ------------------------------------------------------- sessione
    def _save_session(self):
        """Salva geometria, colonne, ordinamento e schede per la prossima sessione."""
        s = load_settings()
        s["window_geometry"] = self.geometry()
        s["show_hidden"] = bool(self.show_hidden)
        s["show_details"] = bool(self.show_details)
        s["preview_panel"] = bool(self.preview_var.get())
        tab = self.active_tab
        if tab is not None:
            widths = {}
            for col in ("name", "type", "size", "date", "path"):
                try:
                    widths[col] = int(tab.tree.column(col, "width"))
                except tk.TclError:
                    pass
            s["column_widths"] = widths
            s["sort_col"] = tab.sort_col
            s["sort_desc"] = bool(tab.sort_desc)
        s["tabs"] = [t.current_path for t in self.tabs]
        s["active_tab_index"] = self.tabbar.selected
        save_settings(s)

    def _on_close(self):
        """Salva la sessione e chiude l'applicazione."""
        self._save_session()
        self.destroy()

    # ------------------------------------------------------- stato
    def _update_status(self):
        if not self.tabs:
            return
        total = len(self.tree.get_children())
        selected = len(self.tree.selection())
        size_info = ""
        if selected:
            size_info = f"   |   {format_size(self._selection_size())}"
        if self.clipboard["paths"]:
            clip_mode = "copia" if self.clipboard["mode"] == "copy" else "taglia"
            clip_info = f"   |   Negli appunti: {len(self.clipboard['paths'])} elemento/i ({clip_mode})"
        else:
            clip_info = ""
        self.status_left.config(
            text=f"{total} elemento/i   |   {selected} selezionato/i{size_info}{clip_info}")
        right = self.current_path
        if self._free_space:
            right += f"   |   {self._free_space}"
        self.status_right.config(text=right)

    def _selection_size(self):
        """Dimensione totale della selezione (file: reale; cartelle: cache)."""
        tab = self.active_tab
        if tab is None:
            return 0
        by_path = {e[0]: e for e in tab.entries}
        total = 0
        for p in self.tree.selection():
            entry = by_path.get(p)
            if entry is None:
                continue
            is_dir, size = entry[2], entry[3]
            if is_dir:
                total += self.size_cache.cached_size(p) or 0
            else:
                total += size
        return total

    def _update_free_space(self):
        """Spazio libero/totale sull'unità della cartella corrente (per lo stato)."""
        try:
            usage = shutil.disk_usage(self.current_path)
            self._free_space = (f"{format_size(usage.free)} liberi "
                                f"/ {format_size(usage.total)}")
        except OSError:
            self._free_space = ""


# --------------------------------------------------------------------------
def main():
    app = FileExplorer()
    if len(sys.argv) > 1:
        start = os.path.abspath(sys.argv[1])
        if os.path.isdir(start):
            app.navigate_to(start, record=True)
    app.mainloop()


if __name__ == "__main__":
    main()
