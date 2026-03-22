# -*- coding: cp1252 -*-
import os
import time
import shutil
import subprocess
import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime

try:
    import win32com.client
    import win32print
    import win32con
    _COM_AVAILABLE = True
except Exception:
    win32com = None
    win32print = None
    _COM_AVAILABLE = False

# ------------------- Pfade -------------------
ROOT_PATH    = r"x:\temp_share\to_print"
SOURCE_PATH  = os.path.join(ROOT_PATH, "input")
TEMP_PATH    = os.path.join(ROOT_PATH, "temp")
ARCHIVE_PATH = os.path.join(ROOT_PATH, "archive")
LOG_PATH     = r"x:\logs\files\to_print"

PDFXCHANGE_PATH = r"C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"
IRFANVIEW_PATH  = r"C:\Programme\IrfanView\i_view64.exe"
NOTEPAD_PATH    = r"C:\Windows\System32\notepad.exe"
WORD_PATH       = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
EXCEL_PATH      = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

WAIT_INTERVAL  = 1
STABLE_TIME    = 3
MAX_ATTEMPTS   = 20
LOG_LEVEL      = logging.DEBUG
PRINT_TIMEOUT  = 60  # Sekunden maximale Wartezeit auf Druckauftrag

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
TEXT_EXTS  = {".txt", ".conf"}
WORD_EXTS  = {".doc", ".docx"}
EXCEL_EXTS = {".xls", ".xlsx"}

# ------------------- Verzeichnisse erstellen -------------------
for path in [SOURCE_PATH, TEMP_PATH, ARCHIVE_PATH, LOG_PATH]:
    os.makedirs(path, exist_ok=True)

# ------------------- Logging -------------------
logger = logging.getLogger("PrintWatcher")
logger.setLevel(LOG_LEVEL)
logger.propagate = False

logfile = os.path.join(LOG_PATH, "print.log")
handler = TimedRotatingFileHandler(
    logfile, when="midnight", interval=1, backupCount=7, encoding="utf-8"
)
handler.setLevel(LOG_LEVEL)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S")
handler.setFormatter(formatter)

logger.handlers.clear()
logger.addHandler(handler)

# ------------------- Hilfsfunktionen -------------------
def is_file_locked(filepath):
    try:
        with open(filepath, "a"):
            return False
    except IOError:
        return True

def wait_for_file_stable(filepath):
    last_size = -1
    stable_counter = 0
    attempts = 0
    while attempts < MAX_ATTEMPTS:
        if not is_file_locked(filepath):
            try:
                size = os.path.getsize(filepath)
            except FileNotFoundError:
                return False
            if size == last_size and size > 0:
                stable_counter += WAIT_INTERVAL
                if stable_counter >= STABLE_TIME:
                    return True
            else:
                stable_counter = 0
                last_size = size
        time.sleep(WAIT_INTERVAL)
        attempts += 1
    return False

def wait_for_print_job(filename, timeout=PRINT_TIMEOUT):
    """
    Wartet, bis ein Druckauftrag für die angegebene Datei in allen Druckern abgeschlossen ist.
    """
    if not _COM_AVAILABLE or not win32print:
        return True  # Fallback, falls win32print nicht verfügbar
    start_time = time.time()
    while time.time() - start_time < timeout:
        job_done = True
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        for flags, description, name, comment in printers:
            try:
                hprinter = win32print.OpenPrinter(name)
                jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
                for job in jobs:
                    if filename.lower() in job["pDocument"].lower():
                        job_done = False
                        break
                win32print.ClosePrinter(hprinter)
            except Exception:
                continue
            if not job_done:
                break
        if job_done:
            return True
        time.sleep(1)
    return False

def print_pdf(path):
    subprocess.run([PDFXCHANGE_PATH, "/silent", "/print", path], check=True)
    return True

def print_image(path):
    subprocess.run([IRFANVIEW_PATH, path, "/print"], check=True)
    return True

def print_text(path):
    subprocess.run([NOTEPAD_PATH, "/p", path], check=True)
    return True

def print_word(path):
    if _COM_AVAILABLE:
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(path)
            doc.PrintOut(Background=False)
            doc.Close(False)
            return True
        except Exception as e:
            logger.exception(f"Word COM Fehler: {e}")
        finally:
            if word:
                word.Quit()
    subprocess.run([WORD_PATH, "/q", "/n", "/mFilePrintDefault", "/mFileExit", path], check=True)
    return True

def print_excel(path):
    if _COM_AVAILABLE:
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(path)
            wb.PrintOut()
            wb.Close(False)
            return True
        except Exception as e:
            logger.exception(f"Excel COM Fehler: {e}")
        finally:
            if excel:
                excel.Quit()
    subprocess.run([EXCEL_PATH, "/q", "/n", path, "/mFilePrintDefault"], check=True)
    return True

def process_file(filepath):
    filename = os.path.basename(filepath)
    ext = os.path.splitext(filename)[1].lower()
    lock_file = filepath + ".lock"

    if os.path.exists(lock_file):
        logger.info(f"Datei wird übersprungen, Lock vorhanden: {filename}")
        return

    open(lock_file, "w").close()
    try:
        if not wait_for_file_stable(filepath):
            logger.warning(f"Datei nicht stabil: {filename}")
            return

        temp_file = os.path.join(TEMP_PATH, filename)
        shutil.move(filepath, temp_file)

        printed = False
        if ext == ".pdf":
            printed = print_pdf(temp_file)
        elif ext in IMAGE_EXTS:
            printed = print_image(temp_file)
        elif ext in TEXT_EXTS:
            printed = print_text(temp_file)
        elif ext in WORD_EXTS:
            printed = print_word(temp_file)
        elif ext in EXCEL_EXTS:
            printed = print_excel(temp_file)

        if printed:
            logger.info(f"Datei gedruckt: {filename}")
            if wait_for_print_job(filename, timeout=PRINT_TIMEOUT):
                now = datetime.now()
                date_folder = os.path.join(ARCHIVE_PATH, now.strftime("%Y-%m-%d"))
                time_folder = os.path.join(date_folder, now.strftime("%H-%M-%S"))
                os.makedirs(time_folder, exist_ok=True)
                shutil.move(temp_file, os.path.join(time_folder, filename))
                logger.info(f"Datei archiviert: {filename}")
            else:
                logger.warning(f"Druckauftrag nicht abgeschlossen nach {PRINT_TIMEOUT}s: {filename}")
    except Exception as e:
        logger.exception(f"Fehler bei Datei {filename}: {e}")
    finally:
        if os.path.exists(lock_file):
            os.remove(lock_file)

# ------------------- Main -------------------
if __name__ == "__main__":
    logger.info("AutomaticPrint gestartet (Task Scheduler Run)")
    found = False
    for entry in list(os.scandir(SOURCE_PATH)):
        if entry.is_file():
            found = True
            process_file(entry.path)
    if not found:
        logger.info("Keine neuen Dateien gefunden.")
    logger.info("AutomaticPrint beendet")
