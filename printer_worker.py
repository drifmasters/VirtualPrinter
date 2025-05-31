# -*- coding: utf-8 -*-

import os
import sys
import time
import threading
import zipfile
import tkinter as tk
from tkinter import messagebox
import win32event
import win32service
import win32serviceutil
import win32print
import xml.etree.ElementTree as ET

WATCH_FOLDER = r"C:\VM_PRINTERS\VIRT1"
REAL_PRINTER_NAME = "Microsoft XPS Document Writer"
SERVICE_NAME = "PyVirtualPrinterWorker"
SERVICE_DISPLAY_NAME = "Python VirtualPrinter Worker Service"

def get_page_size(xps_path):
    sizes = {
        (816, 1056): "Letter",
        (794, 1123): "A4",
        (1123, 1587): "A3",
        (1587, 2245): "A2",
        (559, 794): "A5",
    }
    try:
        with zipfile.ZipFile(xps_path, 'r') as z:
            for name in z.namelist():
                if name.endswith('.fpage'):
                    xml_content = z.read(name)
                    root = ET.fromstring(xml_content)
                    width = round(float(root.attrib.get('Width', '794')))
                    height = round(float(root.attrib.get('Height', '1123')))
                    size_name = sizes.get((width, height), "A4")
                    return size_name
        return "A4"
    except:
        return "A4"

def wait_for_complete_file(file_path, timeout=10):
    initial_size = -1
    for _ in range(timeout):
        if not os.path.exists(file_path):
            return False
        current_size = os.path.getsize(file_path)
        if current_size == initial_size:
            return True
        initial_size = current_size
        time.sleep(1)
    return False

def process_single_xps(file_path):
    if not wait_for_complete_file(file_path):
        return

    base = os.path.basename(file_path)
    job_id = os.path.splitext(base)[0].replace("job_", "")

    if "%d" in job_id:
        job_id = "Неизвестно"

    page_count = 0
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            page_count = sum(1 for name in z.namelist() if name.lower().endswith(".fpage"))
    except:
        page_count = 1

    page_size = get_page_size(file_path)

    root = tk.Tk()
    root.title(f"Задание {job_id}")
    root.geometry("300x150")
    root.resizable(False, False)

    lbl = tk.Label(
        root,
        text=f"ID задания: {job_id}\nСтраниц: {page_count}\nРазмер: {page_size}",
        justify=tk.CENTER
    )
    lbl.pack(pady=10)

    def on_delete():
        try:
            os.remove(file_path)
        except:
            pass
        root.destroy()

    def on_send():
        try:
            hPrinter = win32print.OpenPrinter(REAL_PRINTER_NAME)
            with open(file_path, "rb") as f:
                xps_data = f.read()
            win32print.StartDocPrinter(hPrinter, 1, ("JobFromVirtual", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, xps_data)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отправить на принтер:\n{e}")

        try:
            os.remove(file_path)
        except:
            pass
        root.destroy()

    tk.Button(root, text="Удалить", width=12, command=on_delete).pack(side=tk.LEFT, padx=20, pady=10)
    tk.Button(root, text="Отправить", width=12, command=on_send).pack(side=tk.RIGHT, padx=20, pady=10)

    root.mainloop()

def watch_folder_loop(stop_event):
    if not os.path.exists(WATCH_FOLDER):
        os.makedirs(WATCH_FOLDER, exist_ok=True)

    known_files = {}

    while not stop_event.is_set():
        current_files = {name: os.path.getsize(os.path.join(WATCH_FOLDER, name))
                         for name in os.listdir(WATCH_FOLDER) if name.lower().endswith(".xps")}

        for name, size in current_files.items():
            if name not in known_files or known_files[name] != size:
                process_single_xps(os.path.join(WATCH_FOLDER, name))

        known_files = current_files
        time.sleep(1)

class ServiceFramework(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        watch_thread = threading.Thread(target=watch_folder_loop, args=(self.stop_event,))
        watch_thread.daemon = True
        watch_thread.start()
        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)
        watch_thread.join()

def run_as_console():
    print("Запуск printer_worker в консольном режиме.")
    stop_event = threading.Event()

    def on_ctrl_c(signum, frame):
        print("\nCtrl+C — выход...")
        stop_event.set()

    import signal
    signal.signal(signal.SIGINT, on_ctrl_c)

    watch_folder_loop(stop_event)
    print("Консольный режим завершён.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower() == "--console":
        run_as_console()
    else:
        win32serviceutil.HandleCommandLine(ServiceFramework)