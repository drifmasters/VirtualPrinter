# -*- coding: utf-8 -*-
"""
printer_worker.py

1) Запускается как Windows Service (под SYSTEM) и следит за папкой C:\VM_PRINTERS\VIRT1.
2) При появлении нового файла извлекает JobID, Pages, Size и вызывает интерактивное окно.
3) В окне – две кнопки: «Удалить» / «Отправить» – и в зависимости от выбора
   удаляет задание из очереди виртуального принтера или пересылает его на реальный.
"""

import os
import sys
import time
import threading
import zipfile
import tkinter as tk
from tkinter import messagebox
import win32print
import win32con
import win32event
import win32file
import win32service
import win32serviceutil
import win32event
import win32api

# ------------------------------------------------------------------------------
# КОНСТАНТЫ (дублируются из setup_virtual_printer.py, чтобы сервис знал, с чем работать)
# ------------------------------------------------------------------------------
VIRTUAL_PRINTER_NAME = "MyVirtualPrinterPython"
REAL_PRINTER_NAME = r"\\192.168.1.200\OfficePrinter"
WATCH_FOLDER = r"C:\VM_PRINTERS\VIRT1"

# ------------------------------------------------------------------------------
# СЛУЖЕБНЫЙ КЛАСС (для регистрации как сервиса)
# ------------------------------------------------------------------------------
class ServiceFramework(win32serviceutil.ServiceFramework):
    _svc_name_ = "PyVirtualPrinterWorker"
    _svc_display_name_ = "Python VirtualPrinter Worker Service"
    _svc_description_ = "Следит за папкой виртуального принтера и показывает окно с управлением заданиями."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        # Создаём событие, чтобы знать, когда нужно остановиться
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.stop_requested = False

    def SvcStop(self):
        # Валидируем запрос на остановку
        self.stop_requested = True
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        # Основной метод – запустим рабочий цикл в отдельном потоке
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        worker_thread = threading.Thread(target=self.monitor_folder, daemon=True)
        worker_thread.start()
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)

        # Ждём события остановки
        win32event.WaitForSingleObject(self.hWaitStop, win32con.INFINITE)
        self.stop_requested = True
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # Дадим потоку время завершиться
        time.sleep(1)

    # ------------------------------------------------------------------------------
    # ОСНОВНАЯ ЛОГИКА: МОНИТОРИНГ ПАПКИ
    # ------------------------------------------------------------------------------
    def monitor_folder(self):
        """
        Следим за появлением новых XPS/EMF/SPL файлов и обрабатываем их.
        """
        # Если папки нет, создаём
        if not os.path.exists(WATCH_FOLDER):
            try:
                os.makedirs(WATCH_FOLDER)
            except:
                pass

        # Объект для уведомлений об изменениях
        hDir = win32file.FindFirstChangeNotification(
            WATCH_FOLDER,
            False,
            win32con.FILE_NOTIFY_CHANGE_FILE_NAME | win32con.FILE_NOTIFY_CHANGE_LAST_WRITE
        )
        if hDir == win32con.INVALID_HANDLE_VALUE:
            return

        # Собираем уже существующие файлы, чтобы не обрабатывать их
        processed = set(os.listdir(WATCH_FOLDER))

        while not self.stop_requested:
            # Ждём изменений с таймаутом (чтобы периодически проверять флаг stop_requested)
            rc = win32event.WaitForSingleObject(hDir, 1000)
            if rc == win32con.WAIT_OBJECT_0:
                all_files = set(os.listdir(WATCH_FOLDER))
                new_files = all_files - processed
                for fname in new_files:
                    if not fname.lower().endswith((".xps", ".spl", ".emf")):
                        processed.add(fname)
                        continue

                    full_path = os.path.join(WATCH_FOLDER, fname)
                    # Дожидаемся, пока файл полностью запишется
                    time.sleep(0.5)

                    # Находим JobID, сопоставляя fname с pDocument
                    job_id = self.find_job_id_by_filename(fname)
                    if job_id is None:
                        processed.add(fname)
                        continue

                    # Получаем инфо о задании
                    info = self.get_job_info(job_id)
                    if info is None:
                        processed.add(fname)
                        continue

                    # Если pages == 0 и это XPS – считаем через zip
                    if info["total_pages"] == 0 and fname.lower().endswith(".xps"):
                        pages = self.count_xps_pages(full_path)
                        info["total_pages"] = pages

                    # Конвертируем paper size в мм (сотых дюйма → мм)
                    dm = info["devmode"]
                    paper_w = getattr(dm, "PaperWidth", None)
                    paper_h = getattr(dm, "PaperLength", None)
                    if paper_w and isinstance(paper_w, int):
                        pw_mm = paper_w * 0.254
                        ph_mm = paper_h * 0.254
                        info["paper_size_mm"] = (round(pw_mm,1), round(ph_mm,1))
                    else:
                        info["paper_size_mm"] = ("неизвестно","неизвестно")

                    # Покажем интерактивное окно (в сессии пользователя)
                    user_choice = self.show_job_window_and_wait(info, full_path, job_id)
                    if user_choice == "delete":
                        self.delete_job(job_id)
                        try:
                            os.remove(full_path)
                        except: pass
                    elif user_choice == "send":
                        ok = self.send_to_real_printer(full_path)
                        if ok:
                            self.delete_job(job_id)
                            try:
                                os.remove(full_path)
                            except: pass
                        else:
                            # Если не удалось, уведомляем пользователя
                            self.notify_user("Ошибка", f"Не удалось отправить задание {job_id} на реальный принтер.")
                    else:
                        # Пользователь закрыл окно – просто удаляем задание
                        self.delete_job(job_id)
                    processed.add(fname)

                win32file.FindNextChangeNotification(hDir)

        # Закрываем нотификацию
        win32file.FindCloseChangeNotification(hDir)

    # ------------------------------------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ WORKER
    # ------------------------------------------------------------------------------
    def find_job_id_by_filename(self, fname):
        """
        Перебираем задания в очереди виртуального принтера и ищем то, у которого pDocument содержит имя файла.
        Если не находим по имени – возвращаем None.
        """
        try:
            hPrin = win32print.OpenPrinter(VIRTUAL_PRINTER_NAME)
            jobs = win32print.EnumJobs(hPrin, 0, -1, 2)
        except:
            return None

        for j in jobs:
            docname = j["pDocument"]
            if docname and fname.lower() in docname.lower():
                job_id = j["JobId"]
                win32print.ClosePrinter(hPrin)
                return job_id

        win32print.ClosePrinter(hPrin)
        return None

    def get_job_info(self, job_id):
        """
        Возвращает dict с ключами:
        - job_id
        - total_pages
        - devmode (структура DEVMODE)
        """
        try:
            hPrin = win32print.OpenPrinter(VIRTUAL_PRINTER_NAME)
            job_info = win32print.GetJob(hPrin, job_id, 2)
            total = job_info["TotalPages"]
            devmode = job_info["pDevMode"]
            win32print.ClosePrinter(hPrin)
            return {"job_id": job_id, "total_pages": total, "devmode": devmode}
        except:
            return None

    def count_xps_pages(self, xps_path):
        """
        Считает число FixedPage.fpage в XPS (ZIP-архив).
        """
        try:
            z = zipfile.ZipFile(xps_path, 'r')
            fps = [f for f in z.namelist() if f.lower().endswith("fixedpage.fpage")]
            z.close()
            return len(fps)
        except:
            return 0

    def delete_job(self, job_id):
        """
        Удаляет задание из очереди виртуального принтера.
        """
        try:
            hPrin = win32print.OpenPrinter(VIRTUAL_PRINTER_NAME)
            win32print.SetJob(hPrin, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
            win32print.ClosePrinter(hPrin)
        except:
            pass

    def send_to_real_printer(self, xps_filepath):
        """
        Отправляет XPS/EMF/SPL на реальный принтер как RAW.
        """
        try:
            hPrin = win32print.OpenPrinter(REAL_PRINTER_NAME)
            with open(xps_filepath, "rb") as f:
                data = f.read()
            hJob = win32print.StartDocPrinter(hPrin, 1, ("VPDoc", None, "RAW"))
            win32print.StartPagePrinter(hPrin)
            win32print.WritePrinter(hPrin, data)
            win32print.EndPagePrinter(hPrin)
            win32print.EndDocPrinter(hPrin)
            win32print.ClosePrinter(hPrin)
            return True
        except:
            return False

    def notify_user(self, title, message):
        """
        Пытается показать MessageBox в интерактивной сессии.
        """
        # В сервисе под SYSTEM: MessageBox может не появиться в сессии пользователя.
        # Но в большинстве случаев (при Allow service to interact) он всплывёт.
        try:
            win32api.MessageBox(0, message, title, win32con.MB_ICONERROR | win32con.MB_OK)
        except:
            pass

    # ------------------------------------------------------------------------------
    # ОТОБРАЖЕНИЕ ОКНА В ПЕРСОНАЛЬНОЙ СЕССИИ (Tkinter)
    # ------------------------------------------------------------------------------
    def show_job_window_and_wait(self, info, xps_filepath, job_id):
        """
        Запускает окно с информацией о задании и двумя кнопками.
        Так как сервис запущен в Session 0, нужно явно указать, чтобы окно появилось
        в интерактивной сессии пользователя. Один из вариантов – через Win32 API:
        win32gui.SetForegroundWindow и т.д. В простейшем случае Tkinter сам отрисует
        окно там, где нужно, если службе разрешено взаимодействовать с рабочим столом.
        """

        # Создаём отдельный поток, который поднимет окно в интерактивном режиме.
        result_container = {"choice": None}

        def _show():
            root = tk.Tk()
            root.attributes("-topmost", True)
            root.title(f"Печать: JobID {info['job_id']}")
            root.geometry("350x180")
            root.resizable(False, False)

            lbl1 = tk.Label(root, text=f"ID задания: {info['job_id']}")
            lbl1.pack(pady=5)
            lbl2 = tk.Label(root, text=f"Кол-во страниц: {info['total_pages']}")
            lbl2.pack(pady=5)
            pw, ph = info["paper_size_mm"]
            lbl3 = tk.Label(root, text=f"Размер страницы (мм): {pw}×{ph}")
            lbl3.pack(pady=5)

            btn_frame = tk.Frame(root)
            btn_frame.pack(pady=15)

            def on_delete():
                result_container["choice"] = "delete"
                root.destroy()

            def on_send():
                result_container["choice"] = "send"
                root.destroy()

            btn_del = tk.Button(btn_frame, text="Удалить задание", width=15, command=on_delete)
            btn_del.pack(side=tk.LEFT, padx=5)
            btn_send = tk.Button(btn_frame, text="Отправить на реальный", width=15, command=on_send)
            btn_send.pack(side=tk.LEFT, padx=5)

            root.mainloop()

        # Запускаем окно и ждём, пока пользователь что-то выберет или закроет
        t = threading.Thread(target=_show)
        t.start()
        t.join()  # блокируем сервисный поток до тех пор, пока окно не закроется

        return result_container["choice"]

# ------------------------------------------------------------------------------
# ТОЧКА ВХОДА ДЛЯ СЕРВИСА
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    # Если скрипт запущен с аргументом установки/удаления – передаём управление win32serviceutil
    # Например: python printer_worker.py install / remove / start / stop
    if len(sys.argv) > 1 and sys.argv[1].lower() in ("install", "remove", "start", "stop", "restart"):
        win32serviceutil.HandleCommandLine(ServiceFramework)
    else:
        # Если просто запущен без аргументов – запускаем службу в режиме "in-session"
        # Это позволит интерактивно показывать окна, если скрипт запущен не как сервис,
        # а вручную (для отладки).
        # В продакшене мы регистрируем ServiceFramework и Windows будет вызывать SvcDoRun.
        worker = ServiceFramework(sys.argv)
        worker.SvcDoRun()
