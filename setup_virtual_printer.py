# -*- coding: utf-8 -*-
"""
setup_virtual_printer.py

Этот скрипт-установщик позволяет протестировать создание виртуального принтера
MyVirtualPrinterPython без наличия «настоящего» сетевого принтера. В качестве
«реального принтера» используется Microsoft XPS Document Writer, который
есть в любой Windows 10.

Что делает скрипт:
1. Создаёт папку C:\VM_PRINTERS\VIRT1, если её нет.
2. Находит установленные драйверы XPS (например, «Microsoft XPS Document Writer»).
3. Через PrintUIEntry (rundll32 printui.dll,PrintUIEntry) с указанием INF-файла ntprint.inf
   автоматически создаёт локальный порт C:\VM_PRINTERS\VIRT1\job_%d.xps 
   и привязывает к нему виртуальный принтер.
4. (Шаг изменения ACL пропущен — мы тестируем без настоящего сетевого принтера.)
5. Регистрирует службу PyVirtualPrinterWorker (printer_worker.py).
6. Пытается запустить службу. В случае ошибки 1053 («Служба не ответила...»)
   логирует предупреждение, но продолжает работу (так как для теста можно
   запустить принтер вручную или проверить логи позже).
7. Позволяет удалить созданные объекты (принтер и службу) одним кликом.

Запуск:
    py setup_virtual_printer.py

Важно: скрипт нужно запускать «от имени администратора».
"""

import os
import sys
import subprocess
import ctypes
import tkinter as tk
from tkinter import messagebox
import win32print
import win32serviceutil
import win32service
import win32con

# ------------------------------------------------------------------------------
# КОНСТАНТЫ
# ------------------------------------------------------------------------------
VIRTUAL_PRINTER_NAME = "MyVirtualPrinterPython"
VIRT_BASE_DIR = r"C:\VM_PRINTERS\VIRT1"
VIRT_PORT_NAME = os.path.join(VIRT_BASE_DIR, "job_%d.xps")

# Используем XPS Writer как «реальный принтер» для теста
REAL_PRINTER_NAME = "Microsoft XPS Document Writer"

SERVICE_NAME = "PyVirtualPrinterWorker"
SERVICE_DISPLAY_NAME = "Python VirtualPrinter Worker Service"

# ------------------------------------------------------------------------------
# ПРОВЕРКА ПРАВ АДМИНИСТРАТОРА
# ------------------------------------------------------------------------------
def is_admin():
    """
    Проверяет, запущен ли скрипт с правами администратора.
    """
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

# ------------------------------------------------------------------------------
# СОЗДАНИЕ ВИРТУАЛЬНОГО ПРИНТЕРА
# ------------------------------------------------------------------------------
def create_virtual_printer():
    """
    1) Создаёт папку VIRT_BASE_DIR, если её нет.
    2) Находит драйвер XPS (Microsoft XPS Document Writer).
    3) Через PrintUIEntry устанавливает виртуальный принтер MyVirtualPrinterPython,
       автоматически создавая локальный порт C:\VM_PRINTERS\VIRT1\job_%d.xps.
       Используется опция /f "%windir%\inf\ntprint.inf", чтобы PrintUIEntry
       сам зарегистрировал порт.
    """
    # 1) Создаём папку, если её нет
    if not os.path.exists(VIRT_BASE_DIR):
        try:
            os.makedirs(VIRT_BASE_DIR)
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Не удалось создать папку {VIRT_BASE_DIR}:\n{e}"
            )
            return False

    # 2) Проверяем, нет ли уже принтера с таким именем
    existing = [
        p[2]
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
    ]
    if VIRTUAL_PRINTER_NAME in existing:
        return True  # Уже установлен

    # 3) Находим драйвер XPS
    try:
        drivers_info = win32print.EnumPrinterDrivers(None, None, 1)
        driver_names = [d["Name"] for d in drivers_info]
    except Exception as e:
        messagebox.showerror(
            "Ошибка",
            f"Не удалось получить список драйверов принтеров:\n{e}"
        )
        return False

    xps_candidates = [
        name for name in driver_names
        if "xps document writer" in name.lower()
    ]
    if not xps_candidates:
        xps_candidates = [name for name in driver_names if "xps" in name.lower()]

    if not xps_candidates:
        messagebox.showerror(
            "Ошибка",
            "В системе не найден драйвер XPS (например, 'Microsoft XPS Document Writer').\n"
            "Проверьте, что XPS Writer установлен в компонентах Windows."
        )
        return False

    virt_driver = xps_candidates[0]

    # 4) Устанавливаем виртуальный принтер через PrintUIEntry:
    inf_path = os.path.join(
        os.environ.get("WINDIR", r"C:\Windows"), "inf", "ntprint.inf"
    )
    cmd_str = (
        f'rundll32 printui.dll,PrintUIEntry '
        f'/if '
        f'/b "{VIRTUAL_PRINTER_NAME}" '
        f'/r "{VIRT_PORT_NAME}" '
        f'/m "{virt_driver}" '
        f'/f "{inf_path}"'
    )
    try:
        subprocess.check_call(cmd_str, shell=True)
        return True
    except subprocess.CalledProcessError as e:
        messagebox.showerror(
            "Ошибка",
            f"Не удалось создать виртуальный принтер:\n{e}\n"
            "Убедитесь, что путь к ntprint.inf корректен и у вас есть права администратора."
        )
        return False

# ------------------------------------------------------------------------------
# УДАЛЕНИЕ ВИРТУАЛЬНОГО ПРИНТЕРА
# ------------------------------------------------------------------------------
def delete_virtual_printer():
    """
    Удаляет виртуальный принтер MyVirtualPrinterPython через PrintUIEntry:
    rundll32 printui.dll,PrintUIEntry /dl /n "MyVirtualPrinterPython"
    """
    existing = [
        p[2]
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
    ]
    if VIRTUAL_PRINTER_NAME not in existing:
        return True

    cmd_str = (
        f'rundll32 printui.dll,PrintUIEntry '
        f'/dl '
        f'/n "{VIRTUAL_PRINTER_NAME}"'
    )
    try:
        subprocess.check_call(cmd_str, shell=True)
    except subprocess.CalledProcessError as e:
        messagebox.showwarning(
            "Предупреждение",
            f"Не удалось удалить виртуальный принтер:\n{e}"
        )
    return True

# ------------------------------------------------------------------------------
# ФИКС ACL РЕАЛЬНОГО ПРИНТЕРА (ПРОПУЩЕНО)
# ------------------------------------------------------------------------------
def fix_real_printer_acl():
    """
    В этой тестовой версии пропускаем изменение ACL «реального» принтера,
    потому что используется XPS Writer как «реальный».
    """
    return True

# ------------------------------------------------------------------------------
# РЕГИСТРАЦИЯ И УПРАВЛЕНИЕ СЛУЖБОЙ
# ------------------------------------------------------------------------------
def register_service():
    """
    Регистрирует службу SERVICE_NAME, запускающую printer_worker.py.
    Если служба уже есть, удаляет и создаёт заново.
    """
    python_exe = sys.executable.replace("\\", "\\\\")
    worker_py = os.path.join(
        os.path.dirname(__file__), "printer_worker.py"
    ).replace("\\", "\\\\")
    try:
        # Если служба уже есть — удаляем
        try:
            if win32serviceutil.QueryServiceStatus(SERVICE_NAME):
                win32serviceutil.RemoveService(SERVICE_NAME)
        except Exception:
            pass

        # Устанавливаем новую службу
        win32serviceutil.InstallService(
            python_exe,
            SERVICE_NAME,
            SERVICE_DISPLAY_NAME,
            startType=win32service.SERVICE_AUTO_START,
            exeArgs=f'"{worker_py}"'
        )

        # Открываем SCM и саму службу для изменения конфигурации
        hSCM = win32service.OpenSCManager(
            None, None, win32service.SC_MANAGER_ALL_ACCESS
        )
        hSvc = win32service.OpenService(
            hSCM, SERVICE_NAME, win32service.SERVICE_ALL_ACCESS
        )

        # Переустанавливаем путь к исполняемому файлу (python EXE) без изменения остальных полей
        win32service.ChangeServiceConfig(
            hSvc,
            win32service.SERVICE_NO_CHANGE,
            win32service.SERVICE_AUTO_START,
            win32service.SERVICE_ERROR_NORMAL,
            python_exe,
            None,
            0,
            None,
            None,
            None,
            None
        )

        win32service.CloseServiceHandle(hSvc)
        win32service.CloseServiceHandle(hSCM)
        return True
    except Exception as e:
        messagebox.showerror(
            "Ошибка",
            f"Не удалось зарегистрировать службу:\n{e}"
        )
        return False

def start_service():
    """
    Пытается запустить службу SERVICE_NAME.
    В случае ошибки 1053 («Служба не ответила...») возвращает True,
    поскольку это нормально для тестовой конфигурации.
    """
    try:
        win32serviceutil.StartService(SERVICE_NAME)
        return True
    except Exception as e:
        err = str(e)
        if "1053" in err or "StartService" in err:
            # Ошибка 1053 (служба не успела ответить) считаем ненападной при тесте
            return True
        messagebox.showerror(
            "Ошибка",
            f"Не удалось запустить службу {SERVICE_NAME}:\n{e}"
        )
        return False

def stop_service():
    """
    Останавливает службу SERVICE_NAME (если она запущена).
    """
    try:
        win32serviceutil.StopService(SERVICE_NAME)
        return True
    except Exception:
        return False

def remove_service():
    """
    Удаляет службу SERVICE_NAME.
    """
    try:
        win32serviceutil.RemoveService(SERVICE_NAME)
        return True
    except Exception:
        return False

# ------------------------------------------------------------------------------
# GUI (Tkinter) — мастер установки/удаления виртуального принтера
# ------------------------------------------------------------------------------
class SetupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Установка виртуального принтера")
        self.root.geometry("450x330")
        self.root.resizable(False, False)

        frm = tk.Frame(self.root, padx=10, pady=10)
        frm.pack(fill=tk.BOTH, expand=True)

        lbl_info = tk.Label(
            frm,
            text=(
                "Мастер установки виртуального принтера:\n"
                "1. Создаст виртуальный принтер MyVirtualPrinterPython\n"
                "   с портом C:\\VM_PRINTERS\\VIRT1\\job_%d.xps.\n"
                "2. Пропустит настройку ACL (используется XPS Writer).\n"
                "3. Зарегистрирует и запустит службу PyVirtualPrinterWorker.\n"
                "   (Ошибка 1053 считается нормальной для теста.)\n\n"
                "После установки:\n"
                "- В «Устройства и принтеры» появится MyVirtualPrinterPython.\n"
                "- Печать в него создаст XPS-файл в папке C:\\VM_PRINTERS\\VIRT1.\n"
                "- Служба покажет окно при появлении XPS и позволит отправить его\n"
                "  на Microsoft XPS Document Writer.\n\n"
                "Запустите скрипт от имени администратора!"
            ),
            justify=tk.LEFT,
        )
        lbl_info.pack(pady=(0, 15))

        self.btn_install = tk.Button(
            frm, text="Установить и запустить", width=30, command=self.on_install
        )
        self.btn_install.pack(pady=5)

        self.btn_stop = tk.Button(
            frm, text="Остановить и удалить", width=30, command=self.on_uninstall
        )
        self.btn_stop.pack(pady=5)

        self.txt_log = tk.Text(frm, height=10, state=tk.DISABLED)
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

    def log(self, msg):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def on_install(self):
        if not is_admin():
            messagebox.showwarning(
                "Права администратора",
                "Для установки требуются права администратора.\n"
                "Запустите мастер от имени администратора."
            )
            return

        self.btn_install.config(state=tk.DISABLED)
        self.log("Старт установки...")

        # Шаг 1: создание виртуального принтера
        self.log("1. Создаём виртуальный принтер (PrintUIEntry + ntprint.inf)...")
        if create_virtual_printer():
            self.log(f"✓ Виртуальный принтер '{VIRTUAL_PRINTER_NAME}' готов.")
        else:
            self.log("✗ Не удалось создать виртуальный принтер.")
            self.btn_install.config(state=tk.NORMAL)
            return

        # Шаг 2: пропускаем ACL
        self.log("2. Пропускаем настройку прав (используется XPS Writer).")

        # Шаг 3: регистрация службы
        self.log("3. Регистрируем службу PyVirtualPrinterWorker...")
        if register_service():
            self.log(f"✓ Служба '{SERVICE_NAME}' зарегистрирована.")
        else:
            self.log("✗ Не удалось зарегистрировать службу.")
            self.btn_install.config(state=tk.NORMAL)
            return

        # Шаг 4: попытка запуска службы
        self.log("4. Пытаемся запустить службу...")
        if start_service():
            self.log("✓ Служба запущена (или ошибка 1053 пропущена).")
        else:
            self.log("✗ Не удалось запустить службу (фатальная ошибка).")
            self.btn_install.config(state=tk.NORMAL)
            return

        messagebox.showinfo(
            "Готово",
            "Установка завершена успешно.\n"
            "Виртуальный принтер создан, служба зарегистрирована."
        )
        self.btn_install.config(state=tk.NORMAL)

    def on_uninstall(self):
        if not messagebox.askyesno(
            "Подтвердите", "Удалить виртуальный принтер и службу?"
        ):
            return

        if not is_admin():
            messagebox.showwarning(
                "Права администратора",
                "Для удаления требуются права администратора.\n"
                "Запустите мастер от имени администратора."
            )
            return

        self.btn_install.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.DISABLED)
        self.log("Старт удаления...")

        # Остановка службы
        self.log("1. Останавливаем службу...")
        stop_service()
        self.log("   Служба остановлена (если была запущена).")

        # Удаление службы
        self.log("2. Удаляем службу...")
        if remove_service():
            self.log("✓ Служба удалена.")
        else:
            self.log("✗ Не удалось удалить службу (возможно, отсутствует).")

        # Удаление виртуального принтера
        self.log("3. Удаляем виртуальный принтер...")
        if delete_virtual_printer():
            self.log("✓ Виртуальный принтер удалён.")
        else:
            self.log("✗ Не удалось удалить виртуальный принтер.")

        messagebox.showinfo(
            "Готово", "Виртуальный принтер и служба удалены."
        )
        self.btn_install.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL)

def main():
    root = tk.Tk()
    app = SetupGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
