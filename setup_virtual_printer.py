# -*- coding: utf-8 -*-
"""
setup_virtual_printer.py

Этот скрипт-установщик создаёт и удаляет «виртуальный принтер» под Windows 10 x64 с помощью Python:
1. Создаёт папку C:\VM_PRINTERS\VIRT1, если её нет.
2. Находит в системе драйвер XPS (например, «Microsoft XPS Document Writer»).
3. Создаёт локальный порт с именем "C:\VM_PRINTERS\VIRT1\job_%d.xps" через PowerShell Add-PrinterPort.
4. Вызывает PrintUIEntry (rundll32 printui.dll,PrintUIEntry) для установки принтера
   MyVirtualPrinterPython с этим локальным портом.
5. Настраивает ACL у реального сетевого принтера (удаляет право «Print» у Users,
   оставляя его только у SYSTEM и Administrators), чтобы пользователь не мог
   печатать напрямую в обход виртуального.
6. Регистрирует Windows-службу PyVirtualPrinterWorker (printer_worker.py),
   которая в фоне следит за папкой C:\VM_PRINTERS\VIRT1 и при появлении файла
   показывает окно с информацией (ID задания, количество страниц, размер страницы)
   и кнопками «Удалить» / «Отправить на реальный принтер».
7. Позволяет одним кликом «Установить и запустить» или «Остановить и удалить».

Важно: для выполнения операций установки принтера и регистрации службы требуются
права администратора. Если скрипт запущен без повышенных прав, мастер выдаст предупреждение.
"""

import os
import sys
import subprocess
import ctypes
import tkinter as tk
from tkinter import messagebox
import win32print
import win32con
import win32security
import win32serviceutil
import win32service

# ------------------------------------------------------------------------------
# КОНСТАНТЫ (имена принтеров, пути и т. д.)
# ------------------------------------------------------------------------------
VIRTUAL_PRINTER_NAME = "MyVirtualPrinterPython"
VIRT_BASE_DIR = r"C:\VM_PRINTERS\VIRT1"
VIRT_PORT_NAME = os.path.join(VIRT_BASE_DIR, "job_%d.xps")
REAL_PRINTER_NAME = r"\\192.168.1.200\OfficePrinter"
SERVICE_NAME = "PyVirtualPrinterWorker"
SERVICE_DISPLAY_NAME = "Python VirtualPrinter Worker Service"

# ------------------------------------------------------------------------------
# ПРОВЕРКА ПРАВ АДМИНИСТРАТОРА
# ------------------------------------------------------------------------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

# ------------------------------------------------------------------------------
# СОЗДАНИЕ И УДАЛЕНИЕ ВИРТУАЛЬНОГО ПРИНТЕРА
# ------------------------------------------------------------------------------
def create_virtual_printer():
    """
    1) Создаёт папку VIRT_BASE_DIR, если её нет.
    2) Находит в системе драйвер XPS.
    3) Создаёт локальный порт через PowerShell: Add-PrinterPort -Name "{VIRT_PORT_NAME}".
    4) Вызывает PrintUIEntry в виде одной строки (shell=True) для установки принтера:
       rundll32 printui.dll,PrintUIEntry /if /b "{VIRTUAL_PRINTER_NAME}"
       /r "{VIRT_PORT_NAME}" /m "{DriverName}"
    """
    # 1) Создаём папку, если отсутствует
    if not os.path.exists(VIRT_BASE_DIR):
        try:
            os.makedirs(VIRT_BASE_DIR)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать папку {VIRT_BASE_DIR}:\n{e}")
            return False

    # 2) Проверяем, нет ли уже такого принтера
    existing_printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)]
    if VIRTUAL_PRINTER_NAME in existing_printers:
        return True  # Уже установлен

    # 3) Находим драйвер XPS
    try:
        drivers_info = win32print.EnumPrinterDrivers(None, None, 1)
        driver_names = [d["Name"] for d in drivers_info]
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось получить список драйверов принтеров:\n{e}")
        return False

    xps_candidates = [name for name in driver_names if "xps document writer" in name.lower()]
    if not xps_candidates:
        xps_candidates = [name for name in driver_names if "xps" in name.lower()]

    if not xps_candidates:
        messagebox.showerror(
            "Ошибка",
            "В системе не найден драйвер XPS (например, 'Microsoft XPS Document Writer').\n"
            "Установите его через:\n"
            "Панель управления → Программы и компоненты → Включение компонентов Windows →\n"
            "«Службы документов и печати» → отметьте «Служба XPS».")
        return False

    virt_driver = xps_candidates[0]

    # 4) Создаём локальный порт через PowerShell
    ps_command = (
        f'Add-PrinterPort -Name "{VIRT_PORT_NAME}" -ErrorAction SilentlyContinue'
    )
    try:
        subprocess.check_call(
            ["powershell.exe", "-NoProfile", "-Command", ps_command],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception as e:
        # Если порт не смог создать PowerShell
        messagebox.showerror("Ошибка",
            f"Не удалось создать локальный порт \"{VIRT_PORT_NAME}\" через PowerShell:\n{e}\n"
            "Проверьте, что у вас есть права администратора.")
        return False

    # 5) Устанавливаем принтер через PrintUIEntry (единая строка, shell=True!)
    cmd_str = (
        f'rundll32 printui.dll,PrintUIEntry '
        f'/if '
        f'/b "{VIRTUAL_PRINTER_NAME}" '
        f'/r "{VIRT_PORT_NAME}" '
        f'/m "{virt_driver}"'
    )
    try:
        subprocess.check_call(cmd_str, shell=True)
        return True
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка",
            f"Не удалось создать виртуальный принтер:\n{e}\n"
            "Возможная причина: порт не был создан.")
        return False

def delete_virtual_printer():
    """
    Удаляет виртуальный принтер (если он существует) через PrintUIEntry:
    rundll32 printui.dll,PrintUIEntry /dl /n "{VIRTUAL_PRINTER_NAME}"
    """
    existing_printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)]
    if VIRTUAL_PRINTER_NAME not in existing_printers:
        return True

    cmd_str = (
        f'rundll32 printui.dll,PrintUIEntry '
        f'/dl '
        f'/n "{VIRTUAL_PRINTER_NAME}"'
    )
    try:
        subprocess.check_call(cmd_str, shell=True)
    except subprocess.CalledProcessError as e:
        # Предупреждаем, но не прерываем выполнение
        messagebox.showwarning("Предупреждение", f"Не удалось удалить виртуальный принтер:\n{e}")
    return True

# ------------------------------------------------------------------------------
# НАСТРОЙКА ПРАВ У РЕАЛЬНОГО ПРИНТЕРА
# ------------------------------------------------------------------------------
def fix_real_printer_acl():
    try:
        hPrinter = win32print.OpenPrinter(REAL_PRINTER_NAME)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть реальный принтер {REAL_PRINTER_NAME}:\n{e}")
        return False

    try:
        sd = win32print.GetPrinter(hPrinter, 2)["pSecurityDescriptor"]
        priv_dacl = win32security.GetSecurityInfo(
            hPrinter,
            win32security.SE_PRINTER_OBJECT,
            win32security.DACL_SECURITY_INFORMATION
        )[1]

        new_dacl = win32security.ACL()
        system_sid = win32security.CreateWellKnownSid(win32security.WinLocalSystemSid, None)
        new_dacl.AddAccessAllowedAce(win32security.ACL_REVISION, win32con.PRINTER_ALL_ACCESS, system_sid)

        admins_sid, _, _ = win32security.LookupAccountName("", "Administrators")
        new_dacl.AddAccessAllowedAce(win32security.ACL_REVISION, win32con.PRINTER_ALL_ACCESS, admins_sid)

        win32security.SetSecurityInfo(
            hPrinter,
            win32security.SE_PRINTER_OBJECT,
            win32security.DACL_SECURITY_INFORMATION,
            None, None, new_dacl, None
        )
        win32print.ClosePrinter(hPrinter)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось изменить права у принтера {REAL_PRINTER_NAME}:\n{e}")
        win32print.ClosePrinter(hPrinter)
        return False

# ------------------------------------------------------------------------------
# РЕГИСТРАЦИЯ И УПРАВЛЕНИЕ СЛУЖБОЙ (printer_worker.py)
# ------------------------------------------------------------------------------
def register_service():
    python_exe = sys.executable.replace("\\", "\\\\")
    worker_py = os.path.join(os.path.dirname(__file__), "printer_worker.py").replace("\\", "\\\\")
    try:
        try:
            if win32serviceutil.QueryServiceStatus(SERVICE_NAME):
                win32serviceutil.RemoveService(SERVICE_NAME)
        except Exception:
            pass

        win32serviceutil.InstallService(
            python_exe,
            SERVICE_NAME,
            SERVICE_DISPLAY_NAME,
            startType=win32service.SERVICE_AUTO_START,
            exeArgs=f'"{worker_py}"'
        )
        hSCM = win32service.OpenSCManager(None, None, win32con.SC_MANAGER_ALL_ACCESS)
        hSvc = win32service.OpenService(hSCM, SERVICE_NAME, win32con.SERVICE_ALL_ACCESS)
        win32service.ChangeServiceConfig(
            hSvc,
            win32service.SERVICE_NO_CHANGE,
            win32service.SERVICE_AUTO_START,
            win32service.SERVICE_ERROR_NORMAL,
            python_exe,
            None, 0, None, None, None, None
        )
        win32service.CloseServiceHandle(hSvc)
        win32service.CloseServiceHandle(hSCM)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось зарегистрировать службу:\n{e}")
        return False

def start_service():
    try:
        win32serviceutil.StartService(SERVICE_NAME)
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить службу {SERVICE_NAME}:\n{e}")
        return False

def stop_service():
    try:
        win32serviceutil.StopService(SERVICE_NAME)
        return True
    except Exception:
        return False

def remove_service():
    try:
        win32serviceutil.RemoveService(SERVICE_NAME)
        return True
    except Exception:
        return False

# ------------------------------------------------------------------------------
# GUI (Tkinter) – мастер установки/удаления виртуального принтера
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
                "Этот мастер автоматически:\n"
                "1. Создаст виртуальный принтер (через PowerShell + PrintUI) с портом\n"
                "   C:\\VM_PRINTERS\\VIRT1\\job_%d.xps.\n"
                "2. Настроит права у реального принтера (RESTRICTED).\n"
                "3. Зарегистрирует и запустит службу для обработки заданий.\n\n"
                "После этого пользователю останется лишь печатать\n"
                "в виртуальный принтер — всё остальное произойдёт автоматически.\n\n"
                "Для корректной работы требуются права администратора."
            ),
            justify=tk.LEFT,
        )
        lbl_info.pack(pady=(0, 15))

        self.btn_install = tk.Button(frm, text="Установить и запустить", width=30, command=self.on_install)
        self.btn_install.pack(pady=5)

        self.btn_stop = tk.Button(frm, text="Остановить и удалить", width=30, command=self.on_uninstall)
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

        self.log("1. Создаём виртуальный принтер (через PowerShell + PrintUIEntry)...")
        if create_virtual_printer():
            self.log(f"✓ Виртуальный принтер '{VIRTUAL_PRINTER_NAME}' готов.")
        else:
            self.log("✗ Не удалось создать виртуальный принтер.")
            self.btn_install.config(state=tk.NORMAL)
            return

        self.log("2. Настраиваем права у реального принтера...")
        if fix_real_printer_acl():
            self.log(f"✓ Права у '{REAL_PRINTER_NAME}' обновлены.")
        else:
            self.log("✗ Не удалось настроить права у реального принтера.")
            self.btn_install.config(state=tk.NORMAL)
            return

        self.log("3. Регистрируем службу обработчика...")
        if register_service():
            self.log(f"✓ Служба '{SERVICE_NAME}' зарегистрирована.")
        else:
            self.log("✗ Не удалось зарегистрировать службу.")
            self.btn_install.config(state=tk.NORMAL)
            return

        self.log("4. Запускаем службу...")
        if start_service():
            self.log("✓ Служба запущена и работает.")
        else:
            self.log("✗ Не удалось запустить службу.")
            self.btn_install.config(state=tk.NORMAL)
            return

        messagebox.showinfo("Готово", "Установка завершена успешно.\n"
                             "Виртуальный принтер создан, служба запущена.")
        self.btn_install.config(state=tk.NORMAL)

    def on_uninstall(self):
        if not messagebox.askyesno("Подтвердите", "Удалить виртуальный принтер и службу?"):
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

        self.log("1. Останавливаем службу...")
        stop_service()
        self.log("   Служба остановлена (если была запущена).")

        self.log("2. Удаляем службу...")
        if remove_service():
            self.log("✓ Служба удалена.")
        else:
            self.log("✗ Служба не удалена (возможно, уже отсутствует).")

        self.log("3. Удаляем виртуальный принтер...")
        if delete_virtual_printer():
            self.log("✓ Виртуальный принтер удалён.")
        else:
            self.log("✗ Не удалось удалить виртуальный принтер.")

        messagebox.showinfo("Готово", "Виртуальный принтер и служба удалены.")
        self.btn_install.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL)

def main():
    root = tk.Tk()
    app = SetupGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
