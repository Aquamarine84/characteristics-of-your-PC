# -*- coding: utf-8 -*-
"""
PC Info Reporter
Простой скрипт для Windows, который показывает характеристики ПК
и сохраняет их в TXT-файл.

Запуск:
    python pc_info_gui.py

Требования:
    - Windows
    - Python 3.10+
    - PowerShell (обычно уже есть в Windows)
"""

import json
import os
import platform
import shutil
import socket
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext


APP_TITLE = "Характеристики ПК"
DEFAULT_FILENAME_PREFIX = "Характеристики_ПК"


def run_powershell_json(command: str):
    """
    Выполняет PowerShell-команду и пытается вернуть JSON.
    Возвращает Python-объект или None.
    """
    try:
        full_cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            command,
        ]
        result = subprocess.run(
            full_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=25,
        )
        if result.returncode != 0:
            return None

        output = result.stdout.strip()
        if not output:
            return None

        return json.loads(output)
    except Exception:
        return None


def run_powershell_text(command: str) -> str:
    """Выполняет PowerShell-команду и возвращает текст."""
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                command,
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
        return (result.stdout or "").strip()
    except Exception:
        return ""


def to_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def safe_get(dct, key, default="Не удалось определить"):
    if not isinstance(dct, dict):
        return default
    value = dct.get(key)
    if value in (None, "", []):
        return default
    return value


def bytes_to_gb(num) -> str:
    try:
        return f"{int(num) / (1024 ** 3):.2f} ГБ"
    except Exception:
        return "Не удалось определить"


def bytes_to_mb(num) -> str:
    try:
        return f"{int(num) / (1024 ** 2):.0f} МБ"
    except Exception:
        return "Не удалось определить"


def format_freq(mhz) -> str:
    try:
        mhz = float(mhz)
        ghz = mhz / 1000
        return f"{ghz:.2f} ГГц"
    except Exception:
        return "Не удалось определить"


def get_computer_system():
    data = run_powershell_json(
        r"Get-CimInstance Win32_ComputerSystem | "
        r"Select-Object Manufacturer, Model, TotalPhysicalMemory, UserName | "
        r"ConvertTo-Json -Depth 3"
    ) or {}

    return {
        "Производитель": safe_get(data, "Manufacturer"),
        "Модель": safe_get(data, "Model"),
        "Всего ОЗУ": bytes_to_gb(data.get("TotalPhysicalMemory")) if data.get("TotalPhysicalMemory") else "Не удалось определить",
        "Текущий пользователь": safe_get(data, "UserName"),
    }


def get_os_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_OperatingSystem | "
        r"Select-Object Caption, Version, BuildNumber, OSArchitecture, LastBootUpTime | "
        r"ConvertTo-Json -Depth 3"
    ) or {}

    return {
        "ОС": safe_get(data, "Caption", platform.platform()),
        "Версия": safe_get(data, "Version", platform.version()),
        "Сборка": safe_get(data, "BuildNumber"),
        "Разрядность": safe_get(data, "OSArchitecture", platform.architecture()[0]),
        "Последняя загрузка": safe_get(data, "LastBootUpTime"),
        "Имя компьютера": socket.gethostname(),
    }


def get_cpu_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_Processor | "
        r"Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, Manufacturer | "
        r"ConvertTo-Json -Depth 3"
    )
    cpu = to_list(data)
    if not cpu:
        return [{
            "Название": platform.processor() or "Не удалось определить",
            "Производитель": "Не удалось определить",
            "Ядер": "Не удалось определить",
            "Потоков": "Не удалось определить",
            "Макс. частота": "Не удалось определить",
        }]

    result = []
    for item in cpu:
        result.append({
            "Название": safe_get(item, "Name"),
            "Производитель": safe_get(item, "Manufacturer"),
            "Ядер": safe_get(item, "NumberOfCores"),
            "Потоков": safe_get(item, "NumberOfLogicalProcessors"),
            "Макс. частота": format_freq(item.get("MaxClockSpeed")),
        })
    return result


def get_ram_modules():
    data = run_powershell_json(
        r"Get-CimInstance Win32_PhysicalMemory | "
        r"Select-Object Manufacturer, PartNumber, Capacity, Speed, DeviceLocator, BankLabel | "
        r"ConvertTo-Json -Depth 3"
    )
    modules = to_list(data)
    result = []
    for item in modules:
        result.append({
            "Слот": safe_get(item, "DeviceLocator"),
            "Банк": safe_get(item, "BankLabel"),
            "Производитель": safe_get(item, "Manufacturer"),
            "Модель": safe_get(item, "PartNumber"),
            "Объём": bytes_to_gb(item.get("Capacity")),
            "Частота": f"{item.get('Speed')} МГц" if item.get("Speed") else "Не удалось определить",
        })
    return result


def get_gpu_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_VideoController | "
        r"Select-Object Name, AdapterRAM, DriverVersion, VideoProcessor, CurrentHorizontalResolution, CurrentVerticalResolution | "
        r"ConvertTo-Json -Depth 4"
    )
    gpus = to_list(data)
    result = []
    for item in gpus:
        w = item.get("CurrentHorizontalResolution")
        h = item.get("CurrentVerticalResolution")
        resolution = f"{w}x{h}" if w and h else "Не удалось определить"

        result.append({
            "Название": safe_get(item, "Name"),
            "Видеопроцессор": safe_get(item, "VideoProcessor"),
            "Память": bytes_to_gb(item.get("AdapterRAM")),
            "Драйвер": safe_get(item, "DriverVersion"),
            "Текущее разрешение": resolution,
        })
    return result


def get_disk_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_DiskDrive | "
        r"Select-Object Model, InterfaceType, MediaType, Size, SerialNumber | "
        r"ConvertTo-Json -Depth 4"
    )
    disks = to_list(data)
    result = []
    for item in disks:
        result.append({
            "Модель": safe_get(item, "Model"),
            "Интерфейс": safe_get(item, "InterfaceType"),
            "Тип": safe_get(item, "MediaType"),
            "Объём": bytes_to_gb(item.get("Size")),
            "Серийный номер": safe_get(item, "SerialNumber"),
        })
    return result


def get_logical_disks():
    data = run_powershell_json(
        r"Get-CimInstance Win32_LogicalDisk -Filter ""DriveType=3"" | "
        r"Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace | "
        r"ConvertTo-Json -Depth 3"
    )
    disks = to_list(data)
    result = []
    for item in disks:
        result.append({
            "Диск": safe_get(item, "DeviceID"),
            "Метка": safe_get(item, "VolumeName"),
            "Файловая система": safe_get(item, "FileSystem"),
            "Объём": bytes_to_gb(item.get("Size")),
            "Свободно": bytes_to_gb(item.get("FreeSpace")),
        })
    return result


def get_motherboard_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_BaseBoard | "
        r"Select-Object Manufacturer, Product, SerialNumber | "
        r"ConvertTo-Json -Depth 3"
    ) or {}
    return {
        "Производитель": safe_get(data, "Manufacturer"),
        "Модель": safe_get(data, "Product"),
        "Серийный номер": safe_get(data, "SerialNumber"),
    }


def get_bios_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_BIOS | "
        r"Select-Object Manufacturer, SMBIOSBIOSVersion, ReleaseDate, SerialNumber | "
        r"ConvertTo-Json -Depth 3"
    ) or {}
    return {
        "Производитель": safe_get(data, "Manufacturer"),
        "Версия BIOS": safe_get(data, "SMBIOSBIOSVersion"),
        "Дата выпуска": safe_get(data, "ReleaseDate"),
        "Серийный номер": safe_get(data, "SerialNumber"),
    }


def get_network_info():
    data = run_powershell_json(
        r"Get-CimInstance Win32_NetworkAdapterConfiguration | "
        r"Where-Object { $_.IPEnabled -eq $true } | "
        r"Select-Object Description, MACAddress, IPAddress, DefaultIPGateway, DHCPEnabled | "
        r"ConvertTo-Json -Depth 5"
    )
    adapters = to_list(data)
    result = []
    for item in adapters:
        ips = item.get("IPAddress")
        if isinstance(ips, list):
            ip_value = ", ".join(str(x) for x in ips)
        else:
            ip_value = str(ips) if ips else "Не удалось определить"

        gw = item.get("DefaultIPGateway")
        if isinstance(gw, list):
            gw_value = ", ".join(str(x) for x in gw)
        else:
            gw_value = str(gw) if gw else "Не удалось определить"

        result.append({
            "Адаптер": safe_get(item, "Description"),
            "MAC": safe_get(item, "MACAddress"),
            "IP": ip_value,
            "Шлюз": gw_value,
            "DHCP": "Да" if item.get("DHCPEnabled") else "Нет",
        })
    return result


def section(title: str) -> str:
    return f"\n{'=' * 70}\n{title}\n{'=' * 70}\n"


def format_dict(data: dict) -> str:
    lines = []
    for key, value in data.items():
        lines.append(f"{key}: {value}")
    return "\n".join(lines)


def format_list_of_dicts(items: list) -> str:
    if not items:
        return "Нет данных"
    blocks = []
    for index, item in enumerate(items, start=1):
        lines = [f"[{index}]"]
        for key, value in item.items():
            lines.append(f"{key}: {value}")
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


def collect_pc_info() -> str:
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    parts = [
        section("ОБЩАЯ ИНФОРМАЦИЯ"),
        f"Дата и время отчёта: {now}",
        format_dict(get_computer_system()),
        section("ОПЕРАЦИОННАЯ СИСТЕМА"),
        format_dict(get_os_info()),
        section("ПРОЦЕССОР"),
        format_list_of_dicts(get_cpu_info()),
        section("ОПЕРАТИВНАЯ ПАМЯТЬ"),
        format_list_of_dicts(get_ram_modules()),
        section("ВИДЕОКАРТА"),
        format_list_of_dicts(get_gpu_info()),
        section("ФИЗИЧЕСКИЕ ДИСКИ"),
        format_list_of_dicts(get_disk_info()),
        section("ЛОГИЧЕСКИЕ ДИСКИ"),
        format_list_of_dicts(get_logical_disks()),
        section("МАТЕРИНСКАЯ ПЛАТА"),
        format_dict(get_motherboard_info()),
        section("BIOS"),
        format_dict(get_bios_info()),
        section("СЕТЕВЫЕ АДАПТЕРЫ"),
        format_list_of_dicts(get_network_info()),
    ]
    return "\n".join(parts).strip() + "\n"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x650")
        self.root.minsize(760, 520)

        self.last_report = ""

        top_frame = tk.Frame(root, padx=10, pady=10)
        top_frame.pack(fill="x")

        title = tk.Label(
            top_frame,
            text="Просмотр характеристик ПК",
            font=("Segoe UI", 16, "bold")
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            top_frame,
            text="Нажмите кнопку, чтобы собрать информацию о компьютере и сохранить её в файл.",
            font=("Segoe UI", 10)
        )
        subtitle.pack(anchor="w", pady=(4, 0))

        btn_frame = tk.Frame(root, padx=10, pady=5)
        btn_frame.pack(fill="x")

        self.collect_btn = tk.Button(
            btn_frame,
            text="Собрать характеристики",
            font=("Segoe UI", 10, "bold"),
            width=24,
            command=self.on_collect
        )
        self.collect_btn.pack(side="left", padx=(0, 8))

        self.save_btn = tk.Button(
            btn_frame,
            text="Сохранить в TXT",
            font=("Segoe UI", 10),
            width=18,
            state="disabled",
            command=self.on_save
        )
        self.save_btn.pack(side="left", padx=(0, 8))

        self.copy_btn = tk.Button(
            btn_frame,
            text="Копировать текст",
            font=("Segoe UI", 10),
            width=18,
            state="disabled",
            command=self.on_copy
        )
        self.copy_btn.pack(side="left")

        self.status_var = tk.StringVar(value="Готово к работе.")
        status = tk.Label(
            root,
            textvariable=self.status_var,
            anchor="w",
            padx=10,
            pady=6,
            relief="groove"
        )
        status.pack(fill="x", side="bottom")

        self.text = scrolledtext.ScrolledText(
            root,
            wrap="word",
            font=("Consolas", 10)
        )
        self.text.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self.text.insert("1.0", "Здесь появятся характеристики компьютера.")
        self.text.configure(state="disabled")

    def set_text(self, value: str):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        self.text.configure(state="disabled")

    def on_collect(self):
        self.status_var.set("Сбор информации...")
        self.root.update_idletasks()

        try:
            report = collect_pc_info()
            self.last_report = report
            self.set_text(report)
            self.save_btn.config(state="normal")
            self.copy_btn.config(state="normal")
            self.status_var.set("Готово. Информация собрана.")
        except Exception as e:
            self.status_var.set("Ошибка при сборе информации.")
            messagebox.showerror("Ошибка", f"Не удалось собрать характеристики.\n\n{e}")

    def on_save(self):
        if not self.last_report.strip():
            messagebox.showwarning("Нет данных", "Сначала соберите характеристики.")
            return

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        initial_name = f"{DEFAULT_FILENAME_PREFIX}_{timestamp}.txt"

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.isdir(desktop):
            desktop = os.path.expanduser("~")

        file_path = filedialog.asksaveasfilename(
            title="Сохранить файл",
            initialdir=desktop,
            initialfile=initial_name,
            defaultextension=".txt",
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8-sig") as f:
                f.write(self.last_report)
            self.status_var.set(f"Файл сохранён: {file_path}")
            messagebox.showinfo("Успех", f"Файл успешно сохранён:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл.\n\n{e}")

    def on_copy(self):
        if not self.last_report.strip():
            messagebox.showwarning("Нет данных", "Сначала соберите характеристики.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_report)
        self.root.update()
        self.status_var.set("Текст скопирован в буфер обмена.")
        messagebox.showinfo("Готово", "Текст скопирован в буфер обмена.")


def main():
    if os.name != "nt":
        print("Этот скрипт рассчитан на Windows.")
        sys.exit(1)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
