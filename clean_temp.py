"""
ATCleaner — очистка временных файлов Windows и базовые инструменты оптимизации.

Вкладка "Очистка" удаляет:
- Temp пользователя (%TEMP%) и системный Windows\\Temp
- Логи Windows (C:\\Windows\\Logs)
- Кэш обновлений Windows (C:\\Windows\\SoftwareDistribution\\Download)
- Кэш миниатюр Explorer и общий кэш интернета (INetCache)
- Кэш DNS-резолвера (команда ipconfig /flushdns)
- Кэш Discord (только папки с картинками/вложениями — Code Cache НЕ трогаем,
  чтобы не сломать структуру Electron-приложения)
- Кэш Telegram, Steam
- Кэш браузеров: Google Chrome, Яндекс Браузер, Opera
- Кэш игровых лаунчеров: Epic Games Launcher, EA App
- Кэш видеоредакторов: CapCut, Adobe (Media Cache)
- Кэш шейдеров видеокарт: NVIDIA (GLCache, DXCache), AMD (DxCache, DxcCache, GLCache)

Вкладка "Инструменты" добавляет:
- "Освободить ОЗУ" — очистка Standby List (кэша ОЗУ) через системный вызов
  NtSetSystemInformation (тот же приём, что использует утилита
  "Empty Standby List"). Требует прав администратора — если их нет,
  программа сообщает об этом и не пытается выполнить системный вызов.
- "Перезапустить проводник" — принудительно завершает explorer.exe и
  запускает его заново (полезно, если завис рабочий стол/панель задач).

Программа НЕ трогает файлы, которые не может удалить (заняты другим
процессом или защищены системой) — такие объекты просто пропускаются.
Если папка какого-то приложения не найдена (программа не установлена
или видеокарта другого производителя), она также пропускается без ошибки.

ВАЖНО про Discord: удаляется только содержимое папок "Cache" (картинки,
вложения, эмодзи). Папка "Code Cache" НИКОГДА не трогается — в ней Electron
хранит скомпилированный JS-код интерфейса, и её удаление приводит к ошибке
запуска Discord.

Требуется библиотека customtkinter:
    pip install customtkinter

Иконка приложения:
    Положите файл ATCleanerLogoMR.ico рядом со скриптом, чтобы иконка
    отображалась в заголовке окна и на панели задач. Если файла нет —
    программа запустится без кастомной иконки, без ошибок.

Рекомендуется запускать от имени администратора — иначе часть системных
папок (Windows\\Logs, SoftwareDistribution\\Download, часть Windows\\Temp)
очистится не полностью, а функция "Освободить ОЗУ" не сработает вовсе.
"""

import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time

import customtkinter as ctk
from tkinter import messagebox

# ---------------------------------------------------------------------------
# Цветовая схема (тёмная тема)
# ---------------------------------------------------------------------------
COLOR_BG = "#120C1F"          # глубокий тёмно-фиолетовый фон окна
COLOR_TEXT = "#FFFFFF"        # белый текст
COLOR_SUBTEXT = "#B9AEDA"     # приглушённый лавандовый для подзаголовка
COLOR_LOG_BG = "#211A33"      # тёмно-серый (с фиолетовым оттенком) фон лога
COLOR_LOG_BORDER = "#4B3F73"  # аккуратная фиолетовая рамка лога
COLOR_LOG_TEXT = "#E8E2F5"    # светлый текст в логе
COLOR_ACCENT = "#B026FF"      # яркий неоново-фиолетовый (основная кнопка)
COLOR_ACCENT_HOVER = "#8E00E0"  # цвет кнопки при наведении/нажатии
COLOR_TOOL_ACCENT = "#7A5CFF"    # чуть более спокойный фиолетовый для кнопок-инструментов
COLOR_TOOL_ACCENT_HOVER = "#5F3FE0"
COLOR_TAB_BG = "#1B1530"      # фон панели вкладок
COLOR_FOOTER = "#FFFFFF"      # подпись автора — явно белый цвет
COLOR_WARNING = "#FFC94D"     # яркий тёплый жёлто-оранжевый для предупреждений

# Путь к иконке приложения (.ico). Если файла не существует — просто
# пропускаем установку иконки, без падения программы.
ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ATCleanerLogoMR.ico")


# ---------------------------------------------------------------------------
# Вспомогательные системные функции
# ---------------------------------------------------------------------------
def is_windows():
    return sys.platform.startswith("win")


def is_admin():
    """Проверяет, запущена ли программа с правами администратора."""
    if not is_windows():
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Список папок для очистки
# ---------------------------------------------------------------------------
def get_target_folders():
    """Возвращает список (название, путь) папок, которые будем чистить."""
    folders = []

    # Папка временных файлов текущего пользователя (%TEMP%)
    user_temp = tempfile.gettempdir()
    if user_temp:
        folders.append(("Temp пользователя", user_temp))

    windir = os.environ.get("WINDIR", r"C:\Windows")

    # Системная папка C:\Windows\Temp
    folders.append(("Системный Temp", os.path.join(windir, "Temp")))

    # Логи Windows
    folders.append(("Логи Windows", os.path.join(windir, "Logs")))

    # Кэш загруженных обновлений Windows (обычно занимает несколько ГБ)
    folders.append((
        "Кэш обновлений Windows",
        os.path.join(windir, "SoftwareDistribution", "Download"),
    ))

    local_appdata = os.environ.get("LOCALAPPDATA")
    roaming_appdata = os.environ.get("APPDATA")
    program_files_x86 = os.environ.get("ProgramFiles(x86)") or os.environ.get("ProgramFiles")

    if local_appdata:
        # Кэш миниатюр Explorer
        folders.append((
            "Кэш миниатюр Explorer",
            os.path.join(local_appdata, "Microsoft", "Windows", "Explorer"),
        ))

        # Общий кэш интернета
        folders.append((
            "Кэш браузера (INetCache)",
            os.path.join(local_appdata, "Microsoft", "Windows", "INetCache"),
        ))

        # --- Discord ---------------------------------------------------
        # ВАЖНО: чистим ТОЛЬКО папку "Cache" (картинки, вложения, эмодзи).
        # Папку "Code Cache" никогда не трогаем — в ней Electron хранит
        # скомпилированный JS-код интерфейса, и её удаление ломает запуск
        # Discord ("не открывается", "белый экран" и т.п.).
        folders.append(("Кэш Discord (LocalAppData)", os.path.join(local_appdata, "discord", "Cache")))

        # --- Браузеры ---
        folders.append((
            "Кэш Google Chrome",
            os.path.join(local_appdata, "Google", "Chrome", "User Data", "Default", "Cache"),
        ))
        folders.append((
            "Кэш Яндекс Браузера",
            os.path.join(local_appdata, "Yandex", "YandexBrowser", "User Data", "Default", "Cache"),
        ))
        folders.append((
            "Кэш Opera",
            os.path.join(local_appdata, "Opera Software", "Opera Stable", "Cache"),
        ))

        # --- Игровые лаунчеры ---
        folders.append((
            "Кэш Epic Games Launcher",
            os.path.join(local_appdata, "EpicGamesLauncher", "Saved", "webcache"),
        ))
        folders.append((
            "Кэш EA App",
            os.path.join(local_appdata, "Electronic Arts", "EA Desktop", "CacheStorage"),
        ))

        # --- Видеоредакторы ---
        folders.append((
            "Кэш CapCut",
            os.path.join(local_appdata, "CapCut", "User Data", "Cache"),
        ))

        # --- Кэш шейдеров видеокарт ---
        # NVIDIA: кэш скомпилированных OpenGL- и DirectX-шейдеров
        folders.append(("Кэш шейдеров NVIDIA (OpenGL)", os.path.join(local_appdata, "NVIDIA", "GLCache")))
        folders.append(("Кэш шейдеров NVIDIA (DirectX)", os.path.join(local_appdata, "NVIDIA", "DXCache")))

        # AMD: кэш шейдеров драйвера (пути отличаются в зависимости от версии
        # драйвера — пробуем самые распространённые; несуществующие просто
        # пропускаются без ошибок)
        folders.append(("Кэш шейдеров AMD (DirectX)", os.path.join(local_appdata, "AMD", "DxCache")))
        folders.append(("Кэш шейдеров AMD (DXC)", os.path.join(local_appdata, "AMD", "DxcCache")))
        folders.append(("Кэш шейдеров AMD (OpenGL)", os.path.join(local_appdata, "AMD", "GLCache")))

    if roaming_appdata:
        # --- Discord (Roaming) ---
        # Как и выше — чистим только "Cache", "Code Cache" не трогаем.
        folders.append(("Кэш Discord (Roaming)", os.path.join(roaming_appdata, "discord", "Cache")))

        # Кэш Telegram Desktop
        folders.append((
            "Кэш Telegram",
            os.path.join(roaming_appdata, "Telegram Desktop", "tdata", "user_data", "cache"),
        ))
        folders.append((
            "Медиа-кэш Telegram",
            os.path.join(roaming_appdata, "Telegram Desktop", "tdata", "user_data", "media_cache"),
        ))

        # Кэш Adobe (Premiere / After Effects и т.д.)
        folders.append((
            "Медиа-кэш Adobe",
            os.path.join(roaming_appdata, "Adobe", "Common", "Media Cache"),
        ))
        folders.append((
            "Файлы медиа-кэша Adobe",
            os.path.join(roaming_appdata, "Adobe", "Common", "Media Cache Files"),
        ))

    if program_files_x86:
        # Кэш Steam
        folders.append(("Кэш Steam (appcache)", os.path.join(program_files_x86, "Steam", "appcache")))
        folders.append(("Кэш Steam (htmlcache)", os.path.join(program_files_x86, "Steam", "htmlcache")))

    return folders


def format_size(num_bytes):
    for unit in ["Б", "КБ", "МБ", "ГБ"]:
        if num_bytes < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} ТБ"


def clean_folder(path, log_callback):
    """Удаляет содержимое папки (не саму папку), считает освобождённый объём.

    Если папки не существует (например, приложение не установлено, или
    видеокарта другого производителя), она просто пропускается без ошибки.
    """
    freed = 0
    deleted_count = 0
    error_count = 0

    if not os.path.isdir(path):
        log_callback(f"  Пропущено (не найдено): {path}")
        return freed, deleted_count, error_count

    for entry in os.scandir(path):
        entry_path = entry.path
        try:
            if entry.is_file(follow_symlinks=False) or entry.is_symlink():
                size = entry.stat(follow_symlinks=False).st_size
                os.remove(entry_path)
                freed += size
                deleted_count += 1
            elif entry.is_dir(follow_symlinks=False):
                size = get_dir_size(entry_path)
                shutil.rmtree(entry_path)
                freed += size
                deleted_count += 1
        except (PermissionError, OSError):
            # Файл используется системой/другим процессом — пропускаем
            error_count += 1

    return freed, deleted_count, error_count


def get_dir_size(path):
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def flush_dns():
    """Выполняет ipconfig /flushdns. Возвращает (успех, сообщение)."""
    if not is_windows():
        return False, "Команда доступна только на Windows."
    try:
        result = subprocess.run(
            ["ipconfig", "/flushdns"],
            capture_output=True,
            text=True,
            shell=False,
        )
        output = (result.stdout or result.stderr or "").strip()
        if result.returncode == 0:
            return True, output or "Кэш DNS успешно очищен."
        return False, output or "Команда завершилась с ошибкой."
    except Exception as e:
        return False, f"Ошибка при выполнении команды: {e}"


# ---------------------------------------------------------------------------
# Инструменты оптимизации
# ---------------------------------------------------------------------------
def clear_standby_list():
    """
    Очищает Standby List (кэшированные страницы ОЗУ) через недокументированный
    системный вызов NtSetSystemInformation — тот же приём, что используют
    утилиты вроде "Empty Standby List". Требует прав администратора и
    включения привилегии SeProfileSingleProcessPrivilege.

    Возвращает (успех: bool, сообщение: str).
    """
    if not is_windows():
        return False, "Доступно только на Windows."
    if not is_admin():
        return False, "[Ошибка] Запустите от имени Администратора для очистки ОЗУ!"

    try:
        ntdll = ctypes.WinDLL("ntdll.dll")
        advapi32 = ctypes.WinDLL("advapi32.dll")
        kernel32 = ctypes.WinDLL("kernel32.dll")

        TOKEN_ADJUST_PRIVILEGES = 0x0020
        TOKEN_QUERY = 0x0008
        SE_PRIVILEGE_ENABLED = 0x00000002

        class LUID(ctypes.Structure):
            _fields_ = [("LowPart", ctypes.c_uint32), ("HighPart", ctypes.c_int32)]

        class LUID_AND_ATTRIBUTES(ctypes.Structure):
            _fields_ = [("Luid", LUID), ("Attributes", ctypes.c_uint32)]

        class TOKEN_PRIVILEGES(ctypes.Structure):
            _fields_ = [("PrivilegeCount", ctypes.c_uint32),
                        ("Privileges", LUID_AND_ATTRIBUTES * 1)]

        h_token = ctypes.c_void_p()
        if not advapi32.OpenProcessToken(
            kernel32.GetCurrentProcess(),
            TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
            ctypes.byref(h_token),
        ):
            return False, "Не удалось открыть токен текущего процесса."

        luid = LUID()
        if not advapi32.LookupPrivilegeValueW(
            None, "SeProfileSingleProcessPrivilege", ctypes.byref(luid)
        ):
            return False, "Не удалось найти привилегию SeProfileSingleProcessPrivilege."

        tp = TOKEN_PRIVILEGES()
        tp.PrivilegeCount = 1
        tp.Privileges[0].Luid = luid
        tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

        if not advapi32.AdjustTokenPrivileges(
            h_token, False, ctypes.byref(tp), 0, None, None
        ):
            return False, "Не удалось включить необходимую привилегию."

        # SystemMemoryListInformation = 0x50, MemoryPurgeStandbyList = 4
        command = ctypes.c_int(4)
        status = ntdll.NtSetSystemInformation(0x50, ctypes.byref(command), ctypes.sizeof(command))

        if status == 0:
            return True, "Standby List (кэш ОЗУ) успешно очищен."
        return False, f"Системный вызов вернул код ошибки: {status}"

    except Exception as e:
        return False, f"Ошибка при очистке ОЗУ: {e}"


def restart_explorer():
    """Принудительно перезапускает explorer.exe. Возвращает (успех, сообщение)."""
    if not is_windows():
        return False, "Доступно только на Windows."
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "explorer.exe"],
            capture_output=True,
            text=True,
        )
        time.sleep(1.5)
        subprocess.Popen("explorer.exe")
        return True, "Проводник (explorer.exe) успешно перезапущен."
    except Exception as e:
        return False, f"Ошибка при перезапуске проводника: {e}"


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------
class CleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")

        self.title("ATCleaner — очистка временных файлов Windows")

        # Принудительно закрепляем иконку сразу после создания окна — так
        # Windows с большей вероятностью подхватит её и для заголовка окна,
        # и для панели задач.
        try:
            self.iconbitmap("ATCleanerLogoMR.ico")
        except Exception:
            pass  # файл не найден рядом со скриптом или мы не на Windows

        # Окно шире и выше, чтобы вкладки и новые кнопки смотрелись свободно
        self.geometry("760x700")
        self.resizable(False, False)
        self.configure(fg_color=COLOR_BG)

        # CustomTkinter иногда сбрасывает иконку окна на стандартную уже
        # после отрисовки интерфейса — переустанавливаем её ещё раз чуть позже.
        self._set_app_icon()
        self.after(200, self._set_app_icon)

        # --- Заголовок -------------------------------------------------
        title = ctk.CTkLabel(
            self,
            text="ATCleaner — очистка кэша и оптимизация Windows",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=COLOR_TEXT,
        )
        title.pack(pady=(20, 6))

        subtitle = ctk.CTkLabel(
            self,
            text="Temp, логи Windows, кэш обновлений, DNS, Discord, Telegram, Steam,\n"
                 "браузеров, лаунчеров, видеоредакторов и шейдеров NVIDIA/AMD.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLOR_SUBTEXT,
            justify="center",
        )
        subtitle.pack(pady=(0, 10))

        # --- Предупреждения ------------------------------------------------
        warning_label = ctk.CTkLabel(
            self,
            text="⚠️ Рекомендуется запускать программу от имени администратора\n"
                 "⚠️ Для лучшего результата закройте Telegram, Discord и Steam перед очисткой",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=COLOR_WARNING,
            justify="center",
        )
        warning_label.pack(pady=(0, 14))

        # --- Вкладки ------------------------------------------------------
        self.tabview = ctk.CTkTabview(
            self,
            width=700,
            height=110,
            fg_color=COLOR_TAB_BG,
            segmented_button_fg_color=COLOR_TAB_BG,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_TAB_BG,
            text_color=COLOR_TEXT,
            corner_radius=10,
        )
        self.tabview.pack(padx=16, pady=(0, 14), fill="x")

        tab_clean = self.tabview.add("Очистка")
        tab_tools = self.tabview.add("Инструменты")
        tab_clean.configure(fg_color=COLOR_TAB_BG)
        tab_tools.configure(fg_color=COLOR_TAB_BG)

        # --- Вкладка "Очистка" ---------------------------------------------
        self.clean_button = ctk.CTkButton(
            tab_clean,
            text="🧹  Очистить сейчас",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="white",
            corner_radius=10,
            width=240,
            height=44,
            command=self.start_cleaning,
        )
        self.clean_button.pack(pady=18)

        # --- Вкладка "Инструменты" -------------------------------------------
        tools_row = ctk.CTkFrame(tab_tools, fg_color=COLOR_TAB_BG)
        tools_row.pack(pady=14)

        self.ram_button = ctk.CTkButton(
            tools_row,
            text="🧠  Освободить ОЗУ",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=COLOR_TOOL_ACCENT,
            hover_color=COLOR_TOOL_ACCENT_HOVER,
            text_color="white",
            corner_radius=10,
            width=220,
            height=42,
            command=self.start_ram_clean,
        )
        self.ram_button.pack(side="left", padx=10)

        self.explorer_button = ctk.CTkButton(
            tools_row,
            text="🗂️  Перезапустить проводник",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=COLOR_TOOL_ACCENT,
            hover_color=COLOR_TOOL_ACCENT_HOVER,
            text_color="white",
            corner_radius=10,
            width=240,
            height=42,
            command=self.start_explorer_restart,
        )
        self.explorer_button.pack(side="left", padx=10)

        # --- Поле логов (общее для обеих вкладок) --------------------------
        self.log_area = ctk.CTkTextbox(
            self,
            width=700,
            height=290,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=COLOR_LOG_BG,
            text_color=COLOR_LOG_TEXT,
            border_width=1,
            border_color=COLOR_LOG_BORDER,
            corner_radius=8,
            scrollbar_button_color=COLOR_LOG_BORDER,
            scrollbar_button_hover_color=COLOR_ACCENT,
            state="disabled",
        )
        self.log_area.pack(padx=16, pady=(0, 14))

        # --- Статус ------------------------------------------------------
        self.status_label = ctk.CTkLabel(
            self,
            text="Готово к работе.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLOR_TEXT,
        )
        self.status_label.pack(pady=(0, 10))

        # --- Подпись автора --------------------------------------------------
        bottom_bar = ctk.CTkFrame(self, fg_color=COLOR_BG)
        bottom_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 12))

        footer = ctk.CTkLabel(
            bottom_bar,
            text="Создано от ArseniosTs",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=COLOR_FOOTER,
        )
        footer.pack(side="right")

    # -----------------------------------------------------------------
    # Общие утилиты интерфейса
    # -----------------------------------------------------------------
    def _set_app_icon(self):
        """Устанавливает иконку окна и панели задач, если файл найден."""
        if is_windows() and os.path.isfile(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass  # некорректный .ico или иная проблема — не критично

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def clear_log(self):
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")

    def set_buttons_state(self, state):
        self.clean_button.configure(state=state)
        self.ram_button.configure(state=state)
        self.explorer_button.configure(state=state)

    # -----------------------------------------------------------------
    # Очистка мусора
    # -----------------------------------------------------------------
    def start_cleaning(self):
        self.set_buttons_state("disabled")
        self.clean_button.configure(text="Очистка...")
        self.status_label.configure(text="Идёт очистка, подождите...")
        self.clear_log()

        thread = threading.Thread(target=self.run_cleaning, daemon=True)
        thread.start()

    def run_cleaning(self):
        total_freed = 0
        total_deleted = 0
        total_errors = 0

        # --- Кэш DNS ---
        self.log("Очистка кэша DNS (ipconfig /flushdns)")
        dns_ok, dns_message = flush_dns()
        self.log(f"  {dns_message}")
        self.log("")

        # --- Папки ---
        for label, path in get_target_folders():
            self.log(f"Очистка: {label}")
            self.log(f"  Путь: {path}")
            freed, deleted, errors = clean_folder(path, self.log)
            total_freed += freed
            total_deleted += deleted
            total_errors += errors
            self.log(
                f"  Удалено объектов: {deleted}, "
                f"освобождено: {format_size(freed)}, "
                f"пропущено (занято/защищено): {errors}"
            )
            self.log("")

        self.log("=" * 50)
        self.log(f"ИТОГО удалено объектов: {total_deleted}")
        self.log(f"ИТОГО освобождено места: {format_size(total_freed)}")
        if total_errors:
            self.log(
                f"Пропущено {total_errors} объектов (использовались системой — "
                f"попробуйте запустить от имени администратора)."
            )

        self.after(0, self.finish_cleaning, total_freed, total_deleted)

    def finish_cleaning(self, total_freed, total_deleted):
        self.set_buttons_state("normal")
        self.clean_button.configure(text="🧹  Очистить сейчас")
        self.status_label.configure(
            text=f"Готово. Удалено {total_deleted} объектов, освобождено {format_size(total_freed)}."
        )
        messagebox.showinfo(
            "Очистка завершена",
            f"Удалено объектов: {total_deleted}\n"
            f"Освобождено места: {format_size(total_freed)}",
        )

    # -----------------------------------------------------------------
    # Инструмент: освобождение ОЗУ
    # -----------------------------------------------------------------
    def start_ram_clean(self):
        self.clear_log()

        # Явная проверка прав администратора ДО запуска потока — если их
        # нет, сразу сообщаем пользователю и не пытаемся выполнять системный
        # вызов, который всё равно завершится неудачей.
        if not is_admin():
            self.log("[Ошибка] Запустите от имени Администратора для очистки ОЗУ!")
            self.status_label.configure(text="Нужны права администратора для очистки ОЗУ.")
            messagebox.showwarning(
                "Требуются права администратора",
                "Для очистки ОЗУ необходимо запустить программу от имени администратора.\n\n"
                "Закройте программу и запустите её снова через правый клик →\n"
                "«Запуск от имени администратора».",
            )
            return

        self.set_buttons_state("disabled")
        self.ram_button.configure(text="Очистка ОЗУ...")
        self.status_label.configure(text="Освобождаем оперативную память...")
        self.log("Очистка Standby List (кэша оперативной памяти)...")

        thread = threading.Thread(target=self.run_ram_clean, daemon=True)
        thread.start()

    def run_ram_clean(self):
        success, message = clear_standby_list()
        self.log(f"  {message}")
        self.after(0, self.finish_ram_clean, success, message)

    def finish_ram_clean(self, success, message):
        self.set_buttons_state("normal")
        self.ram_button.configure(text="🧠  Освободить ОЗУ")
        self.status_label.configure(
            text="ОЗУ очищена." if success else "Не удалось очистить ОЗУ."
        )
        if success:
            messagebox.showinfo("Готово", message)
        else:
            messagebox.showwarning("Не удалось", message)

    # -----------------------------------------------------------------
    # Инструмент: перезапуск проводника
    # -----------------------------------------------------------------
    def start_explorer_restart(self):
        self.set_buttons_state("disabled")
        self.explorer_button.configure(text="Перезапуск...")
        self.status_label.configure(text="Перезапускаем проводник...")
        self.clear_log()
        self.log("Перезапуск процесса explorer.exe...")

        thread = threading.Thread(target=self.run_explorer_restart, daemon=True)
        thread.start()

    def run_explorer_restart(self):
        success, message = restart_explorer()
        self.log(f"  {message}")
        self.after(0, self.finish_explorer_restart, success, message)

    def finish_explorer_restart(self, success, message):
        self.set_buttons_state("normal")
        self.explorer_button.configure(text="🗂️  Перезапустить проводник")
        self.status_label.configure(
            text="Проводник перезапущен." if success else "Не удалось перезапустить проводник."
        )
        if success:
            messagebox.showinfo("Готово", message)
        else:
            messagebox.showwarning("Не удалось", message)


if __name__ == "__main__":
    app = CleanerApp()
    app.mainloop()
