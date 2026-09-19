"""
ATCleaner v1.5 — очистка временных файлов Windows, инструменты оптимизации
и поддержка ПЯТИ языков интерфейса: English (по умолчанию), Русский,
Ελληνικά, Deutsch, Español.

Вкладка "Очистка" / "Cleaning" содержит только лог и кнопку запуска.
Вкладка "Настройки" / "Settings" содержит флажки (CTkCheckBox),
сгруппированные по категориям, — пользователь сам решает, что чистить.
Если флажок снят, соответствующая папка полностью пропускается.

Категории и флажки:
  System / Система:
    - Temporary Files / Временные файлы  (Temp пользователя, системный Temp,
      кэш миниатюр Explorer, INetCache)
    - System Logs / Системные логи        (C:\\Windows\\Logs)
    - Windows Update Cache / Кэш обновлений Windows
      (C:\\Windows\\SoftwareDistribution\\Download)
    - Windows.old Folder / Папка Windows.old (C:\\Windows.old — файлы
      предыдущей установки Windows; безопасно удалять, если система
      работает нормально; освобождает 10-30 ГБ)
  Browsers / Браузеры:
    - Google Chrome, Yandex Browser, Opera & Opera GX, Microsoft Edge
  Apps / Приложения:
    - Telegram, Discord (безопасно — без Code Cache), Steam,
      Graphics Shaders NVIDIA/AMD

Дополнительно (не отображается как флажок, но чистится всегда):
  - Кэш DNS (ipconfig /flushdns)

Вкладка "Инструменты" / "Tools" (сетка 2x2):
  - "Восстановить кэш иконок" / "Rebuild Icon Cache" — завершает explorer.exe,
    удаляет файлы кэша иконок (%LOCALAPPDATA%\\IconCache.db и
    %LOCALAPPDATA%\\Microsoft\\Windows\\Explorer\\iconcache_*.db), затем
    запускает explorer.exe заново. Решает проблему белых/пустых иконок и
    их долгой загрузки.
  - "Перезапустить Проводник" / "Restart Explorer" — принудительно
    завершает explorer.exe и запускает его заново.
  - "Очистить буфер обмена" / "Clear Clipboard" — очищает системный буфер
    обмена через Win32 API (OpenClipboard → EmptyClipboard → CloseClipboard).

Вкладка "О программе" / "About": название и версия программы, автор,
интерактивная кликабельная ссылка на GitHub-репозиторий и текст дисклеймера
на выбранном языке.

ВАЖНО про Discord: чистится только содержимое папок "Cache" (картинки,
вложения, эмодзи). Папка "Code Cache" НИКОГДА не трогается — в ней Electron
хранит скомпилированный JS-код интерфейса, и её удаление ломает запуск
Discord.

ВАЖНО про Windows.old: удаление выполняется тем же безопасным методом, что
и остальные папки — файлы, к которым нет доступа (например, защищённые
TrustedInstaller), просто пропускаются без падения программы. Для полного
удаления Windows.old может потребоваться запуск от имени администратора.

Переключение языка: выпадающий список в ЛЕВОМ нижнем углу окна
(🇬🇧/🇷🇺/🇬🇷/🇩🇪/🇪🇸) мгновенно переводит весь интерфейс — заголовки,
вкладки, кнопки, флажки, крупные предупреждения, статус-строку, подписи
в логе и вкладку "О программе".

Требуется библиотека customtkinter:
    pip install customtkinter

Иконка приложения:
    Положите файл ATCleanerLogoMR.ico рядом со скриптом — иконка жёстко
    закрепляется через self.iconbitmap("ATCleanerLogoMR.ico").
"""

import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser

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
COLOR_CARD_BG = "#1E1832"     # фон карточек/списка флажков
COLOR_FOOTER = "#FFFFFF"      # подпись автора — явно белый цвет
COLOR_WARNING = "#FFC94D"     # яркий тёплый жёлто-оранжевый для предупреждений
COLOR_CATEGORY = "#D6B8FF"    # цвет заголовков категорий флажков

# Путь к иконке приложения (.ico).
ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ATCleanerLogoMR.ico")

# Путь к файлу конфигурации (хранит только выбранный язык, одна строка).
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.txt")


def load_saved_language():
    """Читает config.txt и возвращает сохранённый код языка (ru/en/gr/de/es),
    если файл существует и содержит валидный код. Иначе возвращает None —
    вызывающий код в этом случае использует DEFAULT_LANG (English)."""
    try:
        if os.path.isfile(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = f.read().strip().lower()
            if saved in LANG_OPTIONS.values():
                return saved
    except OSError:
        pass
    return None


def save_language(lang_code):
    """Записывает выбранный код языка в config.txt (перезаписывает файл).
    Ошибки записи (например, папка без прав на запись) тихо игнорируются —
    это не критичная функция, программа должна работать и без неё."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write(lang_code)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Словарь переводов
# ---------------------------------------------------------------------------
LANG_OPTIONS = {
    "🇬🇧 English": "en",
    "🇷🇺 Русский": "ru",
    "🇬🇷 Ελληνικά": "gr",
    "🇩🇪 Deutsch": "de",
    "🇪🇸 Español": "es",
}
LANG_OPTION_LIST = list(LANG_OPTIONS.keys())

TR = {
    "ru": {
        "window_title": "ATCleaner — очистка временных файлов Windows",
        "main_title": "ATCleaner — очистка кэша и оптимизация Windows",
        "subtitle": "Выберите, что очистить, ниже. Кэш DNS обновляется\n"
                     "автоматически при каждом запуске очистки.",
        "warning1": "⚠️ Рекомендуется запускать от имени администратора",
        "warning2": "⚠️ Закройте Telegram, Discord и Steam перед очисткой",
        "tab_clean": "Очистка",
        "tab_settings": "Настройки",
        "tab_tools": "Инструменты",
        "tab_about": "О программе",
        "utilities_title": "Системные утилиты",
        "about_title": "ATCleaner v1.5",
        "about_author": "Автор: ArseniosTs",
        "about_github_label": "🌐 Репозиторий на GitHub",
        "about_disclaimer": "Программа изменяет системные файлы и кэш приложений.\n"
                             "Используйте на свой риск. Перед серьёзными изменениями\n"
                             "рекомендуется иметь резервную копию важных данных.",
        "clean_button": "🧹  Очистить сейчас",
        "clean_button_working": "Очистка...",
        "icon_cache_button": "🖼️  Восстановить кэш иконок",
        "icon_cache_button_working": "Восстановление...",
        "explorer_button": "🗂️  Перезапустить Проводник",
        "explorer_button_working": "Перезапуск...",
        "tool_desc_explorer": "Перезапуск проводника, если зависла панель задач.",
        "tool_desc_icon_cache": "Исправляет белые или неверно отображаемые значки.",
        "tool_desc_clipboard": "Мгновенно удаляет данные буфера обмена из памяти.",
        "lang_switch_hint": "Мгновенное переключение между 5 языками",
        "clipboard_button": "📋  Очистить буфер обмена",
        "clipboard_button_working": "Очистка...",
        "footer": "Создано от ArseniosTs",
        "status_ready": "Готово к работе.",
        "status_cleaning": "Идёт очистка, подождите...",
        "status_done": "Готово. Удалено {count} объектов, освобождено {size}.",
        "status_icon_cache_working": "Восстанавливаем кэш иконок...",
        "status_icon_cache_done": "Кэш иконок восстановлен. Проводник перезапущен.",
        "status_icon_cache_failed": "Не удалось восстановить кэш иконок.",
        "status_explorer_working": "Перезапускаем проводник...",
        "status_explorer_done": "Проводник перезапущен.",
        "status_explorer_failed": "Не удалось перезапустить проводник.",
        "status_clipboard_working": "Очищаем буфер обмена...",
        "status_clipboard_done": "Буфер обмена очищен.",
        "status_clipboard_failed": "Не удалось очистить буфер обмена.",
        "log_dns_title": "Очистка кэша DNS (ipconfig /flushdns)",
        "log_cleaning": "Очистка: {label}",
        "log_path": "  Путь: {path}",
        "log_skipped_missing": "  Пропущено (не найдено): {path}",
        "log_stats": "  Удалено: {deleted}, Освобождено: {freed}, Пропущено: {errors}",
        "log_total_deleted": "ИТОГО удалено объектов: {count}",
        "log_total_freed": "ИТОГО освобождено места: {size}",
        "log_total_skipped_note": "Пропущено {count} объектов (использовались системой — "
                                   "попробуйте запустить от имени администратора).",
        "log_icon_cache_title": "Восстановление кэша иконок (перезапуск проводника)...",
        "log_explorer_title": "Перезапуск процесса explorer.exe...",
        "log_clipboard_title": "Очистка буфера обмена (EmptyClipboard)...",
        "log_user_skipped": "⏭ Пропущено пользователем (флажок снят): {label}",
        "log_extra_section_title": "Дополнительно (лаунчеры, редакторы):",
        "msg_clean_done_title": "Очистка завершена",
        "msg_clean_done_body": "Удалено объектов: {count}\nОсвобождено места: {size}",
        "msg_done_title": "Готово",
        "msg_failed_title": "Не удалось",
        # --- Флажки очистки ---
        "chk_section_hint": "Выберите источники для очистки:",
        "cat_system": "Система",
        "cat_browsers": "Браузеры",
        "cat_apps": "Приложения",
        "chk_temp_files": "Временные файлы",
        "chk_system_logs": "Системные логи",
        "chk_windows_update": "Кэш обновлений Windows",
        "chk_windows_old": "Папка Windows.old",
        "desc_windows_old": "Хранит файлы предыдущей установки Windows. Можно безопасно\n"
                             "удалить, если система работает нормально. Освобождает 10–30 ГБ.",
        "chk_chrome": "Google Chrome",
        "chk_yandex": "Яндекс Браузер",
        "chk_opera_stable": "Opera Stable",
        "chk_opera_gx": "Opera GX",
        "chk_edge": "Microsoft Edge",
        "chk_firefox": "Mozilla Firefox",
        "chk_telegram": "Telegram",
        "chk_discord": "Discord",
        "chk_steam": "Steam",
        "chk_epic": "Epic Games",
        "chk_ea": "EA App",
        "chk_spotify": "Spotify",
        "chk_capcut": "CapCut",
        "chk_adobe_photoshop": "Adobe Photoshop",
        "chk_adobe_premiere": "Adobe Premiere",
        "chk_shaders": "GPU Shaders (Кэш видеокарты NVIDIA/AMD)",
    },
    "en": {
        "window_title": "ATCleaner — Windows Temporary File Cleaner",
        "main_title": "ATCleaner — Cache Cleaner & Windows Optimizer",
        "subtitle": "Choose what to clean below. DNS, launchers, and video editor\n"
                     "caches are cleaned automatically on every run.",
        "warning1": "⚠️ Running as Administrator is highly recommended.",
        "warning2": "⚠️ Close Telegram, Discord, and Steam before cleaning for best results.",
        "tab_clean": "Cleaning",
        "tab_settings": "Settings",
        "tab_tools": "Tools",
        "tab_about": "About",
        "utilities_title": "System Utilities",
        "about_title": "ATCleaner v1.5",
        "about_author": "Author: ArseniosTs",
        "about_github_label": "🌐 GitHub Repository",
        "about_disclaimer": "This tool modifies system and application cache files.\n"
                             "Use at your own risk. Keeping a backup of important data\n"
                             "before major changes is always a good idea.",
        "clean_button": "🧹  Clean Now",
        "clean_button_working": "Cleaning...",
        "icon_cache_button": "🖼️  Rebuild Icon Cache",
        "icon_cache_button_working": "Rebuilding...",
        "explorer_button": "🗂️  Restart Explorer",
        "explorer_button_working": "Restarting...",
        "tool_desc_explorer": "Restarts File Explorer if the taskbar freezes up.",
        "tool_desc_icon_cache": "Fixes icons that appear blank, white, or wrong.",
        "tool_desc_clipboard": "Instantly wipes clipboard data from memory.",
        "lang_switch_hint": "Instant switching between 5 languages",
        "clipboard_button": "📋  Clear Clipboard",
        "clipboard_button_working": "Clearing...",
        "footer": "Created by ArseniosTs",
        "status_ready": "Ready to work.",
        "status_cleaning": "Cleaning in progress, please wait...",
        "status_done": "Done. Deleted {count} items, freed {size}.",
        "status_icon_cache_working": "Rebuilding icon cache...",
        "status_icon_cache_done": "Icon cache rebuilt. Explorer restarted.",
        "status_icon_cache_failed": "Failed to rebuild icon cache.",
        "status_explorer_working": "Restarting Explorer...",
        "status_explorer_done": "Explorer restarted.",
        "status_explorer_failed": "Failed to restart Explorer.",
        "status_clipboard_working": "Clearing clipboard...",
        "status_clipboard_done": "Clipboard cleared.",
        "status_clipboard_failed": "Failed to clear the clipboard.",
        "log_dns_title": "Flushing DNS cache (ipconfig /flushdns)",
        "log_cleaning": "Cleaning: {label}",
        "log_path": "  Path: {path}",
        "log_skipped_missing": "  Skipped (not found): {path}",
        "log_stats": "  Deleted: {deleted}, Total space freed: {freed}, Skipped: {errors}",
        "log_total_deleted": "TOTAL items deleted: {count}",
        "log_total_freed": "TOTAL space freed: {size}",
        "log_total_skipped_note": "Skipped {count} items (in use by the system — "
                                   "try running as Administrator).",
        "log_icon_cache_title": "Rebuilding icon cache (restarting Explorer)...",
        "log_explorer_title": "Restarting explorer.exe process...",
        "log_clipboard_title": "Clearing the clipboard (EmptyClipboard)...",
        "log_user_skipped": "⏭ Skipped by user (checkbox unchecked): {label}",
        "log_extra_section_title": "Additional (launchers, editors):",
        "msg_clean_done_title": "Cleaning Complete",
        "msg_clean_done_body": "Items deleted: {count}\nSpace freed: {size}",
        "msg_done_title": "Done",
        "msg_failed_title": "Failed",
        # --- Cleaning checkboxes ---
        "chk_section_hint": "Select what to clean:",
        "cat_system": "System",
        "cat_browsers": "Browsers",
        "cat_apps": "Apps",
        "chk_temp_files": "Temporary Files",
        "chk_system_logs": "System Logs",
        "chk_windows_update": "Windows Update Cache",
        "chk_windows_old": "Windows.old Folder",
        "desc_windows_old": "Stores files from your previous Windows installation. Safe to\n"
                             "delete if your system works fine. Frees 10-30 GB.",
        "chk_chrome": "Google Chrome",
        "chk_yandex": "Yandex Browser",
        "chk_opera_stable": "Opera Stable",
        "chk_opera_gx": "Opera GX",
        "chk_edge": "Microsoft Edge",
        "chk_firefox": "Mozilla Firefox",
        "chk_telegram": "Telegram",
        "chk_discord": "Discord",
        "chk_steam": "Steam",
        "chk_epic": "Epic Games",
        "chk_ea": "EA App",
        "chk_spotify": "Spotify",
        "chk_capcut": "CapCut",
        "chk_adobe_photoshop": "Adobe Photoshop",
        "chk_adobe_premiere": "Adobe Premiere",
        "chk_shaders": "GPU Shaders (NVIDIA/AMD Cache)",
    },
    "gr": {
        "window_title": "ATCleaner — Εκκαθάριση Προσωρινών Αρχείων Windows",
        "main_title": "ATCleaner — Εκκαθάριση Cache & Βελτιστοποίηση Windows",
        "subtitle": "Επιλέξτε τι θα καθαριστεί παρακάτω. Η cache DNS ανανεώνεται\n"
                     "αυτόματα σε κάθε εκκαθάριση.",
        "warning1": "⚠️ Συνιστάται η εκτέλεση ως Διαχειριστής.",
        "warning2": "⚠️ Κλείστε το Telegram, το Discord και το Steam πριν την εκκαθάριση.",
        "tab_clean": "Εκκαθάριση",
        "tab_settings": "Ρυθμίσεις",
        "tab_tools": "Εργαλεία",
        "tab_about": "Σχετικά",
        "utilities_title": "Εργαλεία Συστήματος",
        "about_title": "ATCleaner v1.5",
        "about_author": "Δημιουργός: ArseniosTs",
        "about_github_label": "🌐 Αποθετήριο GitHub",
        "about_disclaimer": "Αυτό το εργαλείο τροποποιεί αρχεία cache συστήματος και\n"
                             "εφαρμογών. Χρήση με δική σας ευθύνη. Συνιστάται αντίγραφο\n"
                             "ασφαλείας πριν από σημαντικές αλλαγές.",
        "clean_button": "🧹  Εκκαθάριση Τώρα",
        "clean_button_working": "Εκκαθάριση...",
        "icon_cache_button": "🖼️  Επαναφορά Cache Εικονιδίων",
        "icon_cache_button_working": "Επαναφορά...",
        "explorer_button": "🗂️  Επανεκκίνηση Explorer",
        "explorer_button_working": "Επανεκκίνηση...",
        "tool_desc_explorer": "Επανεκκίνηση της Διαχείρισης Αρχείων σε περίπτωση που κολλήσει η γραμμή εργασιών.",
        "tool_desc_icon_cache": "Επιδιόρθωση της προσωρινής μνήμης εικονιδίων.",
        "tool_desc_clipboard": "Άμεση διαγραφή των δεδομένων του προχείρου από τη μνήμη.",
        "lang_switch_hint": "Άμεση εναλλαγή 5 γλωσσών",
        "clipboard_button": "📋  Εκκαθάριση Προχείρου",
        "clipboard_button_working": "Εκκαθάριση...",
        "footer": "Δημιουργήθηκε από ArseniosTs",
        "status_ready": "Έτοιμο.",
        "status_cleaning": "Εκκαθάριση σε εξέλιξη, παρακαλώ περιμένετε...",
        "status_done": "Ολοκληρώθηκε. Διαγράφηκαν {count} στοιχεία, ελευθερώθηκαν {size}.",
        "status_icon_cache_working": "Επαναφορά cache εικονιδίων...",
        "status_icon_cache_done": "Η cache εικονιδίων επαναφέρθηκε. Ο Explorer επανεκκινήθηκε.",
        "status_icon_cache_failed": "Αποτυχία επαναφοράς cache εικονιδίων.",
        "status_explorer_working": "Επανεκκίνηση Explorer...",
        "status_explorer_done": "Ο Explorer επανεκκινήθηκε.",
        "status_explorer_failed": "Αποτυχία επανεκκίνησης Explorer.",
        "status_clipboard_working": "Εκκαθάριση προχείρου...",
        "status_clipboard_done": "Το πρόχειρο εκκαθαρίστηκε.",
        "status_clipboard_failed": "Αποτυχία εκκαθάρισης προχείρου.",
        "log_dns_title": "Εκκαθάριση cache DNS (ipconfig /flushdns)",
        "log_cleaning": "Εκκαθάριση: {label}",
        "log_path": "  Διαδρομή: {path}",
        "log_skipped_missing": "  Παραλείφθηκε (δεν βρέθηκε): {path}",
        "log_stats": "  Διαγράφηκαν: {deleted}, Ελευθερώθηκαν: {freed}, Παραλείφθηκαν: {errors}",
        "log_total_deleted": "ΣΥΝΟΛΟ διαγραμμένων στοιχείων: {count}",
        "log_total_freed": "ΣΥΝΟΛΟ ελεύθερου χώρου: {size}",
        "log_total_skipped_note": "Παραλείφθηκαν {count} στοιχεία (σε χρήση από το σύστημα — "
                                   "δοκιμάστε εκτέλεση ως Διαχειριστής).",
        "log_icon_cache_title": "Επαναφορά cache εικονιδίων (επανεκκίνηση Explorer)...",
        "log_explorer_title": "Επανεκκίνηση της διεργασίας explorer.exe...",
        "log_clipboard_title": "Εκκαθάριση προχείρου (EmptyClipboard)...",
        "log_user_skipped": "⏭ Παραλείφθηκε από τον χρήστη (το πλαίσιο δεν είναι επιλεγμένο): {label}",
        "log_extra_section_title": "Επιπλέον (launchers, επεξεργαστές):",
        "msg_clean_done_title": "Η Εκκαθάριση Ολοκληρώθηκε",
        "msg_clean_done_body": "Διαγραμμένα στοιχεία: {count}\nΕλεύθερος χώρος: {size}",
        "msg_done_title": "Ολοκληρώθηκε",
        "msg_failed_title": "Απέτυχε",
        "chk_section_hint": "Επιλέξτε τι θα καθαριστεί:",
        "cat_system": "Σύστημα",
        "cat_browsers": "Περιηγητές",
        "cat_apps": "Εφαρμογές",
        "chk_temp_files": "Προσωρινά Αρχεία",
        "chk_system_logs": "Αρχεία Καταγραφής Συστήματος",
        "chk_windows_update": "Cache Ενημερώσεων Windows",
        "chk_windows_old": "Φάκελος Windows.old",
        "desc_windows_old": "Περιέχει αρχεία από την προηγούμενη εγκατάσταση Windows.\n"
                             "Ασφαλής διαγραφή αν το σύστημα λειτουργεί κανονικά.\n"
                             "Ελευθερώνει 10–30 GB.",
        "chk_chrome": "Google Chrome",
        "chk_yandex": "Yandex Browser",
        "chk_opera_stable": "Opera Stable",
        "chk_opera_gx": "Opera GX",
        "chk_edge": "Microsoft Edge",
        "chk_firefox": "Mozilla Firefox",
        "chk_telegram": "Telegram",
        "chk_discord": "Discord",
        "chk_steam": "Steam",
        "chk_epic": "Epic Games",
        "chk_ea": "EA App",
        "chk_spotify": "Spotify",
        "chk_capcut": "CapCut",
        "chk_adobe_photoshop": "Adobe Photoshop",
        "chk_adobe_premiere": "Adobe Premiere",
        "chk_shaders": "Shaders Κάρτας Γραφικών (NVIDIA/AMD)",
    },
    "de": {
        "window_title": "ATCleaner — Windows-Temp-Dateien-Bereiniger",
        "main_title": "ATCleaner — Cache-Reiniger & Windows-Optimierer",
        "subtitle": "Wählen Sie unten aus, was bereinigt werden soll. Der DNS-Cache\n"
                     "wird bei jedem Lauf automatisch geleert.",
        "warning1": "⚠️ Die Ausführung als Administrator wird dringend empfohlen.",
        "warning2": "⚠️ Schließen Sie Telegram, Discord und Steam vor der Bereinigung.",
        "tab_clean": "Bereinigung",
        "tab_settings": "Einstellungen",
        "tab_tools": "Werkzeuge",
        "tab_about": "Über",
        "utilities_title": "Systemwerkzeuge",
        "about_title": "ATCleaner v1.5",
        "about_author": "Autor: ArseniosTs",
        "about_github_label": "🌐 GitHub-Repository",
        "about_disclaimer": "Dieses Tool verändert System- und Anwendungs-Cache-Dateien.\n"
                             "Nutzung auf eigenes Risiko. Vor größeren Änderungen wird ein\n"
                             "Backup wichtiger Daten empfohlen.",
        "clean_button": "🧹  Jetzt Bereinigen",
        "clean_button_working": "Bereinigung läuft...",
        "icon_cache_button": "🖼️  Symbol-Cache Wiederherstellen",
        "icon_cache_button_working": "Wird wiederhergestellt...",
        "explorer_button": "🗂️  Explorer Neu Starten",
        "explorer_button_working": "Neustart läuft...",
        "tool_desc_explorer": "Startet den Explorer neu, wenn die Taskleiste einfriert.",
        "tool_desc_icon_cache": "Behebt weiße oder falsch angezeigte Symbole.",
        "tool_desc_clipboard": "Löscht Zwischenablage-Daten sofort aus dem Speicher.",
        "lang_switch_hint": "Sofortiger Wechsel zwischen 5 Sprachen",
        "clipboard_button": "📋  Zwischenablage Leeren",
        "clipboard_button_working": "Wird geleert...",
        "footer": "Erstellt von ArseniosTs",
        "status_ready": "Bereit.",
        "status_cleaning": "Bereinigung läuft, bitte warten...",
        "status_done": "Fertig. {count} Objekte gelöscht, {size} freigegeben.",
        "status_icon_cache_working": "Symbol-Cache wird wiederhergestellt...",
        "status_icon_cache_done": "Symbol-Cache wiederhergestellt. Explorer neu gestartet.",
        "status_icon_cache_failed": "Symbol-Cache konnte nicht wiederhergestellt werden.",
        "status_explorer_working": "Explorer wird neu gestartet...",
        "status_explorer_done": "Explorer wurde neu gestartet.",
        "status_explorer_failed": "Explorer konnte nicht neu gestartet werden.",
        "status_clipboard_working": "Zwischenablage wird geleert...",
        "status_clipboard_done": "Zwischenablage wurde geleert.",
        "status_clipboard_failed": "Zwischenablage konnte nicht geleert werden.",
        "log_dns_title": "DNS-Cache wird geleert (ipconfig /flushdns)",
        "log_cleaning": "Bereinige: {label}",
        "log_path": "  Pfad: {path}",
        "log_skipped_missing": "  Übersprungen (nicht gefunden): {path}",
        "log_stats": "  Gelöscht: {deleted}, Freigegeben: {freed}, Übersprungen: {errors}",
        "log_total_deleted": "GESAMT gelöschte Objekte: {count}",
        "log_total_freed": "GESAMT freigegebener Speicher: {size}",
        "log_total_skipped_note": "{count} Objekte übersprungen (vom System verwendet — "
                                   "versuchen Sie es als Administrator).",
        "log_icon_cache_title": "Symbol-Cache wird wiederhergestellt (Explorer-Neustart)...",
        "log_explorer_title": "explorer.exe wird neu gestartet...",
        "log_clipboard_title": "Zwischenablage wird geleert (EmptyClipboard)...",
        "log_user_skipped": "⏭ Vom Nutzer übersprungen (Kontrollkästchen deaktiviert): {label}",
        "log_extra_section_title": "Zusätzlich (Launcher, Editoren):",
        "msg_clean_done_title": "Bereinigung Abgeschlossen",
        "msg_clean_done_body": "Gelöschte Objekte: {count}\nFreigegebener Speicher: {size}",
        "msg_done_title": "Fertig",
        "msg_failed_title": "Fehlgeschlagen",
        "chk_section_hint": "Wählen Sie, was bereinigt werden soll:",
        "cat_system": "System",
        "cat_browsers": "Browser",
        "cat_apps": "Programme",
        "chk_temp_files": "Temporäre Dateien",
        "chk_system_logs": "Systemprotokolle",
        "chk_windows_update": "Windows-Update-Cache",
        "chk_windows_old": "Windows.old-Ordner",
        "desc_windows_old": "Enthält Dateien der vorherigen Windows-Installation. Sicher zu\n"
                             "löschen, wenn das System einwandfrei läuft. Gibt 10–30 GB frei.",
        "chk_chrome": "Google Chrome",
        "chk_yandex": "Yandex Browser",
        "chk_opera_stable": "Opera Stable",
        "chk_opera_gx": "Opera GX",
        "chk_edge": "Microsoft Edge",
        "chk_firefox": "Mozilla Firefox",
        "chk_telegram": "Telegram",
        "chk_discord": "Discord",
        "chk_steam": "Steam",
        "chk_epic": "Epic Games",
        "chk_ea": "EA App",
        "chk_spotify": "Spotify",
        "chk_capcut": "CapCut",
        "chk_adobe_photoshop": "Adobe Photoshop",
        "chk_adobe_premiere": "Adobe Premiere",
        "chk_shaders": "GPU-Shader-Cache (NVIDIA/AMD)",
    },
    "es": {
        "window_title": "ATCleaner — Limpiador de Archivos Temporales de Windows",
        "main_title": "ATCleaner — Limpiador de Caché y Optimizador de Windows",
        "subtitle": "Elige qué limpiar a continuación. La caché DNS se vacía\n"
                     "automáticamente en cada limpieza.",
        "warning1": "⚠️ Se recomienda ejecutar como Administrador.",
        "warning2": "⚠️ Cierra Telegram, Discord y Steam antes de limpiar.",
        "tab_clean": "Limpieza",
        "tab_settings": "Ajustes",
        "tab_tools": "Herramientas",
        "tab_about": "Acerca de",
        "utilities_title": "Utilidades del Sistema",
        "about_title": "ATCleaner v1.5",
        "about_author": "Autor: ArseniosTs",
        "about_github_label": "🌐 Repositorio de GitHub",
        "about_disclaimer": "Esta herramienta modifica archivos de caché del sistema y\n"
                             "de aplicaciones. Úsala bajo tu propio riesgo. Se recomienda\n"
                             "tener una copia de seguridad antes de cambios importantes.",
        "clean_button": "🧹  Limpiar Ahora",
        "clean_button_working": "Limpiando...",
        "icon_cache_button": "🖼️  Reconstruir Caché de Iconos",
        "icon_cache_button_working": "Reconstruyendo...",
        "explorer_button": "🗂️  Reiniciar Explorador",
        "explorer_button_working": "Reiniciando...",
        "tool_desc_explorer": "Reinicia el Explorador si la barra de tareas se congela.",
        "tool_desc_icon_cache": "Corrige iconos en blanco o mostrados incorrectamente.",
        "tool_desc_clipboard": "Borra al instante los datos del portapapeles de la memoria.",
        "lang_switch_hint": "Cambio instantáneo entre 5 idiomas",
        "clipboard_button": "📋  Vaciar Portapapeles",
        "clipboard_button_working": "Vaciando...",
        "footer": "Creado por ArseniosTs",
        "status_ready": "Listo.",
        "status_cleaning": "Limpieza en curso, espera por favor...",
        "status_done": "Listo. Se eliminaron {count} elementos, se liberaron {size}.",
        "status_icon_cache_working": "Reconstruyendo caché de iconos...",
        "status_icon_cache_done": "Caché de iconos reconstruida. Explorador reiniciado.",
        "status_icon_cache_failed": "No se pudo reconstruir la caché de iconos.",
        "status_explorer_working": "Reiniciando el Explorador...",
        "status_explorer_done": "Explorador reiniciado.",
        "status_explorer_failed": "No se pudo reiniciar el Explorador.",
        "status_clipboard_working": "Vaciando el portapapeles...",
        "status_clipboard_done": "Portapapeles vaciado.",
        "status_clipboard_failed": "No se pudo vaciar el portapapeles.",
        "log_dns_title": "Vaciando caché DNS (ipconfig /flushdns)",
        "log_cleaning": "Limpiando: {label}",
        "log_path": "  Ruta: {path}",
        "log_skipped_missing": "  Omitido (no encontrado): {path}",
        "log_stats": "  Eliminado: {deleted}, Liberado: {freed}, Omitido: {errors}",
        "log_total_deleted": "TOTAL de elementos eliminados: {count}",
        "log_total_freed": "TOTAL de espacio liberado: {size}",
        "log_total_skipped_note": "Se omitieron {count} elementos (en uso por el sistema — "
                                   "intenta ejecutar como Administrador).",
        "log_icon_cache_title": "Reconstruyendo caché de iconos (reiniciando Explorador)...",
        "log_explorer_title": "Reiniciando el proceso explorer.exe...",
        "log_clipboard_title": "Vaciando el portapapeles (EmptyClipboard)...",
        "log_user_skipped": "⏭ Omitido por el usuario (casilla desmarcada): {label}",
        "log_extra_section_title": "Adicional (lanzadores, editores):",
        "msg_clean_done_title": "Limpieza Completa",
        "msg_clean_done_body": "Elementos eliminados: {count}\nEspacio liberado: {size}",
        "msg_done_title": "Listo",
        "msg_failed_title": "Fallido",
        "chk_section_hint": "Selecciona qué limpiar:",
        "cat_system": "Sistema",
        "cat_browsers": "Navegadores",
        "cat_apps": "Aplicaciones",
        "chk_temp_files": "Archivos Temporales",
        "chk_system_logs": "Registros del Sistema",
        "chk_windows_update": "Caché de Windows Update",
        "chk_windows_old": "Carpeta Windows.old",
        "desc_windows_old": "Contiene archivos de la instalación anterior de Windows. Se\n"
                             "puede eliminar de forma segura si el sistema funciona bien.\n"
                             "Libera 10–30 GB.",
        "chk_chrome": "Google Chrome",
        "chk_yandex": "Yandex Browser",
        "chk_opera_stable": "Opera Stable",
        "chk_opera_gx": "Opera GX",
        "chk_edge": "Microsoft Edge",
        "chk_firefox": "Mozilla Firefox",
        "chk_telegram": "Telegram",
        "chk_discord": "Discord",
        "chk_steam": "Steam",
        "chk_epic": "Epic Games",
        "chk_ea": "EA App",
        "chk_spotify": "Spotify",
        "chk_capcut": "CapCut",
        "chk_adobe_photoshop": "Adobe Photoshop",
        "chk_adobe_premiere": "Adobe Premiere",
        "chk_shaders": "Caché de Shaders GPU (NVIDIA/AMD)",
    },
}

# Названия папок очистки (для лога) на обоих языках, привязаны к
# внутренним идентификаторам, которые возвращает get_target_folders().
FOLDER_NAMES = {
    "ru": {
        "temp_user": "Temp пользователя",
        "temp_system": "Системный Temp",
        "windows_logs": "Логи Windows",
        "windows_update_cache": "Кэш обновлений Windows",
        "windows_old": "Папка Windows.old",
        "explorer_thumbcache": "Кэш миниатюр Explorer",
        "inetcache": "Кэш браузера (INetCache)",
        "discord_cache_local": "Кэш Discord (LocalAppData)",
        "discord_cache_roaming": "Кэш Discord (Roaming)",
        "chrome_cache": "Кэш Google Chrome",
        "yandex_cache": "Кэш Яндекс Браузера",
        "opera_cache": "Кэш Opera",
        "opera_gx_cache": "Кэш Opera GX",
        "edge_cache": "Кэш Microsoft Edge",
        "firefox_cache": "Кэш Mozilla Firefox",
        "epic_cache": "Кэш Epic Games Launcher",
        "ea_cache": "Кэш EA App",
        "spotify_cache": "Кэш Spotify",
        "capcut_cache": "Кэш CapCut",
        "adobe_photoshop_cache": "Кэш Adobe Photoshop (Camera Raw)",
        "nvidia_gl_cache": "Кэш шейдеров NVIDIA (OpenGL)",
        "nvidia_dx_cache": "Кэш шейдеров NVIDIA (DirectX)",
        "amd_dx_cache": "Кэш шейдеров AMD (DirectX)",
        "amd_dxc_cache": "Кэш шейдеров AMD (DXC)",
        "amd_gl_cache": "Кэш шейдеров AMD (OpenGL)",
        "telegram_cache": "Кэш Telegram",
        "telegram_media_cache": "Медиа-кэш Telegram",
        "adobe_media_cache": "Медиа-кэш Adobe",
        "adobe_media_cache_files": "Файлы медиа-кэша Adobe",
        "steam_appcache": "Кэш Steam (appcache)",
        "steam_htmlcache": "Кэш Steam (htmlcache)",
    },
    "en": {
        "temp_user": "User Temp",
        "temp_system": "System Temp",
        "windows_logs": "Windows Logs",
        "windows_update_cache": "Windows Update Cache",
        "windows_old": "Windows.old Folder",
        "explorer_thumbcache": "Explorer Thumbnail Cache",
        "inetcache": "Internet Cache (INetCache)",
        "discord_cache_local": "Discord Cache (LocalAppData)",
        "discord_cache_roaming": "Discord Cache (Roaming)",
        "chrome_cache": "Google Chrome Cache",
        "yandex_cache": "Yandex Browser Cache",
        "opera_cache": "Opera Cache",
        "opera_gx_cache": "Opera GX Cache",
        "edge_cache": "Microsoft Edge Cache",
        "firefox_cache": "Mozilla Firefox Cache",
        "epic_cache": "Epic Games Launcher Cache",
        "ea_cache": "EA App Cache",
        "spotify_cache": "Spotify Cache",
        "capcut_cache": "CapCut Cache",
        "adobe_photoshop_cache": "Adobe Photoshop Cache (Camera Raw)",
        "nvidia_gl_cache": "NVIDIA Shader Cache (OpenGL)",
        "nvidia_dx_cache": "NVIDIA Shader Cache (DirectX)",
        "amd_dx_cache": "AMD Shader Cache (DirectX)",
        "amd_dxc_cache": "AMD Shader Cache (DXC)",
        "amd_gl_cache": "AMD Shader Cache (OpenGL)",
        "telegram_cache": "Telegram Cache",
        "telegram_media_cache": "Telegram Media Cache",
        "adobe_media_cache": "Adobe Media Cache",
        "adobe_media_cache_files": "Adobe Media Cache Files",
        "steam_appcache": "Steam Cache (appcache)",
        "steam_htmlcache": "Steam Cache (htmlcache)",
    },
    "gr": {
        "temp_user": "Προσωρινά Αρχεία Χρήστη",
        "temp_system": "Προσωρινά Αρχεία Συστήματος",
        "windows_logs": "Αρχεία Καταγραφής Windows",
        "windows_update_cache": "Cache Ενημερώσεων Windows",
        "windows_old": "Φάκελος Windows.old",
        "explorer_thumbcache": "Cache Μικρογραφιών Explorer",
        "inetcache": "Cache Διαδικτύου (INetCache)",
        "discord_cache_local": "Cache Discord (LocalAppData)",
        "discord_cache_roaming": "Cache Discord (Roaming)",
        "chrome_cache": "Cache Google Chrome",
        "yandex_cache": "Cache Yandex Browser",
        "opera_cache": "Cache Opera",
        "opera_gx_cache": "Cache Opera GX",
        "edge_cache": "Cache Microsoft Edge",
        "firefox_cache": "Cache Mozilla Firefox",
        "epic_cache": "Cache Epic Games Launcher",
        "ea_cache": "Cache EA App",
        "spotify_cache": "Cache Spotify",
        "capcut_cache": "Cache CapCut",
        "adobe_photoshop_cache": "Cache Adobe Photoshop (Camera Raw)",
        "nvidia_gl_cache": "Cache Shader NVIDIA (OpenGL)",
        "nvidia_dx_cache": "Cache Shader NVIDIA (DirectX)",
        "amd_dx_cache": "Cache Shader AMD (DirectX)",
        "amd_dxc_cache": "Cache Shader AMD (DXC)",
        "amd_gl_cache": "Cache Shader AMD (OpenGL)",
        "telegram_cache": "Cache Telegram",
        "telegram_media_cache": "Cache Πολυμέσων Telegram",
        "adobe_media_cache": "Cache Πολυμέσων Adobe",
        "adobe_media_cache_files": "Αρχεία Cache Πολυμέσων Adobe",
        "steam_appcache": "Cache Steam (appcache)",
        "steam_htmlcache": "Cache Steam (htmlcache)",
    },
    "de": {
        "temp_user": "Benutzer-Temp",
        "temp_system": "System-Temp",
        "windows_logs": "Windows-Protokolle",
        "windows_update_cache": "Windows-Update-Cache",
        "windows_old": "Windows.old-Ordner",
        "explorer_thumbcache": "Explorer-Miniaturansicht-Cache",
        "inetcache": "Internet-Cache (INetCache)",
        "discord_cache_local": "Discord-Cache (LocalAppData)",
        "discord_cache_roaming": "Discord-Cache (Roaming)",
        "chrome_cache": "Google-Chrome-Cache",
        "yandex_cache": "Yandex-Browser-Cache",
        "opera_cache": "Opera-Cache",
        "opera_gx_cache": "Opera-GX-Cache",
        "edge_cache": "Microsoft-Edge-Cache",
        "firefox_cache": "Mozilla-Firefox-Cache",
        "epic_cache": "Epic-Games-Launcher-Cache",
        "ea_cache": "EA-App-Cache",
        "spotify_cache": "Spotify-Cache",
        "capcut_cache": "CapCut-Cache",
        "adobe_photoshop_cache": "Adobe-Photoshop-Cache (Camera Raw)",
        "nvidia_gl_cache": "NVIDIA-Shader-Cache (OpenGL)",
        "nvidia_dx_cache": "NVIDIA-Shader-Cache (DirectX)",
        "amd_dx_cache": "AMD-Shader-Cache (DirectX)",
        "amd_dxc_cache": "AMD-Shader-Cache (DXC)",
        "amd_gl_cache": "AMD-Shader-Cache (OpenGL)",
        "telegram_cache": "Telegram-Cache",
        "telegram_media_cache": "Telegram-Medien-Cache",
        "adobe_media_cache": "Adobe-Medien-Cache",
        "adobe_media_cache_files": "Adobe-Medien-Cache-Dateien",
        "steam_appcache": "Steam-Cache (appcache)",
        "steam_htmlcache": "Steam-Cache (htmlcache)",
    },
    "es": {
        "temp_user": "Temp del Usuario",
        "temp_system": "Temp del Sistema",
        "windows_logs": "Registros de Windows",
        "windows_update_cache": "Caché de Windows Update",
        "windows_old": "Carpeta Windows.old",
        "explorer_thumbcache": "Caché de Miniaturas del Explorador",
        "inetcache": "Caché de Internet (INetCache)",
        "discord_cache_local": "Caché de Discord (LocalAppData)",
        "discord_cache_roaming": "Caché de Discord (Roaming)",
        "chrome_cache": "Caché de Google Chrome",
        "yandex_cache": "Caché de Yandex Browser",
        "opera_cache": "Caché de Opera",
        "opera_gx_cache": "Caché de Opera GX",
        "edge_cache": "Caché de Microsoft Edge",
        "firefox_cache": "Caché de Mozilla Firefox",
        "epic_cache": "Caché de Epic Games Launcher",
        "ea_cache": "Caché de EA App",
        "spotify_cache": "Caché de Spotify",
        "capcut_cache": "Caché de CapCut",
        "adobe_photoshop_cache": "Caché de Adobe Photoshop (Camera Raw)",
        "nvidia_gl_cache": "Caché de Shaders NVIDIA (OpenGL)",
        "nvidia_dx_cache": "Caché de Shaders NVIDIA (DirectX)",
        "amd_dx_cache": "Caché de Shaders AMD (DirectX)",
        "amd_dxc_cache": "Caché de Shaders AMD (DXC)",
        "amd_gl_cache": "Caché de Shaders AMD (OpenGL)",
        "telegram_cache": "Caché de Telegram",
        "telegram_media_cache": "Caché Multimedia de Telegram",
        "adobe_media_cache": "Caché Multimedia de Adobe",
        "adobe_media_cache_files": "Archivos de Caché Multimedia de Adobe",
        "steam_appcache": "Caché de Steam (appcache)",
        "steam_htmlcache": "Caché de Steam (htmlcache)",
    },
}

DEFAULT_LANG = "en"

# ---------------------------------------------------------------------------
# Определение флажков очистки: категория, ключ перевода, связанные папки
# ---------------------------------------------------------------------------
# Порядок категорий на экране
CATEGORY_ORDER = ["system", "browsers", "apps"]
CATEGORY_LABEL_KEYS = {
    "system": "cat_system",
    "browsers": "cat_browsers",
    "apps": "cat_apps",
}

# Каждый флажок: id -> (категория, ключ_перевода_названия, [id папок],
#                         ключ_перевода_описания_или_None)
# Теперь у каждого приложения/источника — своя отдельная галочка, никаких
# объединений (Chrome/Yandex, Opera/Opera GX и т.д. больше не смешиваются).
CHECKBOX_DEFS = [
    # --- Система ---
    ("temp_files", "system", "chk_temp_files",
     ["temp_user", "temp_system", "explorer_thumbcache", "inetcache"], None),
    ("system_logs", "system", "chk_system_logs", ["windows_logs"], None),
    ("windows_update", "system", "chk_windows_update", ["windows_update_cache"], None),
    ("windows_old", "system", "chk_windows_old", ["windows_old"], "desc_windows_old"),

    # --- Браузеры (каждый отдельно) ---
    ("chrome", "browsers", "chk_chrome", ["chrome_cache"], None),
    ("yandex", "browsers", "chk_yandex", ["yandex_cache"], None),
    ("opera_stable", "browsers", "chk_opera_stable", ["opera_cache"], None),
    ("opera_gx", "browsers", "chk_opera_gx", ["opera_gx_cache"], None),
    ("edge", "browsers", "chk_edge", ["edge_cache"], None),
    ("firefox", "browsers", "chk_firefox", ["firefox_cache"], None),

    # --- Программы (каждая отдельно) ---
    ("telegram", "apps", "chk_telegram", ["telegram_cache", "telegram_media_cache"], None),
    ("discord", "apps", "chk_discord", ["discord_cache_local", "discord_cache_roaming"], None),
    ("steam", "apps", "chk_steam", ["steam_appcache", "steam_htmlcache"], None),
    ("epic", "apps", "chk_epic", ["epic_cache"], None),
    ("ea", "apps", "chk_ea", ["ea_cache"], None),
    ("spotify", "apps", "chk_spotify", ["spotify_cache"], None),
    ("capcut", "apps", "chk_capcut", ["capcut_cache"], None),
    ("adobe_photoshop", "apps", "chk_adobe_photoshop", ["adobe_photoshop_cache"], None),
    ("adobe_premiere", "apps", "chk_adobe_premiere",
     ["adobe_media_cache", "adobe_media_cache_files"], None),
    ("shaders", "apps", "chk_shaders",
     ["nvidia_gl_cache", "nvidia_dx_cache", "amd_dx_cache", "amd_dxc_cache", "amd_gl_cache"], None),
]


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


def get_firefox_cache_dirs(local_appdata):
    """Firefox хранит кэш внутри папки профиля со случайным именем
    (например 'abcd1234.default-release\\cache2'), поэтому путь нельзя
    прописать напрямую — сканируем папку профилей и собираем все cache2.
    Если Firefox не установлен, возвращает пустой список (без ошибок).
    """
    profiles_root = os.path.join(local_appdata, "Mozilla", "Firefox", "Profiles")
    cache_dirs = []
    if os.path.isdir(profiles_root):
        try:
            for entry in os.scandir(profiles_root):
                if entry.is_dir():
                    cache_dir = os.path.join(entry.path, "cache2")
                    cache_dirs.append(cache_dir)
        except OSError:
            pass
    return cache_dirs


# ---------------------------------------------------------------------------
# Список папок для очистки: folder_id -> путь (строка) ИЛИ список путей
# (например, "firefox_cache" — сразу несколько профилей)
# ---------------------------------------------------------------------------
def get_target_folders():
    """Возвращает словарь {folder_id: путь} для ВСЕХ известных источников."""
    folders = {}

    user_temp = tempfile.gettempdir()
    if user_temp:
        folders["temp_user"] = user_temp

    windir = os.environ.get("WINDIR", r"C:\Windows")
    system_drive = os.environ.get("SystemDrive", "C:")

    folders["temp_system"] = os.path.join(windir, "Temp")
    folders["windows_logs"] = os.path.join(windir, "Logs")
    folders["windows_update_cache"] = os.path.join(windir, "SoftwareDistribution", "Download")
    # C:\Windows.old — файлы предыдущей установки Windows
    folders["windows_old"] = system_drive + os.sep + "Windows.old"

    local_appdata = os.environ.get("LOCALAPPDATA")
    roaming_appdata = os.environ.get("APPDATA")
    program_files_x86 = os.environ.get("ProgramFiles(x86)") or os.environ.get("ProgramFiles")

    if local_appdata:
        folders["explorer_thumbcache"] = os.path.join(local_appdata, "Microsoft", "Windows", "Explorer")
        folders["inetcache"] = os.path.join(local_appdata, "Microsoft", "Windows", "INetCache")

        # Discord: чистим ТОЛЬКО "Cache". "Code Cache" никогда не трогаем —
        # его удаление ломает запуск Discord.
        folders["discord_cache_local"] = os.path.join(local_appdata, "discord", "Cache")

        # Браузеры
        folders["chrome_cache"] = os.path.join(
            local_appdata, "Google", "Chrome", "User Data", "Default", "Cache")
        folders["yandex_cache"] = os.path.join(
            local_appdata, "Yandex", "YandexBrowser", "User Data", "Default", "Cache")
        folders["opera_cache"] = os.path.join(local_appdata, "Opera Software", "Opera Stable", "Cache")
        folders["opera_gx_cache"] = os.path.join(local_appdata, "Opera Software", "Opera GX Stable", "Cache")
        folders["edge_cache"] = os.path.join(
            local_appdata, "Microsoft", "Edge", "User Data", "Default", "Cache")

        # Firefox: несколько профилей возможны — возвращаем список путей
        folders["firefox_cache"] = get_firefox_cache_dirs(local_appdata)

        # Игровые лаунчеры
        folders["epic_cache"] = os.path.join(local_appdata, "EpicGamesLauncher", "Saved", "webcache")
        folders["ea_cache"] = os.path.join(local_appdata, "Electronic Arts", "EA Desktop", "CacheStorage")

        # Spotify (аудио-кэш; путь может отличаться для версии из Microsoft Store)
        folders["spotify_cache"] = os.path.join(local_appdata, "Spotify", "Storage")

        # Видеоредакторы / графика
        folders["capcut_cache"] = os.path.join(local_appdata, "CapCut", "User Data", "Cache")
        # Camera Raw — кэш, который использует Adobe Photoshop
        folders["adobe_photoshop_cache"] = os.path.join(local_appdata, "Adobe", "CameraRaw", "Cache")

        # Кэш шейдеров видеокарт
        folders["nvidia_gl_cache"] = os.path.join(local_appdata, "NVIDIA", "GLCache")
        folders["nvidia_dx_cache"] = os.path.join(local_appdata, "NVIDIA", "DXCache")
        folders["amd_dx_cache"] = os.path.join(local_appdata, "AMD", "DxCache")
        folders["amd_dxc_cache"] = os.path.join(local_appdata, "AMD", "DxcCache")
        folders["amd_gl_cache"] = os.path.join(local_appdata, "AMD", "GLCache")

    if roaming_appdata:
        folders["discord_cache_roaming"] = os.path.join(roaming_appdata, "discord", "Cache")

        folders["telegram_cache"] = os.path.join(
            roaming_appdata, "Telegram Desktop", "tdata", "user_data", "cache")
        folders["telegram_media_cache"] = os.path.join(
            roaming_appdata, "Telegram Desktop", "tdata", "user_data", "media_cache")

        folders["adobe_media_cache"] = os.path.join(roaming_appdata, "Adobe", "Common", "Media Cache")
        folders["adobe_media_cache_files"] = os.path.join(
            roaming_appdata, "Adobe", "Common", "Media Cache Files")

    if program_files_x86:
        folders["steam_appcache"] = os.path.join(program_files_x86, "Steam", "appcache")
        folders["steam_htmlcache"] = os.path.join(program_files_x86, "Steam", "htmlcache")

    return folders


def format_size(num_bytes, lang="ru"):
    units_ru = ["Б", "КБ", "МБ", "ГБ", "ТБ"]
    units_en = ["B", "KB", "MB", "GB", "TB"]
    units = units_ru if lang == "ru" else units_en
    value = float(num_bytes)
    for unit in units[:-1]:
        if value < 1024:
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} {units[-1]}"


def clean_folder(path, log_callback, tr):
    """Удаляет содержимое папки (не саму папку), считает освобождённый объём.

    Безопасный метод: если папки не существует — пропуск без ошибки; если
    файл/подпапка защищены или заняты другим процессом (в т.ч. системные
    файлы внутри Windows.old) — пропускаем именно этот объект и продолжаем,
    программа никогда не падает из-за отказа в доступе.
    """
    freed = 0
    deleted_count = 0
    error_count = 0

    if not os.path.isdir(path):
        log_callback(tr("log_skipped_missing", path=path))
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
                shutil.rmtree(entry_path, onerror=_ignore_rmtree_errors)
                # После rmtree с обработчиком ошибок папка может частично
                # остаться — пересчитываем реально удалённый объём.
                remaining = get_dir_size(entry_path) if os.path.isdir(entry_path) else 0
                freed += max(size - remaining, 0)
                if os.path.isdir(entry_path) and remaining > 0:
                    error_count += 1
                else:
                    deleted_count += 1
        except (PermissionError, OSError):
            error_count += 1

    return freed, deleted_count, error_count


def _ignore_rmtree_errors(func, path, exc_info):
    """Обработчик ошибок для shutil.rmtree — просто пропускаем защищённые
    или занятые файлы (например, внутри Windows.old), не роняя программу."""
    pass


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
    """Выполняет ipconfig /flushdns. Возвращает (успех, текст вывода команды)."""
    if not is_windows():
        return False, "Windows only."
    try:
        result = subprocess.run(
            ["ipconfig", "/flushdns"],
            capture_output=True,
            text=True,
            shell=False,
        )
        output = (result.stdout or result.stderr or "").strip()
        return (result.returncode == 0), output
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# Инструменты оптимизации
# ---------------------------------------------------------------------------
def rebuild_icon_cache():
    """
    Полностью пересобирает кэш иконок Windows — решает проблему, когда
    значки файлов/ярлыков становятся белыми, пустыми или долго загружаются.

    Порядок действий (безопасный метод):
      1. Принудительно завершаем explorer.exe — пока он запущен, файлы
         кэша иконок заблокированы и не могут быть удалены.
      2. Удаляем старый общий файл %LOCALAPPDATA%\\IconCache.db (актуально
         для более старых версий Windows).
      3. Удаляем все файлы %LOCALAPPDATA%\\Microsoft\\Windows\\Explorer\\
         iconcache_*.db (актуально для современных версий Windows —
         кэш иконок хранится не одним файлом, а несколькими).
      4. Запускаем explorer.exe заново — Windows автоматически создаёт
         новый чистый кэш иконок при следующем обращении к файлам.

    Ошибки доступа к отдельным файлам не прерывают процесс — такие файлы
    просто пропускаются.

    Возвращает (успех: bool, техническое сообщение: str или None).
    """
    if not is_windows():
        return False, "windows_only"

    try:
        # 1. Останавливаем проводник, чтобы разблокировать файлы кэша
        subprocess.run(
            ["taskkill", "/F", "/IM", "explorer.exe"],
            capture_output=True,
            text=True,
        )
        time.sleep(1.0)

        local_appdata = os.environ.get("LOCALAPPDATA")
        if local_appdata:
            # 2. Старый единый файл кэша иконок
            legacy_path = os.path.join(local_appdata, "IconCache.db")
            if os.path.isfile(legacy_path):
                try:
                    os.remove(legacy_path)
                except OSError:
                    pass  # занят/защищён — пропускаем, не критично

            # 3. Современные файлы iconcache_*.db
            explorer_dir = os.path.join(local_appdata, "Microsoft", "Windows", "Explorer")
            if os.path.isdir(explorer_dir):
                for entry in os.scandir(explorer_dir):
                    name_lower = entry.name.lower()
                    if entry.is_file() and name_lower.startswith("iconcache_") and name_lower.endswith(".db"):
                        try:
                            os.remove(entry.path)
                        except OSError:
                            pass

        # 4. Запускаем проводник заново — кэш иконок пересоздастся сам
        time.sleep(0.5)
        subprocess.Popen("explorer.exe")
        return True, None

    except Exception as e:
        # Даже если что-то пошло не так, стараемся вернуть рабочий стол
        try:
            subprocess.Popen("explorer.exe")
        except Exception:
            pass
        return False, str(e)


def restart_explorer():
    """Принудительно перезапускает explorer.exe. Возвращает (успех, тех. сообщение)."""
    if not is_windows():
        return False, "windows_only"
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "explorer.exe"],
            capture_output=True,
            text=True,
        )
        time.sleep(1.5)
        subprocess.Popen("explorer.exe")
        return True, None
    except Exception as e:
        return False, str(e)


def clear_clipboard():
    """
    Очищает системный буфер обмена Windows через стандартный вызов Win32 API
    (OpenClipboard → EmptyClipboard → CloseClipboard). Безопасно: если буфер
    занят другим приложением, попытка просто завершается неудачей, ничего
    не ломая.

    Возвращает (успех: bool, техническое сообщение: str или None).
    """
    if not is_windows():
        return False, "windows_only"
    try:
        user32 = ctypes.windll.user32

        if not user32.OpenClipboard(None):
            return False, "open_clipboard_failed"

        try:
            if not user32.EmptyClipboard():
                return False, "empty_clipboard_failed"
        finally:
            user32.CloseClipboard()

        return True, None
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# Интерфейс
# ---------------------------------------------------------------------------
class CleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")

        # Сначала пробуем восстановить язык, сохранённый в config.txt при
        # прошлом запуске. Если файла нет или он повреждён — English.
        self.current_lang = load_saved_language() or DEFAULT_LANG

        self.title(self.tr("window_title"))

        # Принудительно закрепляем иконку сразу после создания окна.
        try:
            self.iconbitmap("ATCleanerLogoMR.ico")
        except Exception:
            pass

        # Окно подобрано так, чтобы вкладка "Очистка" (лог + кнопка),
        # вкладка "Настройки" (список флажков) и вкладка "Инструменты"
        # помещались без обрезки текста.
        self.geometry("540x520")
        self.resizable(False, False)
        self.configure(fg_color=COLOR_BG)

        self._set_app_icon()
        self.after(200, self._set_app_icon)

        # Хранилища ссылок на виджеты флажков/категорий/описаний — нужны
        # для мгновенного перевода при смене языка.
        self.checkbox_vars = {}
        self.checkbox_widgets = {}
        self.category_label_widgets = {}
        self.description_widgets = {}

        # --- Заголовок -------------------------------------------------
        self.title_label = ctk.CTkLabel(
            self,
            text=self.tr("main_title"),
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=COLOR_TEXT,
        )
        self.title_label.pack(pady=(18, 4))

        self.subtitle_label = ctk.CTkLabel(
            self,
            text=self.tr("subtitle"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLOR_SUBTEXT,
            justify="center",
        )
        self.subtitle_label.pack(pady=(0, 8))

        # --- Предупреждения ------------------------------------------------
        self.warning_label = ctk.CTkLabel(
            self,
            text=self._warning_text(),
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=COLOR_WARNING,
            justify="center",
        )
        self.warning_label.pack(pady=(0, 10))

        # --- Вкладки ------------------------------------------------------
        self.tabview = ctk.CTkTabview(
            self,
            width=500,
            height=260,
            fg_color=COLOR_TAB_BG,
            segmented_button_fg_color=COLOR_TAB_BG,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_TAB_BG,
            text_color=COLOR_TEXT,
            corner_radius=10,
        )
        self.tabview.pack(padx=16, pady=(0, 12), fill="x")

        tab_clean = self.tabview.add(self.tr("tab_clean"))
        tab_settings = self.tabview.add(self.tr("tab_settings"))
        tab_tools = self.tabview.add(self.tr("tab_tools"))
        tab_about = self.tabview.add(self.tr("tab_about"))
        tab_clean.configure(fg_color=COLOR_TAB_BG)
        tab_settings.configure(fg_color=COLOR_TAB_BG)
        tab_tools.configure(fg_color=COLOR_TAB_BG)
        tab_about.configure(fg_color=COLOR_TAB_BG)
        self._tab_clean_frame = tab_clean
        self._tab_settings_frame = tab_settings
        self._tab_tools_frame = tab_tools
        self._tab_about_frame = tab_about

        self._build_clean_tab(tab_clean)
        self._build_settings_tab(tab_settings)
        self._build_tools_tab(tab_tools)
        self._build_about_tab(tab_about)

        # --- Статус ------------------------------------------------------
        self.status_label = ctk.CTkLabel(
            self,
            text=self.tr("status_ready"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLOR_TEXT,
        )
        self.status_label.pack(pady=(0, 8))

        # --- Нижняя панель: переключатель языка (слева) + подпись автора (справа) ---
        bottom_bar = ctk.CTkFrame(self, fg_color=COLOR_BG)
        bottom_bar.pack(fill="x", side="bottom", padx=14, pady=(0, 12))

        self.language_menu = ctk.CTkOptionMenu(
            bottom_bar,
            values=LANG_OPTION_LIST,
            command=self.on_language_change,
            width=150,
            height=26,
            font=ctk.CTkFont(family="Segoe UI", size=10),
            dropdown_font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color=COLOR_TOOL_ACCENT,
            button_color=COLOR_TOOL_ACCENT_HOVER,
            button_hover_color=COLOR_ACCENT,
            dropdown_fg_color=COLOR_TAB_BG,
            dropdown_hover_color=COLOR_ACCENT,
            text_color="white",
        )
        default_option = next(
            (option for option, code in LANG_OPTIONS.items() if code == self.current_lang),
            LANG_OPTION_LIST[0],
        )
        self.language_menu.set(default_option)
        self.language_menu.pack(side="left")

        self.lang_switch_hint_label = ctk.CTkLabel(
            bottom_bar,
            text=self.tr("lang_switch_hint"),
            font=ctk.CTkFont(family="Segoe UI", size=8),
            text_color=COLOR_SUBTEXT,
        )
        self.lang_switch_hint_label.pack(side="left", padx=(8, 0))

        self.footer_label = ctk.CTkLabel(
            bottom_bar,
            text=self.tr("footer"),
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=COLOR_FOOTER,
        )
        self.footer_label.pack(side="right")

        # Кликабельная ссылка на репозиторий GitHub — стоит сразу слева от
        # подписи автора, в правой части нижней панели.
        self.lbl_github = ctk.CTkLabel(
            bottom_bar,
            text="💻 GitHub",
            font=ctk.CTkFont(family="Segoe UI", size=10, underline=True),
            text_color="#FFFFFF",
            cursor="hand2",
        )
        self.lbl_github.pack(side="right", padx=(0, 10))
        self.lbl_github.bind("<Button-1>", self.open_github)

    # -----------------------------------------------------------------
    # Построение вкладки "Очистка": только лог + кнопка запуска
    # -----------------------------------------------------------------
    def _build_clean_tab(self, parent):
        self.log_area = ctk.CTkTextbox(
            parent,
            width=460,
            height=110,
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
        self.log_area.pack(padx=16, pady=(16, 14))

        self.clean_button = ctk.CTkButton(
            parent,
            text=self.tr("clean_button"),
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="white",
            corner_radius=10,
            width=260,
            height=46,
            command=self.start_cleaning,
        )
        self.clean_button.pack(pady=(0, 12))

    # -----------------------------------------------------------------
    # Построение вкладки "Настройки": все флажки очистки одним списком
    # -----------------------------------------------------------------
    def _build_settings_tab(self, parent):
        self.chk_hint_label = ctk.CTkLabel(
            parent,
            text=self.tr("chk_section_hint"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COLOR_SUBTEXT,
        )
        self.chk_hint_label.pack(pady=(10, 4), anchor="w", padx=18)

        # Фирменная фиолетовая обводка в цвет приложения
        scroll_frame = ctk.CTkScrollableFrame(
            parent,
            width=460,
            height=150,
            fg_color=COLOR_CARD_BG,
            corner_radius=10,
            border_width=1,
            border_color="#8A2BE2",
            scrollbar_button_color=COLOR_LOG_BORDER,
            scrollbar_button_hover_color=COLOR_ACCENT,
        )
        scroll_frame.pack(padx=16, pady=(0, 14), fill="x")

        current_category = None
        for checkbox_id, category, label_key, folder_ids, desc_key in CHECKBOX_DEFS:
            if category != current_category:
                current_category = category
                cat_label = ctk.CTkLabel(
                    scroll_frame,
                    text=self.tr(CATEGORY_LABEL_KEYS[category]),
                    font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                    text_color=COLOR_CATEGORY,
                )
                cat_label.pack(anchor="w", padx=12, pady=(12, 2))
                self.category_label_widgets[category] = cat_label

            var = ctk.BooleanVar(value=True)  # по умолчанию все включены
            checkbox = ctk.CTkCheckBox(
                scroll_frame,
                text=self.tr(label_key),
                variable=var,
                onvalue=True,
                offvalue=False,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=COLOR_TEXT,
                fg_color=COLOR_ACCENT,
                hover_color=COLOR_ACCENT_HOVER,
                border_color=COLOR_LOG_BORDER,
                checkmark_color="white",
            )
            checkbox.pack(anchor="w", padx=24, pady=(4, 0))

            self.checkbox_vars[checkbox_id] = var
            self.checkbox_widgets[checkbox_id] = checkbox

            if desc_key:
                desc_label = ctk.CTkLabel(
                    scroll_frame,
                    text=self.tr(desc_key),
                    font=ctk.CTkFont(family="Segoe UI", size=10),
                    text_color=COLOR_SUBTEXT,
                    justify="left",
                )
                desc_label.pack(anchor="w", padx=46, pady=(0, 6))
                self.description_widgets[checkbox_id] = desc_label

        # Небольшой отступ снизу списка
        ctk.CTkLabel(scroll_frame, text="", height=1).pack(pady=(0, 4))

    # -----------------------------------------------------------------
    # Построение вкладки "Инструменты": заголовок + сетка 2x2 из 4 кнопок
    # -----------------------------------------------------------------
    def _build_tools_tab(self, parent):
        self.utilities_title_label = ctk.CTkLabel(
            parent,
            text=self.tr("utilities_title"),
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=COLOR_TEXT,
        )
        self.utilities_title_label.pack(pady=(10, 10))

        tools_wrapper = ctk.CTkFrame(parent, fg_color=COLOR_TAB_BG)
        tools_wrapper.pack(expand=True)

        button_kwargs = dict(
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=COLOR_TOOL_ACCENT,
            hover_color=COLOR_TOOL_ACCENT_HOVER,
            text_color="white",
            corner_radius=10,
            width=280,
            height=36,
        )
        desc_font = ctk.CTkFont(family="Segoe UI", size=9)

        self.explorer_button = ctk.CTkButton(
            tools_wrapper,
            text=self.tr("explorer_button"),
            command=self.start_explorer_restart,
            **button_kwargs,
        )
        self.explorer_button.pack(pady=(4, 0))
        self.explorer_desc_label = ctk.CTkLabel(
            tools_wrapper, text=self.tr("tool_desc_explorer"),
            font=desc_font, text_color=COLOR_SUBTEXT,
        )
        self.explorer_desc_label.pack(pady=(2, 10))

        self.icon_cache_button = ctk.CTkButton(
            tools_wrapper,
            text=self.tr("icon_cache_button"),
            command=self.start_icon_cache_rebuild,
            **button_kwargs,
        )
        self.icon_cache_button.pack(pady=(4, 0))
        self.icon_cache_desc_label = ctk.CTkLabel(
            tools_wrapper, text=self.tr("tool_desc_icon_cache"),
            font=desc_font, text_color=COLOR_SUBTEXT,
        )
        self.icon_cache_desc_label.pack(pady=(2, 10))

        self.clipboard_button = ctk.CTkButton(
            tools_wrapper,
            text=self.tr("clipboard_button"),
            command=self.start_clipboard_clear,
            **button_kwargs,
        )
        self.clipboard_button.pack(pady=(4, 0))
        self.clipboard_desc_label = ctk.CTkLabel(
            tools_wrapper, text=self.tr("tool_desc_clipboard"),
            font=desc_font, text_color=COLOR_SUBTEXT,
        )
        self.clipboard_desc_label.pack(pady=(2, 4))

    # -----------------------------------------------------------------
    # Построение вкладки "О программе": название, автор, GitHub, дисклеймер
    # -----------------------------------------------------------------
    def _build_about_tab(self, parent):
        wrapper = ctk.CTkFrame(parent, fg_color=COLOR_TAB_BG)
        wrapper.pack(expand=True, fill="both", pady=(18, 10))

        self.about_title_label = ctk.CTkLabel(
            wrapper,
            text=self.tr("about_title"),
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=COLOR_TEXT,
        )
        self.about_title_label.pack(pady=(4, 6))

        self.about_author_label = ctk.CTkLabel(
            wrapper,
            text=self.tr("about_author"),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COLOR_SUBTEXT,
        )
        self.about_author_label.pack(pady=(0, 10))

        # Интерактивная ссылка на GitHub — тот же принцип, что и в нижней
        # панели: подчёркнутый текст, курсор "рука", клик открывает браузер.
        self.about_github_label = ctk.CTkLabel(
            wrapper,
            text=self.tr("about_github_label"),
            font=ctk.CTkFont(family="Segoe UI", size=13, underline=True, weight="bold"),
            text_color=COLOR_ACCENT,
            cursor="hand2",
        )
        self.about_github_label.pack(pady=(0, 16))
        self.about_github_label.bind("<Button-1>", self.open_github)

        self.about_disclaimer_label = ctk.CTkLabel(
            wrapper,
            text=self.tr("about_disclaimer"),
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=COLOR_SUBTEXT,
            justify="center",
        )
        self.about_disclaimer_label.pack(pady=(0, 4), padx=20)

    # -----------------------------------------------------------------
    # Локализация
    # -----------------------------------------------------------------
    def tr(self, key, **kwargs):
        """Возвращает переведённую строку текущего языка с подстановкой параметров."""
        text = TR[self.current_lang].get(key, key)
        if kwargs:
            text = text.format(**kwargs)
        return text

    def folder_name(self, folder_id):
        return FOLDER_NAMES[self.current_lang].get(folder_id, folder_id)

    def _warning_text(self):
        return f"{self.tr('warning1')}\n{self.tr('warning2')}"

    def on_language_change(self, choice):
        new_lang = LANG_OPTIONS.get(choice, DEFAULT_LANG)
        if new_lang == self.current_lang:
            return

        old_tab_clean = self.tr("tab_clean")
        old_tab_settings = self.tr("tab_settings")
        old_tab_tools = self.tr("tab_tools")
        old_tab_about = self.tr("tab_about")

        self.current_lang = new_lang
        save_language(new_lang)  # запоминаем выбор для следующего запуска
        self.apply_language(old_tab_clean, old_tab_settings, old_tab_tools, old_tab_about)

    def apply_language(self, old_tab_clean, old_tab_settings, old_tab_tools, old_tab_about):
        """Переводит абсолютно все статичные элементы интерфейса, включая
        флажки очистки, вкладку "Инструменты" и вкладку "О программе"."""
        self.title(self.tr("window_title"))
        self.title_label.configure(text=self.tr("main_title"))
        self.subtitle_label.configure(text=self.tr("subtitle"))
        self.warning_label.configure(text=self._warning_text())
        self.footer_label.configure(text=self.tr("footer"))
        self.chk_hint_label.configure(text=self.tr("chk_section_hint"))
        self.utilities_title_label.configure(text=self.tr("utilities_title"))
        self.lang_switch_hint_label.configure(text=self.tr("lang_switch_hint"))

        # Короткие подсказки под кнопками вкладки "Инструменты"
        self.explorer_desc_label.configure(text=self.tr("tool_desc_explorer"))
        self.icon_cache_desc_label.configure(text=self.tr("tool_desc_icon_cache"))
        self.clipboard_desc_label.configure(text=self.tr("tool_desc_clipboard"))

        # Вкладка "О программе"
        self.about_title_label.configure(text=self.tr("about_title"))
        self.about_author_label.configure(text=self.tr("about_author"))
        self.about_github_label.configure(text=self.tr("about_github_label"))
        self.about_disclaimer_label.configure(text=self.tr("about_disclaimer"))

        # Переименовываем вкладки, сохраняя их содержимое
        self.tabview.rename(old_tab_clean, self.tr("tab_clean"))
        self.tabview.rename(old_tab_settings, self.tr("tab_settings"))
        self.tabview.rename(old_tab_tools, self.tr("tab_tools"))
        self.tabview.rename(old_tab_about, self.tr("tab_about"))

        # Заголовки категорий флажков
        for category, widget in self.category_label_widgets.items():
            widget.configure(text=self.tr(CATEGORY_LABEL_KEYS[category]))

        # Тексты флажков
        for checkbox_id, category, label_key, folder_ids, desc_key in CHECKBOX_DEFS:
            widget = self.checkbox_widgets.get(checkbox_id)
            if widget is not None:
                widget.configure(text=self.tr(label_key))
            if desc_key and checkbox_id in self.description_widgets:
                self.description_widgets[checkbox_id].configure(text=self.tr(desc_key))

        # Кнопки возвращаем в состояние "по умолчанию" при смене языка —
        # если операция не выполняется в данный момент, это безопасно.
        if self.clean_button.cget("state") != "disabled":
            self.clean_button.configure(text=self.tr("clean_button"))
        if self.icon_cache_button.cget("state") != "disabled":
            self.icon_cache_button.configure(text=self.tr("icon_cache_button"))
        if self.explorer_button.cget("state") != "disabled":
            self.explorer_button.configure(text=self.tr("explorer_button"))
        if self.clipboard_button.cget("state") != "disabled":
            self.clipboard_button.configure(text=self.tr("clipboard_button"))

        # Статус-строку переводим только в состоянии "готово"
        self.status_label.configure(text=self.tr("status_ready"))

    # -----------------------------------------------------------------
    # Общие утилиты интерфейса
    # -----------------------------------------------------------------
    def _set_app_icon(self):
        if is_windows() and os.path.isfile(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass

    def open_github(self, event=None):
        """Открывает репозиторий проекта на GitHub в браузере по умолчанию."""
        webbrowser.open_new("https://github.com/ArseniosTs/AT-Cleaner")

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
        self.icon_cache_button.configure(state=state)
        self.explorer_button.configure(state=state)
        self.clipboard_button.configure(state=state)

    # -----------------------------------------------------------------
    # Очистка мусора
    # -----------------------------------------------------------------
    def start_cleaning(self):
        self.set_buttons_state("disabled")
        self.clean_button.configure(text=self.tr("clean_button_working"))
        self.status_label.configure(text=self.tr("status_cleaning"))
        self.clear_log()

        # Снимок состояния флажков делаем в основном потоке (GUI-объекты
        # нельзя безопасно опрашивать из фонового потока в некоторых
        # реализациях tkinter), а сама очистка идёт в отдельном потоке.
        checkbox_state = {cid: var.get() for cid, var in self.checkbox_vars.items()}

        thread = threading.Thread(target=self.run_cleaning, args=(checkbox_state,), daemon=True)
        thread.start()

    def run_cleaning(self, checkbox_state):
        total_freed = 0
        total_deleted = 0
        total_errors = 0

        # --- Кэш DNS (всегда) ---
        self.log(self.tr("log_dns_title"))
        dns_ok, dns_output = flush_dns()
        if dns_output:
            self.log(f"  {dns_output}")
        self.log("")

        all_folders = get_target_folders()

        # --- Папки, управляемые флажками (теперь у КАЖДОГО приложения своя
        # отдельная галочка — Chrome, Yandex, Opera Stable, Opera GX, Edge,
        # Firefox, Epic, EA, Spotify, CapCut, обе программы Adobe и т.д.) ---
        for checkbox_id, category, label_key, folder_ids, desc_key in CHECKBOX_DEFS:
            checkbox_label = self.tr(label_key)

            if not checkbox_state.get(checkbox_id, True):
                self.log(self.tr("log_user_skipped", label=checkbox_label))
                continue

            for folder_id in folder_ids:
                path_value = all_folders.get(folder_id)
                if not path_value:
                    continue

                # У Firefox может быть несколько профилей — path_value тогда
                # список путей вместо одной строки; обрабатываем оба случая.
                paths_to_clean = path_value if isinstance(path_value, list) else [path_value]
                if not paths_to_clean:
                    continue

                label = self.folder_name(folder_id)
                for path in paths_to_clean:
                    self.log(self.tr("log_cleaning", label=label))
                    self.log(self.tr("log_path", path=path))
                    freed, deleted, errors = clean_folder(path, self.log, self.tr)
                    total_freed += freed
                    total_deleted += deleted
                    total_errors += errors
                    self.log(self.tr(
                        "log_stats",
                        deleted=deleted,
                        freed=format_size(freed, self.current_lang),
                        errors=errors,
                    ))
                    self.log("")

        self.log("=" * 50)
        self.log(self.tr("log_total_deleted", count=total_deleted))
        self.log(self.tr("log_total_freed", size=format_size(total_freed, self.current_lang)))
        if total_errors:
            self.log(self.tr("log_total_skipped_note", count=total_errors))

        self.after(0, self.finish_cleaning, total_freed, total_deleted)

    def finish_cleaning(self, total_freed, total_deleted):
        self.set_buttons_state("normal")
        self.clean_button.configure(text=self.tr("clean_button"))
        size_str = format_size(total_freed, self.current_lang)
        self.status_label.configure(text=self.tr("status_done", count=total_deleted, size=size_str))
        messagebox.showinfo(
            self.tr("msg_clean_done_title"),
            self.tr("msg_clean_done_body", count=total_deleted, size=size_str),
        )

    # -----------------------------------------------------------------
    # Инструмент: восстановление кэша иконок
    # -----------------------------------------------------------------
    def start_icon_cache_rebuild(self):
        self.set_buttons_state("disabled")
        self.icon_cache_button.configure(text=self.tr("icon_cache_button_working"))
        self.status_label.configure(text=self.tr("status_icon_cache_working"))
        self.clear_log()
        self.log(self.tr("log_icon_cache_title"))

        thread = threading.Thread(target=self.run_icon_cache_rebuild, daemon=True)
        thread.start()

    def run_icon_cache_rebuild(self):
        success, tech_message = rebuild_icon_cache()
        if success:
            self.log(f"  {self.tr('status_icon_cache_done')}")
        else:
            self.log(f"  {self.tr('status_icon_cache_failed')} ({tech_message})")
        self.after(0, self.finish_icon_cache_rebuild, success)

    def finish_icon_cache_rebuild(self, success):
        self.set_buttons_state("normal")
        self.icon_cache_button.configure(text=self.tr("icon_cache_button"))
        self.status_label.configure(
            text=self.tr("status_icon_cache_done") if success else self.tr("status_icon_cache_failed")
        )
        if success:
            messagebox.showinfo(self.tr("msg_done_title"), self.tr("status_icon_cache_done"))
        else:
            messagebox.showwarning(self.tr("msg_failed_title"), self.tr("status_icon_cache_failed"))

    # -----------------------------------------------------------------
    # Инструмент: перезапуск проводника
    # -----------------------------------------------------------------
    def start_explorer_restart(self):
        self.set_buttons_state("disabled")
        self.explorer_button.configure(text=self.tr("explorer_button_working"))
        self.status_label.configure(text=self.tr("status_explorer_working"))
        self.clear_log()
        self.log(self.tr("log_explorer_title"))

        thread = threading.Thread(target=self.run_explorer_restart, daemon=True)
        thread.start()

    def run_explorer_restart(self):
        success, tech_message = restart_explorer()
        if success:
            self.log(f"  {self.tr('status_explorer_done')}")
        else:
            self.log(f"  {self.tr('status_explorer_failed')} ({tech_message})")
        self.after(0, self.finish_explorer_restart, success)

    def finish_explorer_restart(self, success):
        self.set_buttons_state("normal")
        self.explorer_button.configure(text=self.tr("explorer_button"))
        self.status_label.configure(
            text=self.tr("status_explorer_done") if success else self.tr("status_explorer_failed")
        )
        if success:
            messagebox.showinfo(self.tr("msg_done_title"), self.tr("status_explorer_done"))
        else:
            messagebox.showwarning(self.tr("msg_failed_title"), self.tr("status_explorer_failed"))

    # -----------------------------------------------------------------
    # Инструмент: очистка буфера обмена
    # -----------------------------------------------------------------
    def start_clipboard_clear(self):
        self.set_buttons_state("disabled")
        self.clipboard_button.configure(text=self.tr("clipboard_button_working"))
        self.status_label.configure(text=self.tr("status_clipboard_working"))
        self.clear_log()
        self.log(self.tr("log_clipboard_title"))

        thread = threading.Thread(target=self.run_clipboard_clear, daemon=True)
        thread.start()

    def run_clipboard_clear(self):
        success, tech_message = clear_clipboard()
        if success:
            self.log(f"  {self.tr('status_clipboard_done')}")
        else:
            self.log(f"  {self.tr('status_clipboard_failed')} ({tech_message})")
        self.after(0, self.finish_clipboard_clear, success)

    def finish_clipboard_clear(self, success):
        self.set_buttons_state("normal")
        self.clipboard_button.configure(text=self.tr("clipboard_button"))
        self.status_label.configure(
            text=self.tr("status_clipboard_done") if success else self.tr("status_clipboard_failed")
        )
        if success:
            messagebox.showinfo(self.tr("msg_done_title"), self.tr("status_clipboard_done"))
        else:
            messagebox.showwarning(self.tr("msg_failed_title"), self.tr("status_clipboard_failed"))


if __name__ == "__main__":
    app = CleanerApp()
    app.mainloop()
