#!/usr/bin/env python3
import base64
import io
import ipaddress
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import socket
import threading
import time
from typing import List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from PIL import Image, ImageSequence, ImageDraw, ImageTk
import pystray
from pystray import MenuItem as Item
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Divoom Keeper"
IMG_SIZE = 128
DEFAULT_QUALITY = 85
DEFAULT_SPEED = 100
SCREEN_COUNT = 5
PREVIEW_SIZE = 384


def appdata_dir() -> Path:
    base = os.environ.get("APPDATA")
    if not base:
        base = str(Path.home() / ".divoom-keeper")
    path = Path(base) / "DivoomKeeper"
    path.mkdir(parents=True, exist_ok=True)
    return path


CONFIG_PATH = appdata_dir() / "config.json"
LOG_PATH = appdata_dir() / "divoom-keeper.log"


def setup_logging() -> None:
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    handler = RotatingFileHandler(LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    handler.setFormatter(formatter)
    root.handlers.clear()
    root.addHandler(handler)


def default_config() -> dict:
    return {
        "device_ip": "192.168.1.116",
        "interval_minutes": 60,
        "quality": DEFAULT_QUALITY,
        "speed": DEFAULT_SPEED,
        "resend_on_startup": True,
        "start_with_windows": True,
        "ui_theme": "dark",
        "ui_lang": "en",
        "screens": [{"path": ""} for _ in range(SCREEN_COUNT)],
    }


class ConfigStore:
    def __init__(self, path: Path):
        self.path = path
        self.data = default_config()
        self.load()

    def load(self) -> None:
        if self.path.exists():
            try:
                raw = json.loads(self.path.read_text(encoding="utf-8"))
                merged = default_config()
                merged.update(raw)
                screens = raw.get("screens", merged["screens"])
                if not isinstance(screens, list):
                    screens = merged["screens"]
                while len(screens) < SCREEN_COUNT:
                    screens.append({"path": ""})
                merged["screens"] = screens[:SCREEN_COUNT]
                self.data = merged
            except Exception as e:
                logging.exception("Failed loading config, using defaults: %s", e)
                self.data = default_config()
        self.save()

    def save(self) -> None:
        self.path.write_text(json.dumps(self.data, indent=2, ensure_ascii=False), encoding="utf-8")


class DivoomSender:
    @staticmethod
    def lcd_array(screen: int) -> List[int]:
        arr = [0] * SCREEN_COUNT
        arr[screen - 1] = 1
        return arr

    @staticmethod
    def resize_image(img: Image.Image) -> Image.Image:
        return img.convert("RGB").resize((IMG_SIZE, IMG_SIZE), Image.LANCZOS)

    @staticmethod
    def load_frames(path: str) -> List[Image.Image]:
        img = Image.open(path)
        if getattr(img, "is_animated", False):
            return [frame.convert("RGB") for frame in ImageSequence.Iterator(img)]
        return [img.convert("RGB")]

    @classmethod
    def send_to_screen(
        cls,
        ip: str,
        screen: int,
        path: str,
        quality: int,
        speed: int,
        timeout: int = 10,
    ) -> None:
        frames = [cls.resize_image(f) for f in cls.load_frames(path)]
        pic_id = int(time.time())
        lcd = cls.lcd_array(screen)

        for offset, frame in enumerate(frames):
            buf = io.BytesIO()
            frame.save(buf, format="JPEG", quality=quality)
            payload = {
                "Command": "Draw/SendHttpGif",
                "LcdArray": lcd,
                "PicNum": len(frames),
                "PicOffset": offset,
                "PicID": pic_id,
                "PicSpeed": speed,
                "PicWidth": IMG_SIZE,
                "PicData": base64.b64encode(buf.getvalue()).decode(),
            }
            r = requests.post(f"http://{ip}/post", json=payload, timeout=timeout)
            r.raise_for_status()
            body = r.json()
            if body.get("error_code") != 0:
                raise RuntimeError(f"Divoom rejected frame {offset+1}: {body}")
            time.sleep(0.1)

    @staticmethod
    def _private_ip_prefixes(seed_ip: str = "") -> List[str]:
        prefixes = []

        def add_prefix(ip_str: str) -> None:
            try:
                ip = ipaddress.ip_address(ip_str)
                if ip.version == 4 and ip.is_private:
                    prefixes.append(".".join(ip_str.split(".")[:3]))
            except Exception:
                return

        if seed_ip:
            add_prefix(seed_ip)

        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, family=socket.AF_INET):
                add_prefix(info[4][0])
        except Exception:
            pass

        if not prefixes:
            prefixes.append("192.168.1")

        # keep order, remove duplicates
        dedup = []
        seen = set()
        for p in prefixes:
            if p not in seen:
                dedup.append(p)
                seen.add(p)
        return dedup

    @staticmethod
    def _probe_ip(ip: str, timeout: float = 0.45) -> bool:
        payload = {"Command": "Channel/Get5VVoltage"}
        try:
            r = requests.post(f"http://{ip}/post", json=payload, timeout=timeout)
            if r.status_code != 200:
                return False
            body = r.json()
            return isinstance(body, dict) and "error_code" in body
        except Exception:
            return False

    @classmethod
    def discover_devices(cls, seed_ip: str = "", timeout: float = 0.45) -> List[str]:
        prefixes = cls._private_ip_prefixes(seed_ip)
        candidates = [f"{prefix}.{i}" for prefix in prefixes for i in range(1, 255)]

        found: List[str] = []
        with ThreadPoolExecutor(max_workers=64) as ex:
            futures = {ex.submit(cls._probe_ip, ip, timeout): ip for ip in candidates}
            for fut in as_completed(futures):
                ip = futures[fut]
                try:
                    if fut.result():
                        found.append(ip)
                except Exception:
                    continue

        # prioritize subnet of current configured ip first (already first in prefixes)
        found.sort(key=lambda x: tuple(int(part) for part in x.split(".")))
        return found


class StartupManager:
    RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
    VALUE_NAME = "DivoomKeeper"

    @classmethod
    def is_windows(cls) -> bool:
        return os.name == "nt"

    @classmethod
    def set_enabled(cls, enabled: bool) -> None:
        if not cls.is_windows():
            return
        import winreg

        command = f'"{Path(__file__).resolve()}"'
        pythonw = Path(os.environ.get("LOCALAPPDATA", "")) / "Programs/Python"
        # Prefer current interpreter; hide console if packaged with pythonw/PyInstaller windowed
        exe = Path(os.sys.executable)
        if exe.exists():
            command = f'"{exe}" "{Path(__file__).resolve()}"'

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, cls.RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, cls.VALUE_NAME, 0, winreg.REG_SZ, command)
            else:
                try:
                    winreg.DeleteValue(key, cls.VALUE_NAME)
                except FileNotFoundError:
                    pass


class Scheduler(threading.Thread):
    def __init__(self, app: "KeeperApp"):
        super().__init__(daemon=True)
        self.app = app
        self.wakeup = threading.Event()
        self.stop_flag = threading.Event()

    def trigger_now(self) -> None:
        self.wakeup.set()

    def stop(self) -> None:
        self.stop_flag.set()
        self.wakeup.set()

    def run(self) -> None:
        while not self.stop_flag.is_set():
            interval = max(1, int(self.app.cfg.data.get("interval_minutes", 60)))
            triggered = self.wakeup.wait(timeout=interval * 60)
            self.wakeup.clear()
            if self.stop_flag.is_set():
                break
            reason = "manual" if triggered else "scheduled"
            self.app.send_all(reason=reason)


class KeeperUI:
    def __init__(self, app: "KeeperApp"):
        self.app = app
        self.root: Optional[tk.Tk] = None
        self.entries: List[tk.Entry] = []
        self.preview_labels: List[tk.Label] = []
        self.preview_refs: List[Optional[ImageTk.PhotoImage]] = [None] * SCREEN_COUNT
        self.preview_meta_labels: List[tk.Label] = []
        self.preview_anim_frames: List[List[ImageTk.PhotoImage]] = [[] for _ in range(SCREEN_COUNT)]
        self.preview_anim_idx: List[int] = [0] * SCREEN_COUNT
        self.preview_anim_after_id: List[Optional[str]] = [None] * SCREEN_COUNT
        self.health_label: Optional[tk.Label] = None
        self.startup_toggle_btn: Optional[tk.Button] = None
        self.big_preview_label: Optional[tk.Label] = None
        self.big_preview_meta: Optional[tk.Label] = None
        self.big_preview_ref: Optional[ImageTk.PhotoImage] = None
        self.active_preview_idx: int = 0
        self.ip_var = None
        self.interval_var = None
        self.quality_var = None
        self.speed_var = None
        self.resend_var = None
        self.startup_var = None
        self.theme_var = None
        self.theme_box = None
        self.lang_var = None
        self.lang_box = None
        self.lang = str(self.app.cfg.data.get("ui_lang", "en"))
        self.colors = {}

    def t(self, key: str, **kwargs) -> str:
        i18n = {
            "en": {
                "connection_schedule": "Connection & schedule",
                "divoom_ip": "Divoom IP",
                "interval_min": "Interval (min)",
                "quality": "Quality",
                "speed": "Speed",
                "theme": "Theme",
                "language": "Language",
                "resend_on_startup": "Resend on startup",
                "start_with_windows": "Start with Windows",
                "device_checking": "Device: checking...",
                "screen_slots": "Screen slots (with preview)",
                "screen_n": "Screen {n}",
                "no_preview": "No preview",
                "browse": "Browse",
                "send": "Send",
                "save": "Save",
                "send_all": "Send all now",
                "scan_lan": "Scan Divoom in LAN",
                "hide_tray": "Hide to tray",
                "no_file_selected": "No file selected",
                "preview_error": "Preview error",
                "preview_failed": "Preview failed: {err}",
                "select_device": "Select Divoom device",
                "detected_devices": "Detected Divoom-like devices",
                "use_selected": "Use selected",
                "cancel": "Cancel",
                "pick_media": "Select image/GIF for screen {n}",
                "config_saved": "Config saved",
                "save_failed": "Save failed: {err}",
                "screen_no_file": "Screen {n} has no file",
                "startup_enabled": "Startup enabled",
                "startup_disabled": "Startup disabled",
                "startup_toggle_failed": "Failed toggling startup: {err}",
                "no_device_scan": "No Divoom device detected in local subnet scan",
                "detected_one_device": "Detected 1 device: {ip}\n\nSet it as active IP?",
                "detected_multi_devices": "Detected devices:\n- {list}\n\nUse selected device from the dialog.",
                "scan_keep_ip": "Scan finished. Kept current IP unchanged.",
                "active_ip_updated": "Active IP updated to {ip}",
                "scan_failed": "Scan failed: {err}",
                "device_online": "Device {ip}: ONLINE",
                "device_offline": "Device {ip}: OFFLINE",
                "startup_on": "Auto-start: ON (click to disable)",
                "startup_off": "Auto-start: OFF (click to enable)",
                "tray_open": "Open",
                "tray_send_now": "Send now",
                "tray_scan_network": "Scan network",
                "tray_quit": "Quit",
                "media_gif": "GIF (animated)",
                "media_img": "IMG",
            },
            "es": {
                "connection_schedule": "Conexión y programación",
                "divoom_ip": "IP de Divoom",
                "interval_min": "Intervalo (min)",
                "quality": "Calidad",
                "speed": "Velocidad",
                "theme": "Tema",
                "language": "Idioma",
                "resend_on_startup": "Reenviar al iniciar",
                "start_with_windows": "Iniciar con Windows",
                "device_checking": "Dispositivo: comprobando...",
                "screen_slots": "Ranuras de pantalla (con vista previa)",
                "screen_n": "Pantalla {n}",
                "no_preview": "Sin vista previa",
                "browse": "Buscar",
                "send": "Enviar",
                "save": "Guardar",
                "send_all": "Enviar todo ahora",
                "scan_lan": "Buscar Divoom en LAN",
                "hide_tray": "Ocultar en bandeja",
                "no_file_selected": "No hay archivo seleccionado",
                "preview_error": "Error de vista previa",
                "preview_failed": "Fallo de vista previa: {err}",
                "select_device": "Seleccionar dispositivo Divoom",
                "detected_devices": "Dispositivos Divoom detectados",
                "use_selected": "Usar seleccionado",
                "cancel": "Cancelar",
                "pick_media": "Seleccionar imagen/GIF para pantalla {n}",
                "config_saved": "Configuración guardada",
                "save_failed": "Error al guardar: {err}",
                "screen_no_file": "La pantalla {n} no tiene archivo",
                "startup_enabled": "Inicio automático activado",
                "startup_disabled": "Inicio automático desactivado",
                "startup_toggle_failed": "Error al cambiar inicio automático: {err}",
                "no_device_scan": "No se detectó ningún Divoom en el escaneo de subred local",
                "detected_one_device": "Se detectó 1 dispositivo: {ip}\n\n¿Usarlo como IP activa?",
                "detected_multi_devices": "Dispositivos detectados:\n- {list}\n\nUsa el diálogo para elegir el dispositivo.",
                "scan_keep_ip": "Escaneo terminado. Se mantiene la IP actual.",
                "active_ip_updated": "IP activa actualizada a {ip}",
                "scan_failed": "Error de escaneo: {err}",
                "device_online": "Dispositivo {ip}: EN LÍNEA",
                "device_offline": "Dispositivo {ip}: DESCONECTADO",
                "startup_on": "Autoarranque: ON (clic para desactivar)",
                "startup_off": "Autoarranque: OFF (clic para activar)",
                "tray_open": "Abrir",
                "tray_send_now": "Enviar ahora",
                "tray_scan_network": "Buscar en red",
                "tray_quit": "Salir",
                "media_gif": "GIF (animado)",
                "media_img": "IMG",
            },
        }
        lang = self.lang if self.lang in i18n else "en"
        template = i18n[lang].get(key, i18n["en"].get(key, key))
        return template.format(**kwargs)

    def _palette(self, theme: str) -> dict:
        if theme == "light":
            return {
                "bg": "#f4f6fb",
                "fg": "#1e2430",
                "muted": "#5f6b7a",
                "input": "#ffffff",
                "accent": "#0078d4",
                "button": "#e2e8f0",
                "ok": "#0a8f48",
                "warn": "#9a6b00",
                "err": "#b42318",
                "preview_bg": "#e7ecf4",
            }
        return {
            "bg": "#1d1f24",
            "fg": "#e8ecf2",
            "muted": "#9ca8b7",
            "input": "#2a2f38",
            "accent": "#1d9bf0",
            "button": "#344054",
            "ok": "#57d38c",
            "warn": "#f6c177",
            "err": "#ff8b8b",
            "preview_bg": "#2a2f38",
        }

    def ensure_window(self) -> None:
        if self.root is not None:
            self.root.deiconify()
            self.root.lift()
            return

        theme = str(self.app.cfg.data.get("ui_theme", "dark")).lower()
        self.lang = str(self.app.cfg.data.get("ui_lang", "en")).lower()
        self.colors = self._palette("light" if theme == "light" else "dark")

        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1320x860")
        self.root.minsize(1180, 760)
        self.root.configure(bg=self.colors["bg"])
        self.root.protocol("WM_DELETE_WINDOW", self.hide)

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Card.TLabelframe", background=self.colors["bg"], foreground=self.colors["fg"])
        style.configure("Card.TLabelframe.Label", background=self.colors["bg"], foreground=self.colors["accent"])

        self.ip_var = tk.StringVar(master=self.root, value=self.app.cfg.data["device_ip"])
        self.interval_var = tk.StringVar(master=self.root, value=str(self.app.cfg.data["interval_minutes"]))
        self.quality_var = tk.StringVar(master=self.root, value=str(self.app.cfg.data["quality"]))
        self.speed_var = tk.StringVar(master=self.root, value=str(self.app.cfg.data["speed"]))
        self.resend_var = tk.BooleanVar(master=self.root, value=bool(self.app.cfg.data.get("resend_on_startup", True)))
        self.startup_var = tk.BooleanVar(master=self.root, value=bool(self.app.cfg.data.get("start_with_windows", True)))
        self.theme_var = tk.StringVar(master=self.root, value=("Light" if theme == "light" else "Dark"))
        self.lang_var = tk.StringVar(master=self.root, value=("Español" if self.lang == "es" else "English"))

        top = ttk.LabelFrame(self.root, text=self.t("connection_schedule"), style="Card.TLabelframe")
        top.pack(fill="x", padx=12, pady=(12, 8))

        tk.Label(top, text=self.t("divoom_ip"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=0, sticky="w")
        tk.Entry(top, textvariable=self.ip_var, width=16, bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"]).grid(row=0, column=1, sticky="w", padx=6)

        tk.Label(top, text=self.t("interval_min"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=2, sticky="w")
        tk.Entry(top, textvariable=self.interval_var, width=8, bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"]).grid(row=0, column=3, sticky="w", padx=6)

        tk.Label(top, text=self.t("quality"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=4, sticky="w")
        tk.Entry(top, textvariable=self.quality_var, width=6, bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"]).grid(row=0, column=5, sticky="w", padx=6)

        tk.Label(top, text=self.t("speed"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=6, sticky="w")
        tk.Entry(top, textvariable=self.speed_var, width=6, bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"]).grid(row=0, column=7, sticky="w", padx=6)

        tk.Label(top, text=self.t("theme"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=8, sticky="w")
        self.theme_box = ttk.Combobox(top, textvariable=self.theme_var, values=["Dark", "Light"], width=8, state="readonly")
        self.theme_box.grid(row=0, column=9, sticky="w", padx=6)
        self.theme_box.bind("<<ComboboxSelected>>", lambda _e: self.on_theme_changed())

        tk.Label(top, text=self.t("language"), bg=self.colors["bg"], fg=self.colors["fg"]).grid(row=0, column=10, sticky="w")
        self.lang_box = ttk.Combobox(top, textvariable=self.lang_var, values=["English", "Español"], width=10, state="readonly")
        self.lang_box.grid(row=0, column=11, sticky="w", padx=6)
        self.lang_box.bind("<<ComboboxSelected>>", lambda _e: self.on_language_changed())

        tk.Checkbutton(top, text=self.t("resend_on_startup"), variable=self.resend_var, bg=self.colors["bg"], fg=self.colors["fg"], selectcolor=self.colors["bg"], activebackground=self.colors["bg"]).grid(row=1, column=0, columnspan=3, sticky="w", pady=6)
        tk.Checkbutton(top, text=self.t("start_with_windows"), variable=self.startup_var, bg=self.colors["bg"], fg=self.colors["fg"], selectcolor=self.colors["bg"], activebackground=self.colors["bg"]).grid(row=1, column=3, columnspan=3, sticky="w", pady=6)

        self.health_label = tk.Label(top, text=self.t("device_checking"), bg=self.colors["bg"], fg=self.colors["warn"], font=("Segoe UI", 9, "bold"))
        self.health_label.grid(row=1, column=6, columnspan=2, sticky="e", padx=6)

        grid = ttk.LabelFrame(self.root, text=self.t("screen_slots"), style="Card.TLabelframe")
        grid.pack(fill="both", expand=True, padx=12, pady=8)

        for i in range(SCREEN_COUNT):
            row_base = i * 2
            tk.Label(grid, text=self.t("screen_n", n=i+1), bg=self.colors["bg"], fg=self.colors["accent"], font=("Segoe UI", 10, "bold")).grid(row=row_base, column=0, sticky="nw", pady=(8, 2), padx=(8, 4))

            preview = tk.Label(grid, text=self.t("no_preview"), width=1, height=1, bg=self.colors["input"], fg=self.colors["muted"], relief="groove", padx=8, pady=8, cursor="hand2")
            preview.grid(row=row_base, column=1, rowspan=2, sticky="w", pady=(8, 8), padx=(0, 8))
            preview.bind("<Button-1>", lambda _e, idx=i: (self.set_active_preview(idx), self.open_preview_zoom(idx)))
            self.preview_labels.append(preview)

            entry = tk.Entry(grid, width=78, bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"])
            entry.insert(0, self.app.cfg.data["screens"][i].get("path", ""))
            entry.grid(row=row_base, column=2, sticky="we", padx=6)
            self.entries.append(entry)
            entry.bind("<FocusOut>", lambda _e, idx=i: self.refresh_preview(idx))
            entry.bind("<FocusIn>", lambda _e, idx=i: self.set_active_preview(idx))

            actions = tk.Frame(grid, bg=self.colors["bg"])
            actions.grid(row=row_base, column=3, sticky="e", padx=4)
            tk.Button(actions, text=self.t("browse"), command=lambda idx=i: self.pick_file(idx), bg=self.colors["button"], fg=self.colors["fg"]).pack(side="left", padx=2)
            tk.Button(actions, text=self.t("send"), command=lambda idx=i: self.send_one(idx), bg=self.colors["accent"], fg=self.colors["fg"]).pack(side="left", padx=2)

            meta = tk.Label(grid, text="", bg=self.colors["bg"], fg=self.colors["muted"], anchor="w")
            meta.grid(row=row_base + 1, column=2, columnspan=2, sticky="we", padx=6, pady=(0, 6))
            self.preview_meta_labels.append(meta)

        grid.grid_columnconfigure(2, weight=1)

        for i in range(SCREEN_COUNT):
            self.refresh_preview(i)

        preview_panel = ttk.LabelFrame(self.root, text="Live preview", style="Card.TLabelframe")
        preview_panel.pack(fill="both", expand=False, padx=12, pady=(0, 8))
        self.big_preview_label = tk.Label(preview_panel, text=self.t("no_preview"), bg=self.colors["input"], fg=self.colors["muted"], relief="groove", padx=8, pady=8)
        self.big_preview_label.pack(side="left", padx=10, pady=10)
        self.big_preview_meta = tk.Label(preview_panel, text="", bg=self.colors["bg"], fg=self.colors["muted"], justify="left", anchor="w")
        self.big_preview_meta.pack(side="left", fill="both", expand=True, padx=10)

        # Default to first non-empty slot
        first_non_empty = 0
        for i, e in enumerate(self.entries):
            if e.get().strip():
                first_non_empty = i
                break
        self.active_preview_idx = first_non_empty
        self.refresh_big_preview(self.active_preview_idx)

        self.refresh_health()

        footer = tk.Frame(self.root, bg=self.colors["bg"])
        footer.pack(fill="x", padx=10, pady=8)
        tk.Button(footer, text=self.t("save"), command=self.save, bg=self.colors["button"], fg=self.colors["fg"]).pack(side="left")
        tk.Button(footer, text=self.t("send_all"), command=self.send_all_now, bg=self.colors["accent"], fg=self.colors["fg"]).pack(side="left", padx=8)
        tk.Button(footer, text=self.t("scan_lan"), command=self.scan_devices, bg=self.colors["accent"], fg=self.colors["fg"]).pack(side="left", padx=8)
        self.startup_toggle_btn = tk.Button(footer, command=self.toggle_startup_now, bg=self.colors["button"], fg=self.colors["fg"])
        self.startup_toggle_btn.pack(side="left", padx=8)
        self.update_startup_button()
        tk.Button(footer, text=self.t("hide_tray"), command=self.hide, bg=self.colors["button"], fg=self.colors["fg"]).pack(side="right")

    def _cancel_preview_anim(self, idx: int) -> None:
        if self.root is not None and self.preview_anim_after_id[idx] is not None:
            try:
                self.root.after_cancel(self.preview_anim_after_id[idx])
            except Exception:
                pass
        self.preview_anim_after_id[idx] = None
        self.preview_anim_frames[idx] = []
        self.preview_anim_idx[idx] = 0

    def _tick_preview_anim(self, idx: int) -> None:
        if self.root is None:
            return
        frames = self.preview_anim_frames[idx]
        if not frames:
            self.preview_anim_after_id[idx] = None
            return
        self.preview_anim_idx[idx] = (self.preview_anim_idx[idx] + 1) % len(frames)
        frame = frames[self.preview_anim_idx[idx]]
        self.preview_labels[idx].configure(image=frame, text="")
        self.preview_refs[idx] = frame
        self.preview_anim_after_id[idx] = self.root.after(180, lambda: self._tick_preview_anim(idx))

    def set_active_preview(self, idx: int) -> None:
        self.active_preview_idx = idx
        self.refresh_big_preview(idx)

    def refresh_big_preview(self, idx: int) -> None:
        if self.root is None or self.big_preview_label is None or idx >= len(self.entries):
            return
        path = self.entries[idx].get().strip()
        if not path or not Path(path).exists():
            self.big_preview_label.configure(image="", text=self.t("no_preview"), bg=self.colors["input"], fg=self.colors["muted"])
            if self.big_preview_meta is not None:
                self.big_preview_meta.configure(text=self.t("no_file_selected"))
            self.big_preview_ref = None
            return
        try:
            img = Image.open(path).convert("RGB")
            img.thumbnail((560, 560), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            self.big_preview_label.configure(image=tk_img, text="")
            self.big_preview_ref = tk_img
            if self.big_preview_meta is not None:
                self.big_preview_meta.configure(text=f"{self.t('screen_n', n=idx+1)}\n{Path(path).name}\n{img.width}x{img.height}")
        except Exception as e:
            self.big_preview_label.configure(image="", text=self.t("preview_error"), bg=self.colors["preview_bg"], fg=self.colors["err"])
            if self.big_preview_meta is not None:
                self.big_preview_meta.configure(text=self.t("preview_failed", err=e))
            self.big_preview_ref = None

    def open_preview_zoom(self, idx: int) -> None:
        if idx >= len(self.entries) or self.root is None:
            return
        path = self.entries[idx].get().strip()
        if not path or not Path(path).exists():
            return

        dlg = tk.Toplevel(self.root)
        dlg.title(self.t("screen_n", n=idx + 1))
        dlg.configure(bg=self.colors["bg"])
        dlg.geometry("760x760")
        dlg.minsize(520, 520)

        holder = tk.Label(dlg, bg=self.colors["bg"])
        holder.pack(fill="both", expand=True, padx=12, pady=12)

        try:
            img = Image.open(path).convert("RGB")
            img.thumbnail((700, 700), Image.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            holder.configure(image=tk_img)
            holder.image = tk_img
        except Exception as e:
            holder.configure(text=self.t("preview_failed", err=e), fg=self.colors["err"])

    def refresh_preview(self, idx: int) -> None:
        if idx >= len(self.entries):
            return
        path = self.entries[idx].get().strip()
        label = self.preview_labels[idx]
        meta = self.preview_meta_labels[idx]

        self._cancel_preview_anim(idx)

        if not path or not Path(path).exists():
            label.configure(image="", text=self.t("no_preview"), bg=self.colors["input"], fg=self.colors["muted"])
            meta.configure(text=self.t("no_file_selected"))
            self.preview_refs[idx] = None
            return

        try:
            img = Image.open(path)
            is_gif = getattr(img, "is_animated", False)

            if is_gif:
                frames = []
                for frame in ImageSequence.Iterator(img):
                    thumb = frame.convert("RGB").resize((PREVIEW_SIZE, PREVIEW_SIZE), Image.LANCZOS)
                    frames.append(ImageTk.PhotoImage(thumb))
                    if len(frames) >= 24:
                        break
                if frames:
                    self.preview_anim_frames[idx] = frames
                    self.preview_anim_idx[idx] = 0
                    label.configure(image=frames[0], text="")
                    self.preview_refs[idx] = frames[0]
                    self.preview_anim_after_id[idx] = self.root.after(180, lambda: self._tick_preview_anim(idx))
            else:
                frame = img.convert("RGB")
                thumb = frame.resize((PREVIEW_SIZE, PREVIEW_SIZE), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(thumb)
                label.configure(image=tk_img, text="")
                self.preview_refs[idx] = tk_img

            ext = Path(path).suffix.lower().lstrip(".")
            note = self.t("media_gif") if is_gif else ext.upper() if ext else self.t("media_img")
            meta.configure(text=f"{note} • {img.width}x{img.height} • {Path(path).name}")
        except Exception as e:
            label.configure(image="", text=self.t("preview_error"), bg=self.colors["preview_bg"], fg=self.colors["err"])
            meta.configure(text=self.t("preview_failed", err=e))
            self.preview_refs[idx] = None

        if idx == self.active_preview_idx:
            self.refresh_big_preview(idx)

    def probe_health(self) -> bool:
        ip = self.ip_var.get().strip() if self.ip_var else self.app.cfg.data.get("device_ip", "")
        if not ip:
            return False
        return DivoomSender._probe_ip(ip, timeout=0.6)

    def refresh_health(self) -> None:
        if self.root is None or self.health_label is None:
            return

        def worker():
            ok = self.probe_health()
            ip = self.ip_var.get().strip() if self.ip_var else self.app.cfg.data.get("device_ip", "")

            def done():
                if self.health_label is None:
                    return
                if ok:
                    self.health_label.configure(text=self.t("device_online", ip=ip), fg=self.colors["ok"])
                else:
                    self.health_label.configure(text=self.t("device_offline", ip=ip), fg=self.colors["err"])
                if self.root is not None:
                    self.root.after(15000, self.refresh_health)

            if self.root is not None:
                self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _retheme_widget_tree(self, widget) -> None:
        cls = widget.winfo_class()
        try:
            if cls in {"Frame", "TFrame", "Labelframe", "TLabelframe", "Toplevel"}:
                widget.configure(bg=self.colors["bg"])
            elif cls == "Label":
                fg = widget.cget("fg")
                if fg in ("#7fdcff", "#1d9bf0", self.colors.get("accent")):
                    fg = self.colors["accent"]
                elif fg in ("#9ca8b7", self.colors.get("muted")):
                    fg = self.colors["muted"]
                elif fg in ("#57d38c", self.colors.get("ok")):
                    fg = self.colors["ok"]
                elif fg in ("#ff8b8b", self.colors.get("err")):
                    fg = self.colors["err"]
                else:
                    fg = self.colors["fg"]
                widget.configure(bg=self.colors["bg"], fg=fg)
            elif cls == "Entry":
                widget.configure(bg=self.colors["input"], fg=self.colors["fg"], insertbackground=self.colors["fg"])
            elif cls == "Button":
                txt = str(widget.cget("text")).lower()
                if "send" in txt or "scan" in txt:
                    widget.configure(bg=self.colors["accent"], fg=self.colors["fg"], activebackground=self.colors["button"])
                else:
                    widget.configure(bg=self.colors["button"], fg=self.colors["fg"], activebackground=self.colors["accent"])
            elif cls == "Checkbutton":
                widget.configure(bg=self.colors["bg"], fg=self.colors["fg"], selectcolor=self.colors["bg"], activebackground=self.colors["bg"])
            elif cls == "Listbox":
                widget.configure(bg=self.colors["input"], fg=self.colors["fg"], selectbackground=self.colors["accent"])
        except Exception:
            pass

        for child in widget.winfo_children():
            self._retheme_widget_tree(child)

    def on_theme_changed(self) -> None:
        if self.root is None or self.theme_var is None:
            return
        selected = self.theme_var.get().lower()
        self.colors = self._palette("light" if selected == "light" else "dark")
        self.app.cfg.data["ui_theme"] = "light" if selected == "light" else "dark"
        self.app.cfg.save()

        self.root.configure(bg=self.colors["bg"])
        style = ttk.Style(self.root)
        style.configure("Card.TLabelframe", background=self.colors["bg"], foreground=self.colors["fg"])
        style.configure("Card.TLabelframe.Label", background=self.colors["bg"], foreground=self.colors["accent"])
        self._retheme_widget_tree(self.root)

        for i in range(len(self.entries)):
            self.refresh_preview(i)
        self.update_startup_button()
        self.refresh_health()

    def rebuild_window(self) -> None:
        if self.root is None:
            return
        was_visible = self.root.state() != "withdrawn"
        try:
            self.root.destroy()
        except Exception:
            pass
        self.root = None
        self.entries = []
        self.preview_labels = []
        self.preview_meta_labels = []
        self.preview_refs = [None] * SCREEN_COUNT
        self.preview_anim_frames = [[] for _ in range(SCREEN_COUNT)]
        self.preview_anim_idx = [0] * SCREEN_COUNT
        self.preview_anim_after_id = [None] * SCREEN_COUNT
        self.ensure_window()
        if not was_visible:
            self.hide()

    def on_language_changed(self) -> None:
        if self.lang_var is None:
            return
        selected = self.lang_var.get()
        self.lang = "es" if selected.lower().startswith("es") else "en"
        self.app.cfg.data["ui_lang"] = self.lang
        self.app.cfg.save()
        self.rebuild_window()
        self.app.refresh_tray_menu()

    def choose_device_dialog(self, found: List[str], suggested: str) -> Optional[str]:
        if self.root is None:
            return None

        result = {"ip": None}
        dlg = tk.Toplevel(self.root)
        dlg.title(self.t("select_device"))
        dlg.geometry("360x280")
        dlg.configure(bg=self.colors["bg"])
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text=self.t("detected_devices"), bg=self.colors["bg"], fg=self.colors["fg"], font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

        lst = tk.Listbox(dlg, bg=self.colors["input"], fg=self.colors["fg"], selectbackground=self.colors["accent"], height=10)
        for ip in found:
            lst.insert(tk.END, ip)
        try:
            idx = found.index(suggested)
        except ValueError:
            idx = 0
        lst.select_set(idx)
        lst.pack(fill="both", expand=True, padx=10, pady=6)

        def use_selected():
            sel = lst.curselection()
            if sel:
                result["ip"] = lst.get(sel[0])
            dlg.destroy()

        btns = tk.Frame(dlg, bg=self.colors["bg"])
        btns.pack(fill="x", padx=10, pady=10)
        tk.Button(btns, text=self.t("use_selected"), command=use_selected, bg=self.colors["accent"], fg=self.colors["fg"]).pack(side="left")
        tk.Button(btns, text=self.t("cancel"), command=dlg.destroy, bg=self.colors["button"], fg=self.colors["fg"]).pack(side="right")

        self.root.wait_window(dlg)
        return result["ip"]

    def pick_file(self, idx: int) -> None:
        path = filedialog.askopenfilename(
            title=self.t("pick_media", n=idx+1),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All", "*.*")],
        )
        if path:
            self.entries[idx].delete(0, tk.END)
            self.entries[idx].insert(0, path)
            self.refresh_preview(idx)

    def save(self) -> None:
        try:
            self.app.cfg.data["device_ip"] = self.ip_var.get().strip()
            self.app.cfg.data["interval_minutes"] = max(1, int(self.interval_var.get().strip()))
            self.app.cfg.data["quality"] = min(100, max(30, int(self.quality_var.get().strip())))
            self.app.cfg.data["speed"] = max(1, int(self.speed_var.get().strip()))
            self.app.cfg.data["resend_on_startup"] = bool(self.resend_var.get())
            self.app.cfg.data["start_with_windows"] = bool(self.startup_var.get())
            self.app.cfg.data["ui_theme"] = "light" if (self.theme_var and self.theme_var.get().lower() == "light") else "dark"
            for i, entry in enumerate(self.entries):
                self.app.cfg.data["screens"][i]["path"] = entry.get().strip()
                self.refresh_preview(i)
            self.app.cfg.save()
            StartupManager.set_enabled(self.app.cfg.data["start_with_windows"])
            self.update_startup_button()
            self.app.scheduler.trigger_now()
            self.refresh_health()
            messagebox.showinfo(APP_NAME, self.t("config_saved"))
        except Exception as e:
            logging.exception("Save failed")
            messagebox.showerror(APP_NAME, self.t("save_failed", err=e))

    def send_one(self, idx: int) -> None:
        path = self.entries[idx].get().strip()
        if not path:
            messagebox.showwarning(APP_NAME, self.t("screen_no_file", n=idx+1))
            return
        self.app.send_screen(idx + 1, path)

    def send_all_now(self) -> None:
        self.save()
        self.app.scheduler.trigger_now()

    def update_startup_button(self) -> None:
        btn = getattr(self, "startup_toggle_btn", None)
        enabled = bool(self.app.cfg.data.get("start_with_windows", False))
        if btn is None:
            return
        if enabled:
            btn.configure(text=self.t("startup_on"), bg=self.colors["ok"], fg=self.colors["bg"])
        else:
            btn.configure(text=self.t("startup_off"), bg=self.colors["button"], fg=self.colors["fg"])

    def toggle_startup_now(self) -> None:
        try:
            enabled = not bool(self.app.cfg.data.get("start_with_windows", False))
            StartupManager.set_enabled(enabled)
            if self.startup_var is not None:
                self.startup_var.set(enabled)
            self.app.cfg.data["start_with_windows"] = enabled
            self.app.cfg.save()
            self.update_startup_button()
            messagebox.showinfo(APP_NAME, self.t("startup_enabled") if enabled else self.t("startup_disabled"))
        except Exception as e:
            messagebox.showerror(APP_NAME, self.t("startup_toggle_failed", err=e))

    def scan_devices(self) -> None:
        self.save()

        def worker():
            try:
                seed_ip = self.ip_var.get().strip() if self.ip_var else self.app.cfg.data.get("device_ip", "")
                found = DivoomSender.discover_devices(seed_ip=seed_ip)

                def done():
                    if not found:
                        messagebox.showwarning(APP_NAME, self.t("no_device_scan"))
                        return

                    current = self.ip_var.get().strip() if self.ip_var else ""
                    chosen = found[0]

                    if len(found) == 1:
                        apply_ip = messagebox.askyesno(
                            APP_NAME,
                            self.t("detected_one_device", ip=chosen),
                        )
                        if not apply_ip:
                            if current:
                                messagebox.showinfo(APP_NAME, self.t("scan_keep_ip"))
                            return
                    else:
                        selected = self.choose_device_dialog(found, suggested=chosen)
                        if not selected:
                            if current:
                                messagebox.showinfo(APP_NAME, self.t("scan_keep_ip"))
                            return
                        chosen = selected

                    if self.ip_var is not None:
                        self.ip_var.set(chosen)
                    self.app.cfg.data["device_ip"] = chosen
                    self.app.cfg.save()
                    self.refresh_health()
                    messagebox.showinfo(APP_NAME, self.t("active_ip_updated", ip=chosen))

                if self.root is not None:
                    self.root.after(0, done)
            except Exception as e:
                logging.exception("Scan failed: %s", e)
                if self.root is not None:
                    self.root.after(0, lambda: messagebox.showerror(APP_NAME, self.t("scan_failed", err=e)))

        threading.Thread(target=worker, daemon=True).start()

    def hide(self) -> None:
        if self.root is not None:
            self.root.withdraw()

    def run(self) -> None:
        self.ensure_window()
        self.hide()
        self.root.mainloop()


class KeeperApp:
    def __init__(self):
        setup_logging()
        self.cfg = ConfigStore(CONFIG_PATH)
        self.scheduler = Scheduler(self)
        self.ui = KeeperUI(self)
        self.icon: Optional[pystray.Icon] = None

    def refresh_tray_menu(self) -> None:
        if not self.icon:
            return
        self.icon.menu = pystray.Menu(
            Item(self.ui.t("tray_open"), self.open_ui),
            Item(self.ui.t("tray_send_now"), self.send_now),
            Item(self.ui.t("tray_scan_network"), self.scan_network),
            Item(self.ui.t("tray_quit"), self.quit),
        )
        self.icon.update_menu()

    def send_screen(self, screen: int, path: str) -> bool:
        if not Path(path).exists():
            logging.warning("Screen %s skipped, missing file: %s", screen, path)
            return False
        try:
            DivoomSender.send_to_screen(
                ip=self.cfg.data["device_ip"],
                screen=screen,
                path=path,
                quality=int(self.cfg.data["quality"]),
                speed=int(self.cfg.data["speed"]),
            )
            logging.info("Screen %s sent: %s", screen, path)
            return True
        except Exception as e:
            logging.exception("Screen %s send failed: %s", screen, e)
            return False

    def send_all(self, reason: str = "manual") -> None:
        logging.info("Send all triggered (%s)", reason)
        targets = []
        for i in range(SCREEN_COUNT):
            path = self.cfg.data["screens"][i].get("path", "").strip()
            if path:
                targets.append((i + 1, path))

        if not targets:
            return

        ok = 0
        for screen, path in targets:
            if self.send_screen(screen, path):
                ok += 1

        if ok == 0:
            logging.warning("No screen sent successfully. Running LAN scan for Divoom fallback.")
            found = DivoomSender.discover_devices(seed_ip=self.cfg.data.get("device_ip", ""))
            if found:
                self.cfg.data["device_ip"] = found[0]
                self.cfg.save()
                logging.info("Auto-discovery selected device IP: %s", found[0])
                for screen, path in targets:
                    self.send_screen(screen, path)

    def tray_image(self) -> Image.Image:
        img = Image.new("RGB", (64, 64), color=(35, 35, 40))
        draw = ImageDraw.Draw(img)
        draw.rectangle((8, 8, 56, 56), outline=(0, 210, 255), width=3)
        draw.text((18, 20), "D", fill=(0, 210, 255))
        return img

    def open_ui(self, icon=None, item=None):
        def show():
            self.ui.ensure_window()
        threading.Thread(target=show, daemon=True).start()

    def send_now(self, icon=None, item=None):
        self.scheduler.trigger_now()

    def scan_network(self, icon=None, item=None):
        def worker():
            found = DivoomSender.discover_devices(seed_ip=self.cfg.data.get("device_ip", ""))
            if found:
                self.cfg.data["device_ip"] = found[0]
                self.cfg.save()
                logging.info("Tray scan selected IP: %s", found[0])
            else:
                logging.warning("Tray scan found no devices")
        threading.Thread(target=worker, daemon=True).start()

    def quit(self, icon=None, item=None):
        logging.info("Shutting down")
        self.scheduler.stop()
        if self.icon:
            self.icon.stop()
        if self.ui.root:
            self.ui.root.quit()

    def run(self) -> None:
        StartupManager.set_enabled(bool(self.cfg.data.get("start_with_windows", True)))

        self.scheduler.start()

        if self.cfg.data.get("resend_on_startup", True):
            threading.Thread(target=lambda: self.send_all(reason="startup"), daemon=True).start()

        menu = pystray.Menu(
            Item(self.ui.t("tray_open"), self.open_ui),
            Item(self.ui.t("tray_send_now"), self.send_now),
            Item(self.ui.t("tray_scan_network"), self.scan_network),
            Item(self.ui.t("tray_quit"), self.quit),
        )
        self.icon = pystray.Icon("divoom_keeper", self.tray_image(), APP_NAME, menu)

        ui_thread = threading.Thread(target=self.ui.run, daemon=True)
        ui_thread.start()

        self.icon.run()


if __name__ == "__main__":
    app = KeeperApp()
    app.run()
