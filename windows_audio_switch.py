import argparse
import ctypes
import datetime as _datetime
import json
import os
import shutil
import sys
import threading
import tkinter as tk
import traceback
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable, Optional
from winreg import HKEY_CURRENT_USER, KEY_SET_VALUE, REG_SZ, DeleteValue, OpenKey, QueryValueEx, SetValueEx

import comtypes
from comtypes import CLSCTX_ALL, COMMETHOD, GUID, HRESULT, IUnknown, POINTER
from ctypes import byref, c_int, c_ushort, c_void_p, c_wchar_p, windll
from ctypes import wintypes

try:
    import pystray
    from PIL import Image, ImageDraw
except ImportError:
    pystray = None
    Image = None
    ImageDraw = None


APP_NAME = "Windows Audio Switch"
CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / APP_NAME
CONFIG_PATH = CONFIG_DIR / "config.json"
LOG_PATH = CONFIG_DIR / "WindowsAudioSwitch.log"
STARTUP_RUN_NAME = "Windows Audio Switch"
INSTANCE_MUTEX_NAME = "Local\\WindowsAudioSwitchDaemon"

CONFIG_VERSION = 2
DEFAULT_HOTKEY = "ctrl+alt+a"
ERROR_ALREADY_EXISTS = 183

DEVICE_STATE_ACTIVE = 0x00000001
E_RENDER = 0
E_CONSOLE = 0
E_MULTIMEDIA = 1
E_COMMUNICATIONS = 2
STGM_READ = 0
VT_LPWSTR = 31

class PROPERTYKEY(ctypes.Structure):
    _fields_ = [("fmtid", GUID), ("pid", wintypes.DWORD)]


class PROPVARIANT(ctypes.Structure):
    _fields_ = [
        ("vt", c_ushort),
        ("wReserved1", c_ushort),
        ("wReserved2", c_ushort),
        ("wReserved3", c_ushort),
        ("pwszVal", c_wchar_p),
    ]


class IPropertyStore(IUnknown):
    _iid_ = GUID("{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(wintypes.DWORD), "cProps")),
        COMMETHOD([], HRESULT, "GetAt", (["in"], wintypes.DWORD, "iProp"), (["out"], POINTER(PROPERTYKEY), "pkey")),
        COMMETHOD([], HRESULT, "GetValue", (["in"], POINTER(PROPERTYKEY), "key"), (["out"], POINTER(PROPVARIANT), "pv")),
        COMMETHOD([], HRESULT, "SetValue", (["in"], POINTER(PROPERTYKEY), "key"), (["in"], POINTER(PROPVARIANT), "propvar")),
        COMMETHOD([], HRESULT, "Commit"),
    ]


class IMMDevice(IUnknown):
    _iid_ = GUID("{d666063f-1587-4e43-81f1-b948e807363f}")
    _methods_ = [
        COMMETHOD([], HRESULT, "Activate", (["in"], POINTER(GUID), "iid"), (["in"], wintypes.DWORD, "dwClsCtx"), (["in"], POINTER(PROPVARIANT), "pActivationParams"), (["out"], POINTER(c_void_p), "ppInterface")),
        COMMETHOD([], HRESULT, "OpenPropertyStore", (["in"], wintypes.DWORD, "stgmAccess"), (["out"], POINTER(POINTER(IPropertyStore)), "ppProperties")),
        COMMETHOD([], HRESULT, "GetId", (["out"], POINTER(c_wchar_p), "ppstrId")),
        COMMETHOD([], HRESULT, "GetState", (["out"], POINTER(wintypes.DWORD), "pdwState")),
    ]


class IMMDeviceCollection(IUnknown):
    _iid_ = GUID("{0bd7a1be-7a1a-44db-8397-cc5392387b5e}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(wintypes.UINT), "pcDevices")),
        COMMETHOD([], HRESULT, "Item", (["in"], wintypes.UINT, "nDevice"), (["out"], POINTER(POINTER(IMMDevice)), "ppDevice")),
    ]


class IMMDeviceEnumerator(IUnknown):
    _iid_ = GUID("{a95664d2-9614-4f35-a746-de8db63617e6}")
    _methods_ = [
        COMMETHOD([], HRESULT, "EnumAudioEndpoints", (["in"], c_int, "dataFlow"), (["in"], wintypes.DWORD, "dwStateMask"), (["out"], POINTER(POINTER(IMMDeviceCollection)), "ppDevices")),
        COMMETHOD([], HRESULT, "GetDefaultAudioEndpoint", (["in"], c_int, "dataFlow"), (["in"], c_int, "role"), (["out"], POINTER(POINTER(IMMDevice)), "ppEndpoint")),
        COMMETHOD([], HRESULT, "GetDevice", (["in"], c_wchar_p, "pwstrId"), (["out"], POINTER(POINTER(IMMDevice)), "ppDevice")),
        COMMETHOD([], HRESULT, "RegisterEndpointNotificationCallback", (["in"], c_void_p, "pClient")),
        COMMETHOD([], HRESULT, "UnregisterEndpointNotificationCallback", (["in"], c_void_p, "pClient")),
    ]


class IPolicyConfig(IUnknown):
    _iid_ = GUID("{f8679f50-850a-41cf-9c72-430f290290c8}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetMixFormat"),
        COMMETHOD([], HRESULT, "GetDeviceFormat"),
        COMMETHOD([], HRESULT, "ResetDeviceFormat"),
        COMMETHOD([], HRESULT, "SetDeviceFormat"),
        COMMETHOD([], HRESULT, "GetProcessingPeriod"),
        COMMETHOD([], HRESULT, "SetProcessingPeriod"),
        COMMETHOD([], HRESULT, "GetShareMode"),
        COMMETHOD([], HRESULT, "SetShareMode"),
        COMMETHOD([], HRESULT, "GetPropertyValue"),
        COMMETHOD([], HRESULT, "SetPropertyValue"),
        COMMETHOD([], HRESULT, "SetDefaultEndpoint", (["in"], c_wchar_p, "wszDeviceId"), (["in"], c_int, "role")),
        COMMETHOD([], HRESULT, "SetEndpointVisibility"),
    ]


CLSID_MMDEVICE_ENUMERATOR = GUID("{bcde0395-e52f-467c-8e3d-c4579291692e}")
CLSID_POLICY_CONFIG_CLIENT = GUID("{870af99c-171d-4f9e-af0d-e63df40c2bc9}")
PKEY_DEVICE_FRIENDLY_NAME = PROPERTYKEY(GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 14)


@dataclass(frozen=True)
class AudioDevice:
    id: str
    name: str


def _propvariant_clear(prop):
    windll.ole32.PropVariantClear(byref(prop))


def _com_initialized():
    class ComContext:
        def __enter__(self):
            comtypes.CoInitialize()
            return self

        def __exit__(self, exc_type, exc, tb):
            comtypes.CoUninitialize()

    return ComContext()


class AudioManager:
    def _enumerator(self):
        return comtypes.CoCreateInstance(CLSID_MMDEVICE_ENUMERATOR, IMMDeviceEnumerator, CLSCTX_ALL)

    def list_output_devices(self) -> list[AudioDevice]:
        with _com_initialized():
            enumerator = self._enumerator()
            collection = enumerator.EnumAudioEndpoints(E_RENDER, DEVICE_STATE_ACTIVE)
            count = collection.GetCount()
            devices = []
            for index in range(count):
                device = collection.Item(index)
                device_id = device.GetId()
                name = self._friendly_name(device) or device_id
                devices.append(AudioDevice(device_id, name))
            return devices

    def get_current_default_id(self) -> Optional[str]:
        with _com_initialized():
            try:
                device = self._enumerator().GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
                return device.GetId()
            except Exception:
                return None

    def set_default(self, device_id: str) -> None:
        with _com_initialized():
            device = self._enumerator().GetDevice(device_id)
            state = device.GetState()
            if not state & DEVICE_STATE_ACTIVE:
                raise RuntimeError("Audio device is not active")
            policy = comtypes.CoCreateInstance(CLSID_POLICY_CONFIG_CLIENT, IPolicyConfig, CLSCTX_ALL)
            errors = []
            for role in (E_CONSOLE, E_MULTIMEDIA, E_COMMUNICATIONS):
                try:
                    policy.SetDefaultEndpoint(device_id, role)
                except Exception as exc:
                    errors.append(exc)
            if len(errors) == 3:
                raise RuntimeError(f"Failed to set default audio endpoint: {errors[0]}")

    def _friendly_name(self, device) -> Optional[str]:
        store = device.OpenPropertyStore(STGM_READ)
        prop = None
        try:
            prop = store.GetValue(byref(PKEY_DEVICE_FRIENDLY_NAME))
            if prop.vt == VT_LPWSTR and prop.pwszVal:
                return prop.pwszVal
            return None
        finally:
            if prop is not None:
                _propvariant_clear(prop)


def default_config() -> dict:
    return {
        "version": CONFIG_VERSION,
        "audio_toggle": {
            "hotkey": DEFAULT_HOTKEY,
            "device_ids": [],
            "notifications_enabled": True,
        },
    }


def log_message(message: str) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = _datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with LOG_PATH.open("a", encoding="utf-8") as log:
            log.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


def log_exception(context: str, exc: BaseException) -> None:
    log_message(f"{context}: {exc}\n{traceback.format_exc()}")


def show_error_message(title: str, message: str) -> None:
    try:
        windll.user32.MessageBoxW(None, message, title, 0x00000010)
    except Exception:
        pass


def config_bool(value, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in ("1", "true", "yes", "on"):
            return True
        if lowered in ("0", "false", "no", "off"):
            return False
    return default


def normalize_config(config: dict) -> dict:
    normalized = default_config()
    source_toggle = config.get("audio_toggle", {}) if isinstance(config, dict) else {}
    if not isinstance(source_toggle, dict):
        source_toggle = {}

    toggle = normalized["audio_toggle"]
    toggle["hotkey"] = str(source_toggle.get("hotkey") or DEFAULT_HOTKEY).strip().lower()
    device_ids = []
    seen_device_ids = set()
    for item in source_toggle.get("device_ids", []):
        device_id = str(item).strip()
        if device_id and device_id not in seen_device_ids:
            device_ids.append(device_id)
            seen_device_ids.add(device_id)
    toggle["device_ids"] = device_ids
    toggle["notifications_enabled"] = config_bool(source_toggle.get("notifications_enabled"), True)
    return normalized


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return default_config()
    try:
        loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        normalized = normalize_config(loaded)
        if loaded != normalized:
            save_config(normalized)
        return normalized
    except Exception:
        backup = CONFIG_PATH.with_suffix(".broken.json")
        try:
            CONFIG_PATH.replace(backup)
        except Exception:
            pass
        return default_config()


def save_config(config: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(normalize_config(config), indent=2), encoding="utf-8")


def remove_local_config() -> None:
    local_app_data = Path(os.environ.get("LOCALAPPDATA", Path.home())).resolve()
    config_dir = CONFIG_DIR.resolve()
    if config_dir == local_app_data or local_app_data not in config_dir.parents:
        raise RuntimeError(f"Refusing to remove unexpected config path: {CONFIG_DIR}")
    if CONFIG_DIR.exists():
        shutil.rmtree(CONFIG_DIR)


def is_configured(config: dict) -> bool:
    return len(normalize_config(config)["audio_toggle"].get("device_ids", [])) >= 2


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def acquire_daemon_mutex():
    handle = windll.kernel32.CreateMutexW(None, False, INSTANCE_MUTEX_NAME)
    if not handle:
        raise ctypes.WinError()
    if windll.kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        windll.kernel32.CloseHandle(handle)
        return None
    return handle


def quote_arg(value: str) -> str:
    return f'"{value}"'


def startup_command() -> str:
    if is_frozen():
        return f'{quote_arg(sys.executable)} daemon'
    launcher = Path(sys.argv[0]).resolve()
    return f'{quote_arg(sys.executable)} {quote_arg(str(launcher))} daemon'


def set_startup_enabled(enabled: bool) -> None:
    with OpenKey(HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_SET_VALUE) as key:
        if enabled:
            SetValueEx(key, STARTUP_RUN_NAME, 0, REG_SZ, startup_command())
        else:
            try:
                DeleteValue(key, STARTUP_RUN_NAME)
            except FileNotFoundError:
                pass


def get_startup_enabled() -> bool:
    try:
        with OpenKey(HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as key:
            value, _ = QueryValueEx(key, STARTUP_RUN_NAME)
            return bool(value)
    except FileNotFoundError:
        return False


def open_config_folder() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    os.startfile(CONFIG_DIR)


def device_map(devices: list[AudioDevice]) -> dict[str, AudioDevice]:
    return {device.id: device for device in devices}


def select_device(identifier: str, devices: list[AudioDevice]) -> AudioDevice:
    if identifier.isdigit():
        index = int(identifier)
        if 1 <= index <= len(devices):
            return devices[index - 1]
        if 0 <= index < len(devices):
            return devices[index]
    for device in devices:
        if device.id == identifier or device.name.lower() == identifier.lower():
            return device
    raise ValueError(f"No output device matches '{identifier}'")


def toggle_device(config: dict, audio: AudioManager, notify: Callable[[str, str], None]) -> AudioDevice:
    toggle = normalize_config(config)["audio_toggle"]
    selected_ids = toggle["device_ids"]
    if len(selected_ids) < 2:
        raise RuntimeError("At least two audio output devices must be selected in Configure.")

    available = device_map(audio.list_output_devices())
    selected_available = [available[device_id] for device_id in selected_ids if device_id in available]
    missing_ids = [device_id for device_id in selected_ids if device_id not in available]

    if missing_ids and toggle["notifications_enabled"]:
        notify(APP_NAME, f"Skipped {len(missing_ids)} missing audio device(s).")

    if len(selected_available) < 1:
        raise RuntimeError("None of the selected audio output devices are currently available.")

    current_id = audio.get_current_default_id()
    selected_available_ids = [device.id for device in selected_available]
    if current_id in selected_available_ids:
        next_index = (selected_available_ids.index(current_id) + 1) % len(selected_available_ids)
    else:
        next_index = 0

    next_device = selected_available[next_index]
    audio.set_default(next_device.id)
    if toggle["notifications_enabled"]:
        notify("Audio output", next_device.name)
    return next_device


def create_icon_image():
    if Image is None:
        return None
    image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 18, 30, 46), radius=4, fill=(36, 99, 235, 255))
    draw.polygon([(30, 24), (48, 12), (48, 52), (30, 40)], fill=(36, 99, 235, 255))
    draw.arc((38, 20, 58, 44), start=-45, end=45, fill=(20, 184, 166, 255), width=5)
    draw.arc((32, 12, 68, 52), start=-42, end=42, fill=(20, 184, 166, 180), width=4)
    return image


MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
HOTKEY_ID = 0xA001


VK_ALIASES = {
    "backspace": 0x08,
    "tab": 0x09,
    "enter": 0x0D,
    "return": 0x0D,
    "esc": 0x1B,
    "escape": 0x1B,
    "space": 0x20,
    "pageup": 0x21,
    "pagedown": 0x22,
    "end": 0x23,
    "home": 0x24,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "insert": 0x2D,
    "delete": 0x2E,
}


def parse_hotkey(hotkey: str) -> tuple[int, int]:
    modifiers = 0
    key = None
    parts = [part.strip().lower() for part in hotkey.replace(" ", "").split("+") if part.strip()]
    for part in parts:
        if part in ("ctrl", "control"):
            modifiers |= MOD_CONTROL
        elif part == "alt":
            modifiers |= MOD_ALT
        elif part == "shift":
            modifiers |= MOD_SHIFT
        elif part in ("win", "windows", "meta"):
            modifiers |= MOD_WIN
        elif len(part) == 1 and part.isalpha():
            key = ord(part.upper())
        elif len(part) == 1 and part.isdigit():
            key = ord(part)
        elif part.startswith("f") and part[1:].isdigit() and 1 <= int(part[1:]) <= 24:
            key = 0x70 + int(part[1:]) - 1
        elif part in VK_ALIASES:
            key = VK_ALIASES[part]
        else:
            raise ValueError(f"Unsupported hotkey token: {part}")
    if key is None:
        raise ValueError("Hotkey must include a final key, for example ctrl+alt+a.")
    return modifiers, key


class HotkeyManager:
    def __init__(self, callback: Callable[[], None], error_callback: Callable[[str], None]):
        self.callback = callback
        self.error_callback = error_callback
        self.thread: Optional[threading.Thread] = None
        self.thread_id = 0
        self.lock = threading.Lock()

    def start(self, hotkey: str) -> None:
        self.stop()
        modifiers, vk = parse_hotkey(hotkey)
        with self.lock:
            self.thread = threading.Thread(target=self._run, args=(modifiers, vk, hotkey), daemon=True)
            self.thread.start()

    def stop(self) -> None:
        with self.lock:
            if self.thread and self.thread.is_alive() and self.thread_id:
                windll.user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)
                self.thread.join(timeout=2)
            self.thread = None
            self.thread_id = 0

    def _run(self, modifiers: int, vk: int, hotkey: str) -> None:
        self.thread_id = windll.kernel32.GetCurrentThreadId()
        if not windll.user32.RegisterHotKey(None, HOTKEY_ID, modifiers, vk):
            self.error_callback(f"Could not register hotkey '{hotkey}'. It may already be in use.")
            return
        msg = wintypes.MSG()
        try:
            while windll.user32.GetMessageW(byref(msg), None, 0, 0) != 0:
                if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
                    self.callback()
                windll.user32.TranslateMessage(byref(msg))
                windll.user32.DispatchMessageW(byref(msg))
        finally:
            windll.user32.UnregisterHotKey(None, HOTKEY_ID)


class ConfigureWindow:
    def __init__(self, app=None):
        self.app = app
        self.audio = app.audio if app else AudioManager()
        self.root = tk.Tk()
        self.root.title(f"{APP_NAME} Configure")
        self.root.geometry("620x500")
        self.root.minsize(560, 420)
        self.root.configure(bg="#f6f7fb")
        self.root.protocol("WM_DELETE_WINDOW", self.cancel)
        self.root.bind("<Control-s>", lambda _event: self.save())
        self.root.bind("<Escape>", lambda _event: self.cancel())
        self.config = load_config()
        self.device_vars: dict[str, tk.BooleanVar] = {}
        self.hotkey_var = tk.StringVar(value=self.config["audio_toggle"]["hotkey"])
        self.startup_var = tk.BooleanVar(value=get_startup_enabled())
        self.notifications_var = tk.BooleanVar(value=self.config["audio_toggle"]["notifications_enabled"])
        self.status_var = tk.StringVar(value="")
        self._build()
        self._load_devices()
        self.root.after(150, self.bring_to_front)

    def _build(self):
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Root.TFrame", background="#f6f7fb")
        style.configure("Header.TLabel", background="#f6f7fb", font=("Segoe UI", 15, "bold"))
        style.configure("Subtle.TLabel", background="#f6f7fb", foreground="#475569")
        style.configure("Panel.TLabelframe", background="#f6f7fb")
        style.configure("Panel.TLabelframe.Label", background="#f6f7fb", font=("Segoe UI", 10, "bold"))

        outer = ttk.Frame(self.root, padding=18, style="Root.TFrame")
        outer.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(outer, style="Root.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        title_block = ttk.Frame(header, style="Root.TFrame")
        title_block.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(title_block, text=APP_NAME, style="Header.TLabel").pack(anchor=tk.W)
        ttk.Label(title_block, text="Choose the output devices included in the hotkey toggle loop.", style="Subtle.TLabel").pack(anchor=tk.W, pady=(2, 0))

        devices_frame = ttk.LabelFrame(outer, text="Output devices", padding=10, style="Panel.TLabelframe")
        devices_frame.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        devices_frame.grid_rowconfigure(0, weight=1)
        devices_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(devices_frame, highlightthickness=0, bg="#ffffff", height=190)
        scrollbar = ttk.Scrollbar(devices_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.device_list = ttk.Frame(self.canvas, padding=8)
        self.device_list.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.device_window = self.canvas.create_window((0, 0), window=self.device_list, anchor="nw")
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfigure(self.device_window, width=event.width))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        settings = ttk.Frame(outer, style="Root.TFrame")
        settings.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(settings, text="Hotkey", style="Subtle.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=4)
        ttk.Entry(settings, textvariable=self.hotkey_var, width=28).grid(row=0, column=1, sticky=tk.W, pady=4)
        ttk.Checkbutton(settings, text="Start with Windows", variable=self.startup_var).grid(row=1, column=1, sticky=tk.W, pady=4)
        ttk.Checkbutton(settings, text="Enable notifications", variable=self.notifications_var).grid(row=2, column=1, sticky=tk.W, pady=4)

        footer = ttk.Frame(outer, style="Root.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        footer.grid_columnconfigure(2, weight=1)
        ttk.Button(footer, text="Save", command=self.save).grid(row=0, column=0, sticky=tk.W)
        ttk.Button(footer, text="Cancel", command=self.cancel).grid(row=0, column=1, sticky=tk.W, padx=(8, 0))
        ttk.Label(footer, textvariable=self.status_var, style="Subtle.TLabel").grid(row=0, column=2, sticky="ew", padx=(16, 0))
        ttk.Button(footer, text="Uninstall...", command=self.uninstall).grid(row=0, column=3, sticky=tk.E)

    def _load_devices(self):
        selected = set(self.config["audio_toggle"]["device_ids"])
        try:
            devices = self.audio.list_output_devices()
        except Exception as exc:
            self.status_var.set(f"Could not enumerate output devices: {exc}")
            return
        if not devices:
            self.status_var.set("No active output devices found.")
            return
        for device in devices:
            var = tk.BooleanVar(value=device.id in selected)
            self.device_vars[device.id] = var
            ttk.Checkbutton(self.device_list, text=device.name, variable=var).pack(anchor=tk.W, fill=tk.X, pady=3)
        missing = [device_id for device_id in selected if device_id not in self.device_vars]
        if missing:
            self.status_var.set(f"{len(missing)} configured device(s) are not currently available.")

    def save(self):
        selected_ids = [device_id for device_id, var in self.device_vars.items() if var.get()]
        if len(selected_ids) < 2:
            messagebox.showerror(APP_NAME, "Select at least two output devices.")
            return
        hotkey = self.hotkey_var.get().strip().lower()
        try:
            parse_hotkey(hotkey)
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.config["audio_toggle"].update(
            {
                "hotkey": hotkey,
                "device_ids": selected_ids,
                "notifications_enabled": self.notifications_var.get(),
            }
        )
        try:
            save_config(self.config)
            set_startup_enabled(self.startup_var.get())
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not save settings: {exc}")
            return
        if self.app:
            self.app.reload()
        self.root.destroy()

    def cancel(self):
        self.root.destroy()

    def uninstall(self):
        confirmed = messagebox.askyesno(
            APP_NAME,
            "Remove the Windows startup entry and delete saved configuration/log files?\n\n"
            "The portable executable will remain where you placed it.",
        )
        if not confirmed:
            return
        try:
            set_startup_enabled(False)
            remove_local_config()
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not complete uninstall cleanup: {exc}")
            return
        messagebox.showinfo(APP_NAME, "Startup entry and saved local data were removed.")
        self.root.destroy()
        if self.app:
            self.app.quit()

    def bring_to_front(self):
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(300, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    def run(self):
        self.root.mainloop()


class TrayApp:
    def __init__(self):
        if pystray is None:
            raise RuntimeError("pystray and Pillow are required for tray mode. Install requirements.txt.")
        self.audio = AudioManager()
        self.config = load_config()
        self.config_window_open = False
        self.icon = pystray.Icon(APP_NAME, icon=create_icon_image(), title=APP_NAME)
        self.hotkey = HotkeyManager(self.toggle_from_hotkey, self.hotkey_error)
        self.icon.menu = self.build_menu()

    def notify(self, title: str, message: str) -> None:
        if not self.config["audio_toggle"].get("notifications_enabled", True):
            return
        try:
            self.icon.notify(message, title)
        except Exception:
            log_message(f"Notification failed: {title}: {message}")

    def build_menu(self):
        try:
            devices = self.audio.list_output_devices()
        except Exception:
            devices = []

        direct_items = [
            pystray.MenuItem(device.name, self.set_device_action(device.id))
            for device in devices
        ]
        if not direct_items:
            direct_items = [pystray.MenuItem("No output devices found", None, enabled=False)]

        return pystray.Menu(
            pystray.MenuItem("Toggle audio output", lambda _icon, _item: self.toggle()),
            pystray.MenuItem("Switch directly", pystray.Menu(*direct_items)),
            pystray.MenuItem("Configure", lambda _icon, _item: self.configure()),
            pystray.MenuItem("Start with Windows", self.toggle_startup, checked=lambda _item: get_startup_enabled()),
            pystray.MenuItem("Open config folder", lambda _icon, _item: open_config_folder()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", lambda _icon, _item: self.quit()),
        )

    def set_device_action(self, device_id: str):
        def action(_icon, _item):
            self.set_device(device_id)

        return action

    def run(self):
        self.register_hotkey()
        self.icon.run(setup=self.setup)

    def setup(self, _icon):
        try:
            _icon.visible = True
        except Exception as exc:
            log_exception("Tray icon could not be shown", exc)
            show_error_message(APP_NAME, f"Could not show the tray icon.\n\n{exc}\n\nSee:\n{LOG_PATH}")
        if not is_configured(self.config):
            self.configure()

    def register_hotkey(self):
        try:
            self.hotkey.start(self.config["audio_toggle"].get("hotkey") or DEFAULT_HOTKEY)
        except Exception as exc:
            self.hotkey_error(str(exc))

    def reload(self):
        self.config = load_config()
        self.icon.menu = self.build_menu()
        self.register_hotkey()

    def toggle_from_hotkey(self):
        self.toggle()

    def toggle(self):
        try:
            self.config = load_config()
            toggle_device(self.config, self.audio, self.notify)
        except Exception as exc:
            log_exception("Toggle failed", exc)
            self.notify(APP_NAME, str(exc))

    def set_device(self, device_id: str):
        try:
            devices = device_map(self.audio.list_output_devices())
            if device_id not in devices:
                raise RuntimeError("That output device is not currently available.")
            self.audio.set_default(device_id)
            self.config = load_config()
            if self.config["audio_toggle"].get("notifications_enabled", True):
                self.notify("Audio output", devices[device_id].name)
        except Exception as exc:
            log_exception("Set device failed", exc)
            self.notify(APP_NAME, str(exc))

    def configure(self):
        if self.config_window_open:
            return
        self.config_window_open = True

        def run_window():
            try:
                ConfigureWindow(self).run()
            except Exception as exc:
                log_exception("Configure window failed", exc)
                show_error_message(APP_NAME, f"Could not open Configure.\n\n{exc}\n\nSee:\n{LOG_PATH}")
            finally:
                self.config_window_open = False

        threading.Thread(target=run_window, daemon=True).start()

    def toggle_startup(self, _icon, _item):
        try:
            set_startup_enabled(not get_startup_enabled())
        finally:
            self.icon.update_menu()

    def hotkey_error(self, message: str):
        log_message(message)
        self.notify(APP_NAME, message)

    def quit(self):
        self.hotkey.stop()
        self.icon.stop()


def cli_notify(title: str, message: str) -> None:
    print(f"{title}: {message}")


def command_list(_args) -> int:
    devices = AudioManager().list_output_devices()
    current_id = AudioManager().get_current_default_id()
    for index, device in enumerate(devices, start=1):
        marker = "*" if device.id == current_id else " "
        print(f"{marker} {index}. {device.name}")
        print(f"    {device.id}")
    return 0


def command_set(args) -> int:
    audio = AudioManager()
    devices = audio.list_output_devices()
    device = select_device(args.device, devices)
    audio.set_default(device.id)
    config = load_config()
    if config["audio_toggle"].get("notifications_enabled", True):
        cli_notify("Audio output", device.name)
    return 0


def command_toggle(_args) -> int:
    config = load_config()
    toggle_device(config, AudioManager(), cli_notify)
    return 0


def command_configure(_args) -> int:
    ConfigureWindow().run()
    return 0


def command_daemon(_args) -> int:
    mutex = acquire_daemon_mutex()
    if mutex is None:
        message = "Windows Audio Switch is already running. Check the system tray overflow area."
        log_message(message)
        show_error_message(APP_NAME, message)
        return 0
    try:
        TrayApp().run()
    finally:
        windll.kernel32.CloseHandle(mutex)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="audio-switch", description=f"{APP_NAME} tray utility")
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("list", help="List active audio output devices").set_defaults(func=command_list)
    set_parser = subparsers.add_parser("set", help="Set the default output device by id or list index")
    set_parser.add_argument("device", help="Device id, exact friendly name, or list index")
    set_parser.set_defaults(func=command_set)
    subparsers.add_parser("toggle", help="Cycle to the next configured output device").set_defaults(func=command_toggle)
    subparsers.add_parser("configure", help="Open the configuration window").set_defaults(func=command_configure)
    subparsers.add_parser("daemon", help="Run quietly from the system tray").set_defaults(func=command_daemon)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "command", None):
        args = parser.parse_args(["daemon"])
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        log_exception("Command failed", exc)
        print(f"Error: {exc}", file=sys.stderr)
        if getattr(args, "command", None) in ("daemon", "configure"):
            show_error_message(APP_NAME, f"{exc}\n\nSee:\n{LOG_PATH}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
