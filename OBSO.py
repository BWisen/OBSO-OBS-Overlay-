import os
import sys
import json
import threading
import time
import ctypes
import re
import tkinter as tk
from functools import partial

# Dependency check
missing = []
try:
    import websocket
except ImportError:
    missing.append("obs-websocket-py")
try:
    import win32api
    import win32con
    import win32gui
    import win32process
except ImportError:
    missing.append("pywin32")
try:
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except ImportError:
    missing.append("pycaw")
try:
    import pystray
except ImportError:
    missing.append("pystray")
try:
    import keyboard
except ImportError:
    missing.append("keyboard")

if missing:
    from tkinter import messagebox
    msg = (
        "Missing required modules:\n\n" + "\n".join(missing) +
        "\n\nInstall with:\n" + f"pip install {' '.join(missing)}"
    )
    messagebox.showerror("Missing Dependencies", msg)
    sys.exit(1)

import websocket
import pythoncom
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import POINTER, cast

import win32gui
import win32process
import keyboard

CONFIG_PATH = "config.json"
DEFAULT_CONFIG = {
    "obs_host": "localhost",
    "obs_port": 4455,
    "obs_password": "",
    "hotkey": "f1+shift",
    "overlay_visible": True
}

if not os.path.exists(CONFIG_PATH):
    with open(CONFIG_PATH, "w") as f:
        json.dump(DEFAULT_CONFIG, f, indent=4)

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

def save_config():
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=4)

class OBSConnector:
    def __init__(self, on_status_update, on_event):
        self.ws = None
        self.thread = None
        self.connected = False
        self.on_status_update = on_status_update
        self.on_event = on_event
        self.should_run = True
        self.request_id = 0
        self.connect()

    def connect(self):
        if self.thread and self.thread.is_alive():
            self.should_run = False
            self.thread.join()
            self.should_run = True

        def run():
            while self.should_run:
                try:
                    self.ws = websocket.create_connection(
                        f"ws://{config['obs_host']}:{config['obs_port']}"
                    )
                    # Identify connection (OBS 5.0+ protocol)
                    self.ws.send(json.dumps({"op": 1, "d": {"rpcVersion": 1}}))
                    if config["obs_password"]:
                        self.ws.send(json.dumps({
                            "op": 3,
                            "d": {"rpcVersion": 1, "authentication": config["obs_password"]}
                        }))
                    self.connected = True
                    self.on_status_update("Connected")

                    self.subscribe_events([
                        "RecordStateChanged",
                        "StreamStateChanged",
                        "ReplayBufferStateChanged"
                    ])

                    while self.connected and self.should_run:
                        message = self.ws.recv()
                        if not message:
                            continue
                        data = json.loads(message)
                        if data.get("op") == 5 and "d" in data:
                            event_type = data["d"].get("eventType", "")
                            event_data = data["d"].get("eventData", {})
                            self.on_event(event_type, event_data)
                except Exception:
                    self.connected = False
                    self.on_status_update("Reconnecting...")
                    time.sleep(5)

        self.thread = threading.Thread(target=run, daemon=True)
        self.thread.start()

    def subscribe_events(self, events):
        try:
            self.request_id += 1
            self.ws.send(json.dumps({
                "op": 6,
                "d": {
                    "requestType": "Subscribe",
                    "requestId": f"sub_{self.request_id}",
                    "eventSubscriptions": events
                }
            }))
        except Exception as e:
            print("Failed to subscribe events:", e)

    def send(self, request_type):
        if self.connected:
            try:
                self.ws.send(json.dumps({"op": 6, "d": {
                    "requestType": request_type, "requestId": "1"
                }}))
            except:
                self.connected = False
                self.on_status_update("Disconnected")

    def stop(self):
        self.should_run = False
        if self.ws:
            try:
                self.ws.close()
            except:
                pass

class IndicatorButton(tk.Frame):
    def __init__(self, parent, text, command, width_chars=6):
        super().__init__(parent, bg="#3a3a3a")
        self.button = tk.Button(self, text=text, command=command,
                                bg="#444", fg="white", width=width_chars, height=2,
                                font=("Segoe UI", 10))
        self.button.pack(side=tk.LEFT)
        self.indicator = tk.Canvas(self, width=14, height=14, bg=("#444444"), highlightthickness=0)
        self.indicator_circle = self.indicator.create_oval(2, 2, 12, 12, fill="#d3d3d3", outline="white", width=1.5)
        self.indicator.place(in_=self.button, relx=1, x=-10, y=0, anchor='ne')

    def set_color(self, color):
        self.indicator.itemconfig(self.indicator_circle, fill=color)

class Overlay(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OBS Overlay")
        self.attributes("-topmost", True)
        self.overrideredirect(True)

        self.screen_width = self.winfo_screenwidth()
        self.geometry(f"{self.screen_width}x80+0+0")
        self.configure(bg="#2e2e2e")

        self.opacity = 0.95
        self.visible = config.get("overlay_visible", True)
        self.settings_visible = False

        self.recording_active = False
        self.streaming_active = False
        self.replay_active = False

        self.obs = OBSConnector(self.update_status, self.handle_obs_event)

        self.current_hotkey_keys = self.parse_hotkey_to_set(config["hotkey"])
        self.listening_for_hotkey = False
        self.new_hotkey_keys = set()

        self.setup_ui()

        # Wait a bit then set overlay click-through based on visibility
        self.after(100, self.make_overlay_clickable)

        self.register_global_hotkey(config["hotkey"])

        threading.Thread(target=self.setup_tray_icon, daemon=True).start()

        self.update_media_loop()

    def setup_ui(self):
        self.bar = tk.Frame(self, bg="#2e2e2e")
        self.bar.pack(fill=tk.BOTH, expand=True)

        self.logo_text = tk.Label(self.bar, text="OBSO", fg="white", bg="#2e2e2e",
                                  font=("Segoe UI", 18, "bold"))
        self.logo_text.place(x=15, y=15)

        self.status = tk.Label(self.bar, text="Connecting...", fg="white", bg="#2e2e2e", font=("Segoe UI", 10))
        self.status.place(x=10, y=60, anchor="w")

        margin_left = 100
        margin_right = 60
        available_width = self.screen_width - margin_left - margin_right
        obs_width = int(available_width * 0.45)
        media_width = available_width - obs_width - 10

        obs_x = margin_left
        media_x = obs_x + obs_width + 10

        self.controls = self.outlined_section(x=obs_x, y=20, label="OBS Controls", width=obs_width, height=50)

        button_info = [
            ("Record", "ToggleRecord"),
            ("Stream", "ToggleStream"),
            ("Replay", "ToggleReplayBuffer"),
            ("Save Replay", "SaveReplayBuffer")
        ]
        button_count = len(button_info)
        padding_per_button = 12
        available_btn_width = obs_width - (button_count * padding_per_button)
        btn_char_width = max(int(available_btn_width / button_count / 8), 6)

        self.obs_buttons = {}
        self.obs_indicators = {}

        for label, cmd in button_info:
            if label in ("Record", "Stream", "Replay"):
                btn = IndicatorButton(self.controls, text=label, command=partial(self.send_obs_command, cmd), width_chars=btn_char_width)
                btn.pack(side=tk.LEFT, padx=6, pady=4)
                self.obs_buttons[label] = btn.button
                self.obs_indicators[label] = btn
            else:
                btn = tk.Button(self.controls, text=label, command=partial(self.send_obs_command, cmd),
                                bg="#444", fg="white", width=btn_char_width, height=2, font=("Segoe UI", 10))
                btn.pack(side=tk.LEFT, padx=6, pady=4)
                self.obs_buttons[label] = btn

        self.media = self.outlined_section(x=media_x, y=20, label="Media Controls", width=media_width, height=50)
        self.track_label = tk.Label(self.media, text="Title - Artist", fg="white", bg="#3a3a3a",
                                    font=("Segoe UI", 11), anchor="w")
        self.track_label.pack(side=tk.LEFT, padx=12, fill=tk.X, expand=True)
        for label in ["⏮", "⏯", "⏭", "🔉", "🔊"]:
            b = tk.Button(self.media, text=label, command=partial(self.media_control, label),
                          bg="#444", fg="white", width=4, height=2, font=("Segoe UI", 11))
            b.pack(side=tk.LEFT, padx=5, pady=4)

        self.settings_btn = tk.Button(self.bar, text="⚙", command=self.toggle_settings,
                                      bg="#2e2e2e", fg="white", font=("Segoe UI", 12), borderwidth=0)
        self.settings_btn.place(x=self.screen_width - 40, y=30)

        # Settings window setup
        self.settings_window = tk.Toplevel(self)
        self.settings_window.withdraw()
        self.settings_window.overrideredirect(True)
        self.settings_window.attributes("-topmost", True)
        self.settings_window.configure(bg="#3a3a3a")
        self.settings_window.geometry("330x180+{}+{}".format(self.screen_width - 350, 80))
        self.settings_window.protocol("WM_DELETE_WINDOW", self.toggle_settings)

        self.obs_host = self.setting_field(self.settings_window, "OBS Host", config["obs_host"])
        self.obs_port = self.setting_field(self.settings_window, "Port", str(config["obs_port"]))
        self.obs_pwd = self.setting_field(self.settings_window, "Password", config["obs_password"], True)

        hotkey_frame = tk.Frame(self.settings_window, bg="#3a3a3a")
        hotkey_label = tk.Label(hotkey_frame, text="Hotkey:", fg="white", bg="#3a3a3a", width=10, anchor="w")
        hotkey_label.pack(side=tk.LEFT, padx=5)

        self.hotkey_button = tk.Button(hotkey_frame, text=config["hotkey"], bg="#555", fg="white", width=20,
                                       command=self.start_hotkey_capture)
        self.hotkey_button.pack(side=tk.LEFT, padx=5)

        hotkey_frame.pack(fill=tk.X, pady=2, padx=5)

        self.done_cancel_frame = tk.Frame(self.settings_window, bg="#3a3a3a")
        self.done_btn = tk.Button(self.done_cancel_frame, text="Done", command=self.finish_hotkey_capture,
                                  bg="#4caf50", fg="white", width=8)
        self.cancel_btn = tk.Button(self.done_cancel_frame, text="Cancel", command=self.cancel_hotkey_capture,
                                    bg="#f44336", fg="white", width=8)

        self.apply_btn = tk.Button(self.settings_window, text="Apply", command=self.apply_settings, bg="#555", fg="white")
        self.apply_btn.pack(pady=5, fill=tk.X, padx=10)

    def outlined_section(self, x, y, label, width, height=50):
        frame = tk.LabelFrame(self.bar, text=label, bg="#3a3a3a", fg="white",
                              font=("Segoe UI", 10), bd=2, relief=tk.GROOVE,
                              width=width, height=height)
        frame.place(x=x, y=y)
        frame.pack_propagate(False)
        return frame

    def setting_field(self, parent, label, value, password=False):
        frame = tk.Frame(parent, bg="#3a3a3a")
        lbl = tk.Label(frame, text=label, fg="white", bg="#3a3a3a", width=10, anchor="w")
        lbl.pack(side=tk.LEFT, padx=5)
        ent = tk.Entry(frame, bg="#444", fg="white", show="*" if password else "")
        ent.insert(0, value)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        frame.pack(fill=tk.X, pady=2, padx=5)
        return ent

    def toggle_settings(self):
        if self.settings_visible:
            self.settings_window.withdraw()
        else:
            x = self.screen_width - 350
            y = 80
            self.settings_window.geometry(f"330x180+{x}+{y}")
            self.settings_window.deiconify()
        self.settings_visible = not self.settings_visible

    def apply_settings(self):
        def apply():
            try:
                config["obs_host"] = self.obs_host.get()
                config["obs_port"] = int(self.obs_port.get())
                config["obs_password"] = self.obs_pwd.get()
                save_config()
                threading.Thread(target=self.obs.connect, daemon=True).start()
            except Exception as e:
                print("Error applying settings:", e)
            finally:
                self.after(0, lambda: self.apply_btn.config(state=tk.NORMAL))

        self.apply_btn.config(state=tk.DISABLED)
        threading.Thread(target=apply, daemon=True).start()

    def start_hotkey_capture(self):
        if self.listening_for_hotkey:
            return
        self.listening_for_hotkey = True
        self.new_hotkey_keys = set()
        self.hotkey_button.config(text="Press keys...")
        self.done_cancel_frame.pack(fill=tk.X, pady=5, padx=10)
        self.done_btn.pack(side=tk.LEFT, expand=True, padx=5)
        self.cancel_btn.pack(side=tk.LEFT, expand=True, padx=5)

        keyboard.hook(self.capture_hotkey_event)

    def capture_hotkey_event(self, event):
        if event.event_type == "down":
            key = event.name
            if key not in self.new_hotkey_keys:
                self.new_hotkey_keys.add(key)
            self.hotkey_button.config(text=self.normalize_hotkey_keys(self.new_hotkey_keys))

    def finish_hotkey_capture(self):
        if not self.new_hotkey_keys:
            self.cancel_hotkey_capture()
            return
        hotkey_str = self.normalize_hotkey_keys(self.new_hotkey_keys)
        self.on_hotkey_change(hotkey_str)
        self.hotkey_button.config(text=hotkey_str)
        self.stop_hotkey_capture()

    def cancel_hotkey_capture(self):
        self.hotkey_button.config(text=config["hotkey"])
        self.stop_hotkey_capture()

    def stop_hotkey_capture(self):
        self.listening_for_hotkey = False
        self.new_hotkey_keys = set()
        self.done_cancel_frame.pack_forget()
        keyboard.unhook(self.capture_hotkey_event)

    def on_hotkey_change(self, new_hotkey):
        config["hotkey"] = new_hotkey
        save_config()
        self.current_hotkey_keys = self.parse_hotkey_to_set(new_hotkey)
        self.register_global_hotkey(new_hotkey)

    def parse_hotkey_to_set(self, hotkey_str):
        return set(hotkey_str.lower().split("+"))

    def normalize_hotkey_keys(self, keys):
        key_map = {
            "shift_l": "shift",
            "shift_r": "shift",
            "ctrl_l": "ctrl",
            "ctrl_r": "ctrl",
            "alt_l": "alt",
            "alt_r": "alt",
            "cmd": "windows",
            "win": "windows",
        }
        normalized = set()
        for k in keys:
            lk = k.lower()
            normalized.add(key_map.get(lk, lk))
        return "+".join(sorted(normalized))

    def register_global_hotkey(self, hotkey_str):
        try:
            keyboard.unhook_all_hotkeys()
            keyboard.add_hotkey(hotkey_str, self.toggle_visibility)
            print(f"Registered global hotkey: {hotkey_str}")
        except Exception as e:
            print(f"Failed to register global hotkey '{hotkey_str}': {e}")

    def update_status(self, status):
        self.status.config(text=f"OBS: {status}")

    def send_obs_command(self, cmd):
        self.obs.send(cmd)

    def toggle_visibility(self):
        self.visible = not self.visible
        config["overlay_visible"] = self.visible
        save_config()
        self.attributes("-alpha", self.opacity if self.visible else 0)
        self.make_overlay_clickable()

    def media_control(self, action):
        key_map = {
            "⏯": win32con.VK_MEDIA_PLAY_PAUSE,
            "⏮": win32con.VK_MEDIA_PREV_TRACK,
            "⏭": win32con.VK_MEDIA_NEXT_TRACK,
        }
        if action in key_map:
            key = key_map[action]
            win32api.keybd_event(key, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif action == "🔉":
            self.set_volume(self.get_volume() - 0.05)
        elif action == "🔊":
            self.set_volume(self.get_volume() + 0.05)

    def get_volume(self):
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        return volume.GetMasterVolumeLevelScalar()

    def set_volume(self, level):
        level = max(0, min(1, level))
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level, None)

    def handle_obs_event(self, event_type, event_data):
        if event_type == "RecordStateChanged":
            active = event_data.get("outputActive", False)
            self.recording_active = active
            self.obs_indicators["Record"].set_color("#4caf50" if active else "#d3d3d3")
        elif event_type == "StreamStateChanged":
            active = event_data.get("outputActive", False)
            self.streaming_active = active
            self.obs_indicators["Stream"].set_color("#4caf50" if active else "#d3d3d3")
        elif event_type == "ReplayBufferStateChanged":
            active = event_data.get("outputActive", False)
            self.replay_active = active
            self.obs_indicators["Replay"].set_color("#4caf50" if active else "#d3d3d3")

    def make_overlay_clickable(self):
        hwnd = self.winfo_id()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if self.visible:
            style &= ~win32con.WS_EX_TRANSPARENT
        else:
            style |= win32con.WS_EX_TRANSPARENT
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
        self.attributes("-alpha", self.opacity if self.visible else 0)

    def get_media_title_from_window(self):
        media_processes = {
            "spotify.exe",
            "vlc.exe",
            "wmplayer.exe",
            "itunes.exe",
            "chrome.exe",
            "firefox.exe",
            "msedge.exe"
        }

        titles = []

        def enum_handler(hwnd, results):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    h_process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    exe_name = win32process.GetModuleFileNameEx(h_process, 0)
                    exe_name = os.path.basename(exe_name).lower()
                    win32api.CloseHandle(h_process)
                except Exception:
                    exe_name = ""

                if exe_name in media_processes:
                    title = win32gui.GetWindowText(hwnd)
                    if exe_name in {"chrome.exe", "firefox.exe", "msedge.exe"}:
                        if "youtube" in title.lower():
                            results.append(title)
                    else:
                        results.append(title)

        win32gui.EnumWindows(enum_handler, titles)

        for t in titles:
            if len(t) < 5:
                continue
            if re.match(r".+\s-\s.+", t):
                return t

        if titles:
            return titles[0]

        return "No media playing"

    def update_media_loop(self):
        title = self.get_media_title_from_window()
        self.track_label.config(text=title)
        self.after(2000, self.update_media_loop)

    def setup_tray_icon(self):
        import PIL.Image
        import PIL.ImageDraw

        def on_quit(icon, item):
            self.obs.stop()
            icon.stop()
            self.destroy()
            sys.exit(0)

        def toggle_overlay(icon, item):
            self.toggle_visibility()

        # Create simple black square icon with white O letter
        image = PIL.Image.new('RGB', (64, 64), color='black')
        draw = PIL.ImageDraw.Draw(image)
        draw.text((18, 16), "O", fill="white")

        menu = pystray.Menu(
            pystray.MenuItem("Toggle Overlay", toggle_overlay),
            pystray.MenuItem("Quit", on_quit)
        )
        icon = pystray.Icon("OBSO", image, "OBSO", menu)
        icon.run()

if __name__ == "__main__":
    app = Overlay()
    app.mainloop()
