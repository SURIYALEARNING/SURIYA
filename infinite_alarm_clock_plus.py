import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime as dt
import json
import os
import threading
import time
import winsound

# ---- Optional Windows lock/unlock support (pywin32) ----
PYWIN32_OK = False
try:
    import win32con, win32gui, win32api, win32ts
    PYWIN32_OK = True
except Exception:
    PYWIN32_OK = False

CONFIG_FILE = "alarms_v3.json"

DEFAULT_ALARMS = [
    {"label": "Strategy Planning",     "time": "09:30", "enabled": True},
    {"label": "Video Editing",         "time": "10:00", "enabled": True},
    {"label": "Social Media Posting",  "time": "14:00", "enabled": True},
]

DEFAULT_SETTINGS = {
    "alarms": DEFAULT_ALARMS,
    "default_ringtone": "",           # path to WAV, empty = beep fallback
    "pause_on_lock": True
}

def parse_hhmm(hhmm: str):
    s = hhmm.strip()
    if ":" not in s:
        raise ValueError("Use HH:MM (24h)")
    h_str, m_str = s.split(":", 1)
    if not (h_str.isdigit() and m_str.isdigit()):
        raise ValueError("Hour/Minute must be numbers")
    h, m = int(h_str), int(m_str)
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError("Hour 0–23, Minute 0–59")
    return h, m

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Back-compat: if file is a list, wrap into dict
            if isinstance(data, list):
                return {
                    "alarms": data,
                    "default_ringtone": "",
                    "pause_on_lock": True
                }
            # sanitize
            out = {"alarms": [], "default_ringtone": "", "pause_on_lock": True}
            out["default_ringtone"] = str(data.get("default_ringtone", "")).strip()
            out["pause_on_lock"] = bool(data.get("pause_on_lock", True))
            for item in data.get("alarms", []):
                out["alarms"].append({
                    "label": str(item.get("label", "")).strip(),
                    "time":  str(item.get("time", "")).strip(),
                    "enabled": bool(item.get("enabled", True))
                })
            if not out["alarms"]:
                out["alarms"] = DEFAULT_ALARMS
            return out
        except Exception:
            pass
    return DEFAULT_SETTINGS

def save_config(alarms_vars, default_ringtone, pause_on_lock):
    data = {
        "alarms": [],
        "default_ringtone": default_ringtone or "",
        "pause_on_lock": bool(pause_on_lock)
    }
    for rv in alarms_vars:
        data["alarms"].append({
            "label": rv["label_var"].get().strip(),
            "time":  rv["time_var"].get().strip(),
            "enabled": bool(rv["enabled_var"].get())
        })
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

class Beeper:
    """Fallback beeper (used when no WAV chosen)."""
    def __init__(self):
        self._stop = threading.Event()
        self._thread = None
    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
    def _run(self):
        while not self._stop.is_set():
            winsound.Beep(1000, 500)
            for _ in range(10):
                if self._stop.is_set(): break
                time.sleep(0.03)
    def stop(self):
        self._stop.set()

class SoundPlayer:
    """Plays looping WAV via winsound, or falls back to Beeper."""
    def __init__(self):
        self.beeper = Beeper()
        self._using_wav = False
    def play(self, wav_path: str | None):
        self.stop()
        if wav_path and os.path.exists(wav_path):
            try:
                winsound.PlaySound(
                    wav_path,
                    winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP
                )
                self._using_wav = True
                return
            except Exception:
                self._using_wav = False
        # fallback
        self._using_wav = False
        self.beeper.start()
    def stop(self):
        if self._using_wav:
            try:
                winsound.PlaySound(None, winsound.SND_PURGE)
            except Exception:
                pass
        self.beeper.stop()
        self._using_wav = False

# ---- Windows session watcher (lock/unlock) via pywin32 ----
class SessionWatcher(threading.Thread):
    WM_WTSSESSION_CHANGE = 0x02B1
    WTS_SESSION_LOCK     = 0x7
    WTS_SESSION_UNLOCK   = 0x8
    NOTIFY_FOR_THIS_SESSION = 0

    def __init__(self, on_lock, on_unlock):
        super().__init__(daemon=True)
        self.on_lock = on_lock
        self.on_unlock = on_unlock
        self.hinst = None
        self.classAtom = None
        self.hwnd = None

    def run(self):
        try:
            wc = win32gui.WNDCLASS()
            self.hinst = wc.hInstance = win32api.GetModuleHandle(None)
            wc.lpszClassName = "SessionWatcherHiddenWindow"
            wc.lpfnWndProc = self._wndproc
            self.classAtom = win32gui.RegisterClass(wc)
            self.hwnd = win32gui.CreateWindow(
                self.classAtom, "SessionWatcher", 0, 0, 0, 0, 0,
                0, 0, self.hinst, None
            )
            # Register for this session's notifications
            win32ts.WTSRegisterSessionNotification(self.hwnd, self.NOTIFY_FOR_THIS_SESSION)
            win32gui.PumpMessages()
        except Exception:
            # Silent exit; if this fails we simply won't pause on lock
            pass
        finally:
            try:
                if self.hwnd:
                    win32ts.WTSUnRegisterSessionNotification(self.hwnd)
                if self.classAtom:
                    win32gui.UnregisterClass(self.classAtom, self.hinst)
            except Exception:
                pass

    def _wndproc(self, hwnd, msg, wparam, lparam):
        if msg == self.WM_WTSSESSION_CHANGE:
            if wparam == self.WTS_SESSION_LOCK:
                try:
                    self.on_lock()
                except Exception:
                    pass
            elif wparam == self.WTS_SESSION_UNLOCK:
                try:
                    self.on_unlock()
                except Exception:
                    pass
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

class InfiniteAlarmApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Infinite Alarms (Day Starter) — WAV + Pause-on-Lock")
        self.root.resizable(False, False)

        self.rows_vars = []       # alarms UI rows
        self.armed = False
        self.fired_today = set()
        self.player = SoundPlayer()

        # lock/pause state
        self.paused = False
        self.pause_on_lock_var = tk.BooleanVar(value=True if PYWIN32_OK else False)
        self.default_ringtone_path = ""
        self.ringtone_var = tk.StringVar(value="Ringtone: Beep (default)")

        self._build_ui()
        self._load_existing()
        self._tick()

        # Session watcher (only if pywin32 present)
        if PYWIN32_OK:
            self.pause_on_lock_var.set(self.pause_on_lock_var.get())
            self._start_session_watcher()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        header = ttk.Label(self.root, text="Alarms for Today (HH:MM 24-hour) — unlimited rows", font=("Segoe UI", 12, "bold"))
        header.grid(row=0, column=0, columnspan=6, **pad)

        # Controls row
        btns = ttk.Frame(self.root)
        btns.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0,2))

        self.start_btn = ttk.Button(btns, text="Start All (Today)", command=self._start_all)
        self.stop_btn  = ttk.Button(btns, text="Stop All", command=self._stop_all)
        self.save_btn  = ttk.Button(btns, text="Save", command=self._save)
        self.add_btn   = ttk.Button(btns, text="+ Add", command=self._add_row)
        self.dup_btn   = ttk.Button(btns, text="Duplicate Selected", command=self._duplicate_selected)
        self.del_btn   = ttk.Button(btns, text="Delete Selected", command=self._delete_selected)

        self.start_btn.grid(row=0, column=0, padx=4)
        self.stop_btn.grid(row=0, column=1, padx=4)
        self.save_btn.grid(row=0, column=2, padx=4)
        self.add_btn.grid(row=0, column=3, padx=12)
        self.dup_btn.grid(row=0, column=4, padx=4)
        self.del_btn.grid(row=0, column=5, padx=4)

        # Ringtone + Pause-on-lock row
        tools = ttk.Frame(self.root)
        tools.grid(row=2, column=0, sticky="we", padx=8, pady=(0,6))
        tools.columnconfigure(1, weight=1)

        ttk.Label(tools, textvariable=self.ringtone_var).grid(row=0, column=0, sticky="w", padx=(0,10))
        ttk.Button(tools, text="Browse WAV…", command=self._pick_wav).grid(row=0, column=1, sticky="w")
        ttk.Button(tools, text="Clear", command=self._clear_wav).grid(row=0, column=2, sticky="w", padx=(8,0))

        pause_chk = ttk.Checkbutton(tools, text="Pause on Windows lock (resume + ring missed on unlock)", variable=self.pause_on_lock_var)
        pause_chk.grid(row=0, column=3, sticky="e", padx=(20,0))
        if not PYWIN32_OK:
            pause_chk.state(["disabled"])
            tip = ttk.Label(tools, foreground="#a00", text="(Install pywin32 to enable)")
            tip.grid(row=0, column=4, sticky="w", padx=(6,0))

        # Headers
        hdr = ttk.Frame(self.root)
        hdr.grid(row=3, column=0, sticky="w")
        ttk.Label(hdr, text="On", width=4).grid(row=0, column=0, padx=(8,0))
        ttk.Label(hdr, text="#", width=4).grid(row=0, column=1)
        ttk.Label(hdr, text="Label", width=32).grid(row=0, column=2, sticky="w")
        ttk.Label(hdr, text="Time", width=12).grid(row=0, column=3, sticky="w")
        ttk.Label(hdr, text="T-minus", width=14).grid(row=0, column=4, sticky="w")
        ttk.Label(hdr, text="Select", width=8).grid(row=0, column=5, sticky="w")

        # Scrollable rows area
        self.canvas = tk.Canvas(self.root, width=720, height=360, highlightthickness=0)
        self.scroll_y = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.rows_frame = ttk.Frame(self.canvas)

        self.rows_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.rows_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll_y.set)

        self.canvas.grid(row=4, column=0, sticky="w", padx=8, pady=(0,6))
        self.scroll_y.grid(row=4, column=1, sticky="ns", pady=(0,6))

        # Status
        self.status_var = tk.StringVar(value="Status: Not armed")
        ttk.Label(self.root, textvariable=self.status_var, foreground="#555").grid(row=5, column=0, sticky="w", padx=8, pady=(0,10))

    def _row_widgets(self, idx, data):
        enabled_var = tk.BooleanVar(value=data.get("enabled", True))
        label_var = tk.StringVar(value=data.get("label",""))
        time_var = tk.StringVar(value=data.get("time",""))
        sel_var = tk.BooleanVar(value=False)
        tminus_var = tk.StringVar(value="—")

        r = ttk.Frame(self.rows_frame)
        r.grid(row=idx, column=0, sticky="w")

        chk = ttk.Checkbutton(r, variable=enabled_var, width=2)
        num = ttk.Label(r, text=str(idx+1), width=4)
        ent_label = ttk.Entry(r, width=36, textvariable=label_var)
        ent_time  = ttk.Entry(r, width=12, textvariable=time_var)
        tminus_lbl = ttk.Label(r, textvariable=tminus_var, width=16)
        sel_chk = ttk.Checkbutton(r, variable=sel_var, width=6)

        chk.grid(row=0, column=0, padx=(8,8), pady=4)
        num.grid(row=0, column=1, padx=2)
        ent_label.grid(row=0, column=2, padx=4, sticky="w")
        ent_time.grid(row=0, column=3, padx=4, sticky="w")
        tminus_lbl.grid(row=0, column=4, padx=4, sticky="w")
        sel_chk.grid(row=0, column=5, padx=4, sticky="w")

        return {
            "frame": r,
            "enabled_var": enabled_var,
            "label_var": label_var,
            "time_var": time_var,
            "tminus_var": tminus_var,
            "select_var": sel_var
        }

    def _refresh_numbers(self):
        for i, rv in enumerate(self.rows_vars):
            children = rv["frame"].winfo_children()
            if len(children) >= 2 and isinstance(children[1], ttk.Label):
                children[1].configure(text=str(i+1))

    def _add_row(self, preset=None, at_end=True):
        idx = len(self.rows_vars) if at_end else 0
        data = preset or {"label":"","time":"","enabled":True}
        rv = self._row_widgets(idx, data)
        if at_end:
            self.rows_vars.append(rv)
        else:
            self.rows_vars.insert(0, rv)
            for i, r in enumerate(self.rows_vars):
                r["frame"].grid_forget()
                r["frame"].grid(row=i, column=0, sticky="w")
        self._refresh_numbers()
        self.canvas.yview_moveto(1.0)

    def _duplicate_selected(self):
        sel = [rv for rv in self.rows_vars if rv["select_var"].get()]
        if not sel:
            messagebox.showinfo("Duplicate", "Select at least one row.")
            return
        for rv in sel:
            self._add_row(
                preset={
                    "label": rv["label_var"].get(),
                    "time": rv["time_var"].get(),
                    "enabled": rv["enabled_var"].get()
                },
                at_end=True
            )

    def _delete_selected(self):
        indices = [i for i, rv in enumerate(self.rows_vars) if rv["select_var"].get()]
        if not indices:
            messagebox.showinfo("Delete", "Select at least one row.")
            return
        for i in reversed(indices):
            rv = self.rows_vars.pop(i)
            rv["frame"].destroy()
        self._refresh_numbers()

    def _load_existing(self):
        cfg = load_config()
        alarms = cfg.get("alarms", DEFAULT_ALARMS)
        for item in alarms:
            self._add_row(preset=item, at_end=True)
        self.default_ringtone_path = cfg.get("default_ringtone", "") or ""
        self._update_ringtone_label()
        pol = bool(cfg.get("pause_on_lock", True))
        if PYWIN32_OK:
            self.pause_on_lock_var.set(pol)
        else:
            self.pause_on_lock_var.set(False)

    def _save(self):
        try:
            for rv in self.rows_vars:
                if rv["enabled_var"].get() and rv["time_var"].get().strip():
                    parse_hhmm(rv["time_var"].get())
            save_config(self.rows_vars, self.default_ringtone_path, self.pause_on_lock_var.get())
            messagebox.showinfo("Saved", "Alarm list + settings saved.")
        except Exception as e:
            messagebox.showerror("Invalid time", f"Please fix times: {e}")

    def _pick_wav(self):
        path = filedialog.askopenfilename(
            title="Choose WAV ringtone",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")]
        )
        if path:
            self.default_ringtone_path = path
            self._update_ringtone_label()

    def _clear_wav(self):
        self.default_ringtone_path = ""
        self._update_ringtone_label()

    def _update_ringtone_label(self):
        if self.default_ringtone_path and os.path.exists(self.default_ringtone_path):
            base = os.path.basename(self.default_ringtone_path)
            self.ringtone_var.set(f"Ringtone: {base}")
        else:
            self.ringtone_var.set("Ringtone: Beep (default)")

    def _start_all(self):
        try:
            now = dt.datetime.now()
            self.fired_today.clear()
            any_enabled = False
            for idx, rv in enumerate(self.rows_vars):
                if not rv["enabled_var"].get():
                    continue
                t_str = rv["time_var"].get().strip()
                if not t_str:
                    continue
                h, m = parse_hhmm(t_str)
                target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                if target <= now:
                    # If already passed at start time, mark as fired (skip)
                    self.fired_today.add(idx)
                any_enabled = True
            if not any_enabled:
                messagebox.showwarning("No alarms", "Turn on at least one alarm with a valid time.")
                return
            self.armed = True
            self.status_var.set(f"Status: Armed at {now.strftime('%H:%M:%S')} (today only)")
        except Exception as e:
            messagebox.showerror("Invalid time", f"Please fix times: {e}")

    def _stop_all(self):
        self.armed = False
        self.status_var.set("Status: Not armed")
        self.player.stop()

    def _fmt_tminus(self, secs):
        if secs < 0:
            return "—"
        h = secs // 3600
        m = (secs % 3600) // 60
        s = secs % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        else:
            return f"{m:02d}:{s:02d}"

    def _tick(self):
        now = dt.datetime.now()

        # Update T-minus display
        for idx, rv in enumerate(self.rows_vars):
            t_str = rv["time_var"].get().strip()
            if not t_str:
                rv["tminus_var"].set("—")
                continue
            try:
                h, m = parse_hhmm(t_str)
                target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                delta_sec = int((target - now).total_seconds())
                if (idx in self.fired_today) or (not rv["enabled_var"].get()):
                    rv["tminus_var"].set("—")
                else:
                    rv["tminus_var"].set(self._fmt_tminus(delta_sec))
            except Exception:
                rv["tminus_var"].set("ERR")

        # Fire logic
        if self.armed and not (self.paused and self.pause_on_lock_var.get()):
            for idx, rv in enumerate(self.rows_vars):
                if idx in self.fired_today:
                    continue
                if not rv["enabled_var"].get():
                    continue
                t_str = rv["time_var"].get().strip()
                if not t_str:
                    continue
                try:
                    h, m = parse_hhmm(t_str)
                except Exception:
                    continue
                target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                delta = (now - target).total_seconds()
                if -1 <= delta <= 3:
                    self._fire_alarm(idx, rv["label_var"].get().strip() or f"Alarm {idx+1}")
                    self.fired_today.add(idx)
                elif delta > 3:
                    # Only skip (mark fired) if NOT paused. If paused, we want to catch up on unlock.
                    if not self.paused or not self.pause_on_lock_var.get():
                        self.fired_today.add(idx)

        self.root.after(1000, self._tick)

    def _fire_alarm(self, idx, label_text):
        # Start sound (WAV or beep)
        self.player.play(self.default_ringtone_path)

        popup = tk.Toplevel(self.root)
        popup.title("⏰ Alarm")
        popup.resizable(False, False)

        ttk.Label(popup, text=f"⏰ {label_text}", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=2, padx=16, pady=(16, 8))
        ttk.Label(popup, text=dt.datetime.now().strftime("%H:%M")).grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 8))

        def dismiss():
            self.player.stop()
            popup.destroy()

        def snooze5():
            self.player.stop()
            now = dt.datetime.now()
            new_time = (now + dt.timedelta(minutes=5)).strftime("%H:%M")
            self.rows_vars[idx]["time_var"].set(new_time)
            if idx in self.fired_today:
                self.fired_today.remove(idx)
            popup.destroy()
            messagebox.showinfo("Snoozed", f"Snoozed to {new_time}")

        ttk.Button(popup, text="Dismiss", command=dismiss).grid(row=2, column=0, padx=10, pady=(0, 14))
        ttk.Button(popup, text="Snooze 5 min", command=snooze5).grid(row=2, column=1, padx=10, pady=(0, 14))

        popup.attributes("-topmost", True)
        popup.grab_set()

    # ---------- Pause on lock / resume on unlock ----------
    def _start_session_watcher(self):
        def on_lock():
            # Called from watcher thread → route to Tk main thread
            self.root.after(0, self._handle_lock)
        def on_unlock():
            self.root.after(0, self._handle_unlock)

        self.session_watcher = SessionWatcher(on_lock, on_unlock)
        self.session_watcher.start()

    def _handle_lock(self):
        if not self.pause_on_lock_var.get():
            return
        self.paused = True
        self.player.stop()  # stop any ringing sound
        self.status_var.set("Status: Paused (Windows locked)")

    def _handle_unlock(self):
        if not self.pause_on_lock_var.get():
            return
        self.paused = False
        self.status_var.set("Status: Resumed (Windows unlocked)")
        # Fire any alarms that became due while paused
        if self.armed:
            try:
                now = dt.datetime.now()
                due = []
                for idx, rv in enumerate(self.rows_vars):
                    if idx in self.fired_today: 
                        continue
                    if not rv["enabled_var"].get():
                        continue
                    t_str = rv["time_var"].get().strip()
                    if not t_str:
                        continue
                    try:
                        h, m = parse_hhmm(t_str)
                    except Exception:
                        continue
                    target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if target <= now:
                        due.append((target, idx))
                due.sort()
                for _, idx in due:
                    self._fire_alarm(idx, self.rows_vars[idx]["label_var"].get().strip() or f"Alarm {idx+1}")
                    self.fired_today.add(idx)
            except Exception:
                pass

    def _on_close(self):
        self.player.stop()
        try:
            # Best-effort: send WM_DESTROY to watcher window so PumpMessages exits
            if hasattr(self, "session_watcher") and getattr(self.session_watcher, "hwnd", None):
                try:
                    win32gui.PostMessage(self.session_watcher.hwnd, win32con.WM_DESTROY, 0, 0)
                except Exception:
                    pass
        except Exception:
            pass
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("vista")
    except Exception:
        pass
    app = InfiniteAlarmApp(root)
    root.mainloop()
