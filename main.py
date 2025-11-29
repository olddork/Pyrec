import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.animation as animation
import matplotlib.dates as mdates
import math
import random
import csv
import time
import threading
from datetime import datetime, timedelta, date
import configparser
import os
import glob
import sys

# --- NEW DEPENDENCY CHECK ---
try:
    import serial
    import serial.tools.list_ports
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False

# --- Configuration ---
MAX_CHANNELS = 8
DEFAULT_WINDOW_SIZE = 50 
INI_FILE = "sensor_settings.ini"

# COLORS & FONTS
COLOR_BG = "#f0f2f5"       
COLOR_SIDEBAR = "#ffffff"  
COLOR_ACCENT = "#007acc"   
COLOR_TEXT = "#2d3436"     
FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_HEADER = ("Segoe UI", 12, "bold")
FONT_MONO = ("Consolas", 10)

# PERFORMANCE LIMITS
# Updated to 86400 to allow 24 hours of visibility (60 * 60 * 24)
MAX_POINTS_ON_SCREEN = 86400 

INTERVAL_OPTIONS = {
    "1s": 1, "3s": 3, "5s": 5, "10s": 10, "30s": 30,
    "1min": 60, "2min": 120, "5min": 300, "10min": 600
}

BAUD_RATES = [9600, 19200, 38400, 57600, 115200, 250000]

# ==========================================
#        DATA DRIVER ARCHITECTURE
# ==========================================

class DataSource:
    def __init__(self):
        self.connected = False
        self.latest_data = [0.0] * MAX_CHANNELS
        self.lock = threading.Lock() 

    def connect(self, port, baud):
        self.connected = True
        return True

    def disconnect(self):
        self.connected = False

    def get_data(self):
        with self.lock:
            return list(self.latest_data)

class SimulationSource(DataSource):
    def __init__(self):
        super().__init__()
        self.start_time = time.time()

    def get_data(self):
        # 1. Slow down time (0.2x speed) to make graphs less "hectic"
        t = (time.time() - self.start_time) * 0.2
        
        data = []
        for i in range(MAX_CHANNELS):
            center = 2.0
            amplitude = 2.0 
            frequency = 1.0 + (i * 0.05) 
            phase = i * 0.5 
            
            # Base wave: Result is always 0.0 to 4.0
            val = center + amplitude * math.sin(t * frequency + phase)
            
            # Add tiny noise (+/- 0.05) for realism
            noise = random.uniform(-0.05, 0.05)
            
            data.append(val + noise)
            
        return data

class BalkonLoggerDriver(DataSource):
    def __init__(self):
        super().__init__()
        self.ser = None
        self.running = False
        self.thread = None

    def connect(self, port, baud):
        if not SERIAL_AVAILABLE:
            messagebox.showerror("Error", "pyserial not installed.\nRun: pip install pyserial")
            return False
            
        try:
            self.ser = serial.Serial(port, baud, timeout=2)
            self.running = True
            self.thread = threading.Thread(target=self._reader_loop, daemon=True)
            self.thread.start()
            self.connected = True
            return True
        except Exception as e:
            messagebox.showerror("Connection Failed", str(e))
            return False

    def disconnect(self):
        self.running = False
        self.connected = False
        if self.ser and self.ser.is_open:
            self.ser.close()

    def _reader_loop(self):
        print("Serial thread started...")
        while self.running and self.ser and self.ser.is_open:
            try:
                line = self.ser.readline().decode('utf-8', errors='ignore').strip()
                if line == "eof":
                    temp_data = []
                    for _ in range(16):
                        val_str = self.ser.readline().decode('utf-8', errors='ignore').strip()
                        try:
                            val = float(val_str)
                            temp_data.append(val)
                        except ValueError:
                            pass 
                    
                    if len(temp_data) >= 8:
                        with self.lock:
                            self.latest_data = temp_data[:8]
            except Exception as e:
                time.sleep(1) 

DEVICE_DRIVERS = {
    "BalkonLogger": BalkonLoggerDriver,
}

# ==========================================
#              MAIN APPLICATION
# ==========================================

class CustomToolbar(NavigationToolbar2Tk):
    toolitems = [
        ('Home', 'Reset view', 'home', 'home'),
        ('Pan', 'Pan axes', 'move', 'pan'),
        ('Zoom', 'Zoom to rect', 'zoom_to_rect', 'zoom'),
        ('Save', 'Save image', 'filesave', 'save_figure'),
    ]

    def __init__(self, canvas, window, app_instance):
        self.app = app_instance
        super().__init__(canvas, window)
        self.config(background=COLOR_BG)
        
        self._message_label.pack_forget()

        for child in self.winfo_children():
            if isinstance(child, tk.Button):
                child.config(relief="flat", bg=COLOR_BG, activebackground="#e1e4e8", bd=0)
                child.pack_configure(padx=2, pady=4)

        # Common style for custom buttons to ensure uniformity
        btn_style = {
            "relief": "flat",
            "bg": COLOR_BG,
            "font": FONT_MAIN,
            "cursor": "hand2",
            "width": 12, # Fixed width prevents UI jumping
            "activebackground": "#dfe6e9"
        }

        # --- LEFT TOOLS ---
        ttk.Separator(self, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10, pady=5)
        
        # Export Button
        self.btn_export = tk.Button(self, text="📊 Export View", command=self.app.open_export_window, **btn_style)
        self.btn_export.pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(self, orient='vertical').pack(side=tk.LEFT, fill='y', padx=10, pady=5)
        
        # Scroll Button
        self.btn_scroll = tk.Button(self, text="🔒 Scroll: ON", command=self.toggle_scroll, **btn_style)
        self.btn_scroll.config(fg="#27ae60") # Start green
        self.btn_scroll.pack(side=tk.LEFT, padx=2)
        
        # Pause Button
        self.btn_pause = tk.Button(self, text="⏸ Pause", command=self.toggle_pause, **btn_style)
        self.btn_pause.pack(side=tk.LEFT, padx=2)

        # --- RIGHT CONNECTION BAR ---
        conn_frame = tk.Frame(self, bg=COLOR_BG)
        conn_frame.pack(side=tk.RIGHT, padx=5)

        # CONNECT BUTTON
        self.btn_connect = tk.Button(conn_frame, text="🔌 Connect", command=self.toggle_connect,
                                     bg="#c0392b", fg="white", font=FONT_BOLD, relief="flat", width=15)
        self.btn_connect.pack(side=tk.RIGHT, padx=5)

        self.device_var = tk.StringVar(value="BalkonLogger")
        self.cb_device = ttk.Combobox(conn_frame, textvariable=self.device_var, values=list(DEVICE_DRIVERS.keys()), 
                                      state="readonly", width=12, font=FONT_MAIN)
        self.cb_device.pack(side=tk.RIGHT, padx=2)
        
        self.baud_var = tk.StringVar(value="9600")
        self.cb_baud = ttk.Combobox(conn_frame, textvariable=self.baud_var, values=BAUD_RATES, 
                                    state="readonly", width=7, font=FONT_MAIN)
        self.cb_baud.pack(side=tk.RIGHT, padx=2)

        self.port_var = tk.StringVar(value="Simulation")
        self.cb_port = ttk.Combobox(conn_frame, textvariable=self.port_var, values=["Simulation"], 
                                    state="readonly", width=15, font=FONT_MAIN, 
                                    postcommand=self.refresh_ports)
        self.cb_port.pack(side=tk.RIGHT, padx=2)
        self.cb_port.bind("<<ComboboxSelected>>", self.on_port_change)

        self._message_label.config(background=COLOR_BG, font=FONT_MAIN)
        self._message_label.pack(side=tk.RIGHT, padx=10)

        self.on_port_change() 

    def home(self, *args):
        self.app.reset_view_to_defaults()
        if hasattr(self, '_nav_stack'): self._nav_stack.clear()
        if hasattr(self, '_views'): self._views.clear()
        if hasattr(self, '_positions'): self._positions.clear()
        self.draw()

    def toggle_scroll(self):
        current = self.app.auto_scroll.get()
        new_state = not current
        self.app.auto_scroll.set(new_state)
        if new_state: 
            self.btn_scroll.config(text="🔒 Scroll: ON", fg="#27ae60") 
        else: 
            self.btn_scroll.config(text="🔓 Scroll: OFF", fg="#c0392b") 

    def toggle_pause(self):
        self.app.toggle_capture()
        if self.app.is_running: 
            self.btn_pause.config(text="⏸ Pause", fg=COLOR_TEXT)
        else: 
            self.btn_pause.config(text="▶ RESUME", fg="#2980b9") 

    def refresh_ports(self):
        values = ["Simulation"]
        if SERIAL_AVAILABLE:
            try:
                ports = serial.tools.list_ports.comports()
                for p in ports: values.append(p.device)
            except: pass
        self.cb_port['values'] = values

    def on_port_change(self, event=None):
        pass

    def toggle_connect(self):
        if not self.app.connected:
            port = self.port_var.get()
            baud = int(self.baud_var.get())
            device_name = self.device_var.get()
            
            if self.app.connect_to_source(port, baud, device_name):
                self.btn_connect.config(text="✔ Disconnect", bg="#27ae60")
                self.cb_port.config(state="disabled")
                self.cb_baud.config(state="disabled")
                self.cb_device.config(state="disabled")
        else:
            self.app.disconnect_source()
            self.btn_connect.config(text="🔌 Connect", bg="#c0392b")
            self.cb_port.config(state="readonly")
            self.cb_baud.config(state="readonly")
            self.cb_device.config(state="readonly")

class SensorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pro Logger v3.1 - Refined UI")
        self.root.geometry("1280x850")
        self.root.configure(bg=COLOR_BG)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- Data Source ---
        self.data_source = None
        self.connected = False
        
        # --- Data Storage ---
        self.timestamps = []
        self.datetime_cache = [] 
        # Stores RAW data (no factor/offset applied)
        self.channel_data = [[] for _ in range(MAX_CHANNELS)]
        
        # --- State ---
        self.is_running = True
        self.start_time = time.time()
        self.last_capture_time = 0 
        self.current_log_date = None
        self.log_file = None
        self.csv_writer = None
        self.log_filename = ""
        
        # --- UI Vars ---
        self.auto_scroll = tk.BooleanVar(value=True)
        self.window_size_var = tk.IntVar(value=DEFAULT_WINDOW_SIZE)
        self.y_min = tk.DoubleVar(value=-0.5) 
        self.y_max = tk.DoubleVar(value=4.5)  
        self.save_on_exit_var = tk.BooleanVar(value=True)
        self.interval_var = tk.StringVar(value="1s") 
        self.current_interval_sec = 1.0 
        
        self.ch_vars = {
            'active': [tk.BooleanVar(value=True if i < 3 else False) for i in range(MAX_CHANNELS)],
            'factor': [tk.DoubleVar(value=1.0) for _ in range(MAX_CHANNELS)],
            'offset': [tk.DoubleVar(value=0.0) for _ in range(MAX_CHANNELS)],
            'colors': ['#2980b9', '#27ae60', '#c0392b', '#16a085', '#8e44ad', '#f39c12', '#2c3e50', '#d35400'],
            'current_val_str': [tk.StringVar(value="0.00") for _ in range(MAX_CHANNELS)]
        }

        self._apply_theme()
        self.load_settings()
        self.init_daily_log_system()
        self._setup_ui()
        self.update_slider_range()
        
        # Auto-connect Simulation
        self.connect_to_source("Simulation", 9600, "BalkonLogger")
        self.toolbar_ref.btn_connect.config(text="✔ Disconnect", bg="#27ae60")
        self.toolbar_ref.cb_port.config(state="disabled")
        self.toolbar_ref.cb_baud.config(state="disabled")
        self.toolbar_ref.cb_device.config(state="disabled")
        
        self.ani = animation.FuncAnimation(self.fig, self.update_plot, interval=200, blit=False, cache_frame_data=False)

    def _apply_theme(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=FONT_MAIN)
        style.configure("TLabel", background=COLOR_BG, font=FONT_MAIN)
        style.configure("TButton", font=FONT_BOLD, background="#dfe6e9", borderwidth=0)
        style.map("TButton", background=[("active", "#b2bec3")])
        style.configure("Card.TFrame", background=COLOR_SIDEBAR, relief="flat")
        style.configure("Card.TLabel", background=COLOR_SIDEBAR, font=FONT_MAIN)
        style.configure("Header.TLabel", background=COLOR_SIDEBAR, font=FONT_HEADER, foreground=COLOR_ACCENT)
        style.configure("Card.TCheckbutton", background=COLOR_SIDEBAR)
        style.configure("Status.TLabel", background=COLOR_SIDEBAR, foreground="#27ae60", font=("Consolas", 8))

    def connect_to_source(self, port, baud, device_name):
        if port == "Simulation":
            DriverClass = SimulationSource
        else:
            DriverClass = DEVICE_DRIVERS.get(device_name, BalkonLoggerDriver)
            
        self.data_source = DriverClass()
        if self.data_source.connect(port, baud):
            self.connected = True
            self.lbl_status.config(text=f"Connected: {port}\nDevice: {device_name}")
            return True
        else:
            self.data_source = None
            self.connected = False
            return False

    def disconnect_source(self):
        if self.data_source:
            self.data_source.disconnect()
        self.data_source = None
        self.connected = False
        self.lbl_status.config(text=f"Status: Disconnected\nLog: {self.log_filename}")

    # --- LOGGING ---
    def get_daily_filename(self):
        return f"log_{datetime.now().strftime('%Y-%m-%d')}.csv"

    def init_daily_log_system(self):
        now = datetime.now()
        today_str = now.strftime('%Y-%m-%d')
        self.current_log_date = today_str
        self.log_filename = self.get_daily_filename()
        
        # Reset caches
        self.timestamps = []
        self.datetime_cache = []
        self.channel_data = [[] for _ in range(MAX_CHANNELS)]

        # --- LOAD HISTORY (Last 3 days) ---
        all_logs = glob.glob("log_*.csv")
        all_logs.sort()
        
        relevant_files = []
        cutoff_date = (now - timedelta(days=3)).date() 
        
        for fname in all_logs:
            try:
                d_str = os.path.basename(fname).replace("log_", "").replace(".csv", "")
                f_date = datetime.strptime(d_str, "%Y-%m-%d").date()
                if f_date >= cutoff_date:
                    relevant_files.append(fname)
            except Exception:
                pass 

        prev_ts = None
        
        for fname in relevant_files:
            try:
                with open(fname, 'r') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if not row or row[0].startswith("#") or row[0].startswith("Timestamp"): continue
                        try:
                            ts_val = float(row[0])
                            
                            # Gap detection
                            if prev_ts is not None and (ts_val - prev_ts) > (self.current_interval_sec * 5):
                                pad_ts = ts_val - 0.001
                                self.timestamps.append(pad_ts)
                                self.datetime_cache.append(datetime.fromtimestamp(pad_ts))
                                for i in range(MAX_CHANNELS): self.channel_data[i].append(float('nan'))
                            
                            # Real Data
                            self.timestamps.append(ts_val)
                            self.datetime_cache.append(datetime.fromtimestamp(ts_val))
                            
                            for i in range(MAX_CHANNELS):
                                col_idx = 2 + i
                                # Read as RAW (assuming CSV now contains raw data)
                                val = float(row[col_idx]) if col_idx < len(row) else 0.0
                                self.channel_data[i].append(val)
                            prev_ts = ts_val
                        except ValueError: continue
            except Exception as e:
                print(f"Error reading {fname}: {e}")

        if self.timestamps:
            self.last_capture_time = self.timestamps[-1]

        try:
            file_exists = os.path.exists(self.log_filename)
            self.log_file = open(self.log_filename, mode='a', newline='')
            self.csv_writer = csv.writer(self.log_file)
            if not file_exists:
                self.csv_writer.writerow(["# DAILY LOG START", today_str])
                self.csv_writer.writerow(["Timestamp_Unix", "Timestamp_ISO"] + [f"Ch_{i+1}" for i in range(MAX_CHANNELS)])
        except Exception as e:
            messagebox.showerror("File Error", str(e))
            self.is_running = False

    def check_rollover(self):
        if datetime.now().strftime('%Y-%m-%d') != self.current_log_date:
            if self.log_file: self.log_file.close()
            self.init_daily_log_system()
            if hasattr(self, 'lbl_status'):
                self.lbl_status.config(text=f"Active Log:\n{self.log_filename}")

    # --- UI LAYOUT ---
    def _setup_ui(self):
        sidebar = ttk.Frame(self.root, padding="15", width=350, style="Card.TFrame")
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        
        self.lbl_status = ttk.Label(sidebar, text=f"Source: Simulation\nLog: {self.log_filename}", style="Status.TLabel")
        self.lbl_status.pack(pady=(0, 15), anchor="w")
        
        ttk.Label(sidebar, text="Sampling Rate", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        interval_cb = ttk.Combobox(sidebar, textvariable=self.interval_var, values=list(INTERVAL_OPTIONS.keys()), state="readonly", font=FONT_MAIN)
        interval_cb.pack(fill=tk.X, pady=(0, 15))
        interval_cb.bind("<<ComboboxSelected>>", self.on_interval_change)
        
        ttk.Label(sidebar, text="Channels", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        grid_frame = ttk.Frame(sidebar, style="Card.TFrame")
        grid_frame.pack(fill=tk.X)
        headers = ["Ch", "On", "Value", "Factor", "Offset"]
        for col, txt in enumerate(headers):
            ttk.Label(grid_frame, text=txt, style="Card.TLabel", font=("Segoe UI", 8, "bold")).grid(row=0, column=col, sticky="w", padx=2)

        for i in range(MAX_CHANNELS):
            row = i + 1
            lbl = ttk.Label(grid_frame, text=f"■ {i+1}", foreground=self.ch_vars['colors'][i], font=FONT_BOLD, style="Card.TLabel")
            lbl.grid(row=row, column=0, padx=2, pady=4, sticky="w")
            cb = ttk.Checkbutton(grid_frame, variable=self.ch_vars['active'][i], style="Card.TCheckbutton")
            cb.grid(row=row, column=1)
            v_lbl = ttk.Label(grid_frame, textvariable=self.ch_vars['current_val_str'][i], width=8, anchor="e", font=FONT_MONO, background="#f1f2f6")
            v_lbl.grid(row=row, column=2, padx=5)
            
            # Use Spinbox for arrows (increment=0.1)
            spin_f = ttk.Spinbox(grid_frame, from_=-1000.0, to=1000.0, increment=0.1, 
                                 textvariable=self.ch_vars['factor'][i], width=5, font=FONT_MONO)
            spin_f.grid(row=row, column=3, padx=2)
            
            spin_o = ttk.Spinbox(grid_frame, from_=-10000.0, to=10000.0, increment=0.1, 
                                 textvariable=self.ch_vars['offset'][i], width=5, font=FONT_MONO)
            spin_o.grid(row=row, column=4, padx=2)

        ttk.Separator(sidebar, orient='horizontal').pack(fill='x', pady=20)
        ttk.Label(sidebar, text="Display Settings", style="Header.TLabel").pack(anchor="w", pady=(0, 5))
        ttk.Label(sidebar, text="Time Window Size", style="Card.TLabel").pack(anchor="w")
        self.time_slider = ttk.Scale(sidebar, from_=10, to=1000, variable=self.window_size_var, orient='horizontal', command=self.on_slider_move)
        self.time_slider.pack(fill=tk.X, pady=5)
        
        y_frame = ttk.Frame(sidebar, style="Card.TFrame")
        y_frame.pack(fill=tk.X, pady=5)
        ttk.Label(y_frame, text="Y-Min:", style="Card.TLabel").pack(side=tk.LEFT)
        e1 = ttk.Entry(y_frame, textvariable=self.y_min, width=6, font=FONT_MONO)
        e1.pack(side=tk.LEFT, padx=5)
        e1.bind('<Return>', lambda e: self.apply_scale())
        ttk.Label(y_frame, text="Y-Max:", style="Card.TLabel").pack(side=tk.LEFT)
        e2 = ttk.Entry(y_frame, textvariable=self.y_max, width=6, font=FONT_MONO)
        e2.pack(side=tk.LEFT, padx=5)
        e2.bind('<Return>', lambda e: self.apply_scale())
        ttk.Button(sidebar, text="Update Scale", command=self.apply_scale).pack(fill=tk.X, pady=2)
        
        footer = ttk.Frame(sidebar, style="Card.TFrame")
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Checkbutton(footer, text="Save Settings on Exit", variable=self.save_on_exit_var, style="Card.TCheckbutton").pack(anchor="w", pady=(0, 10))
        
        tk.Button(footer, text="EXIT APPLICATION", command=self.on_close,
                  bg="#ffcccc", fg="red", font=("Segoe UI", 10, "bold"), relief="raised").pack(fill=tk.X, pady=5)

        main_frame = ttk.Frame(self.root)
        main_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.fig, self.ax = plt.subplots(figsize=(5, 4), dpi=100, constrained_layout=True)
        
        self.fig.patch.set_facecolor(COLOR_BG) 
        self.ax.set_facecolor("white")
        self.ax.set_title("Live Sensor Data", fontsize=12, fontweight='bold', color=COLOR_TEXT)
        self.ax.tick_params(colors=COLOR_TEXT, labelsize=9)
        self.ax.xaxis.label.set_color(COLOR_TEXT)
        self.ax.yaxis.label.set_color(COLOR_TEXT)
        self.ax.spines['bottom'].set_color('#b2bec3')
        self.ax.spines['top'].set_color('#b2bec3') 
        self.ax.spines['right'].set_color('#b2bec3')
        self.ax.spines['left'].set_color('#b2bec3')
        self.ax.grid(True, color='#dfe6e9', linestyle='--', linewidth=0.5)
        self.ax.set_ylim(self.y_min.get(), self.y_max.get())
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
        self.lines = []
        for i in range(MAX_CHANNELS):
            line, = self.ax.plot([], [], label=f"Ch {i+1}", color=self.ch_vars['colors'][i], linewidth=1.5)
            self.lines.append(line)
        self.canvas = FigureCanvasTkAgg(self.fig, master=main_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.toolbar_ref = CustomToolbar(self.canvas, main_frame, self)
        self.toolbar_ref.update()
        self.toolbar_ref.pack(side=tk.BOTTOM, fill=tk.X) 

    # --- VIEW RESET ---
    def reset_view_to_defaults(self):
        self.auto_scroll.set(True)
        if hasattr(self, 'btn_scroll'): pass 
        self.apply_scale()
        if self.timestamps:
            window_pts = self.window_size_var.get()
            seconds_width = window_pts * self.current_interval_sec
            latest_time = self.datetime_cache[-1]
            start_time = latest_time - timedelta(seconds=seconds_width)
            self.ax.set_xlim(start_time, latest_time)
            self.canvas.draw_idle()

    # --- SLIDER LOGIC ---
    def update_slider_range(self):
        self.time_slider.config(to=MAX_POINTS_ON_SCREEN)
        if self.window_size_var.get() > MAX_POINTS_ON_SCREEN:
            self.window_size_var.set(MAX_POINTS_ON_SCREEN)

    def on_interval_change(self, event):
        text_val = self.interval_var.get()
        if text_val in INTERVAL_OPTIONS:
            self.current_interval_sec = INTERVAL_OPTIONS[text_val]
            self.last_capture_time = 0 
            self.update_slider_range()

    def on_slider_move(self, val):
        pass

    # --- EXPORT LOGIC ---
    def open_export_window(self):
        top = tk.Toplevel(self.root)
        top.title("Export View Data")
        top.geometry("500x500") # Increased height for channels
        top.configure(bg=COLOR_BG)
        c_frame = ttk.Frame(top, padding=20)
        c_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(c_frame, text="Export Settings", font=FONT_HEADER).pack(pady=(0,10))
        
        # --- Channel Selection ---
        ch_frame = ttk.LabelFrame(c_frame, text="Select Channels to Export", padding=10)
        ch_frame.pack(fill=tk.X, pady=10)
        
        export_vars = []
        for i in range(MAX_CHANNELS):
            # Default to active state in main window
            is_active = self.ch_vars['active'][i].get()
            var = tk.BooleanVar(value=is_active)
            export_vars.append(var)
            
            # 2 columns of checkboxes
            col = i % 2
            row = i // 2
            cb = ttk.Checkbutton(ch_frame, text=f"Channel {i+1}", variable=var)
            cb.grid(row=row, column=col, sticky="w", padx=10, pady=2)

        # --- Time Range ---
        ttk.Label(c_frame, text="Time Range", font=FONT_BOLD).pack(anchor="w", pady=(10, 5))
        
        x_min, x_max = self.ax.get_xlim()
        try:
            start_dt = mdates.num2date(x_min).replace(tzinfo=None) 
            end_dt = mdates.num2date(x_max).replace(tzinfo=None)
        except:
            start_dt = datetime.now() - timedelta(hours=1)
            end_dt = datetime.now()

        start_def = start_dt.strftime("%Y-%m-%d %H:%M")
        end_def = end_dt.strftime("%Y-%m-%d %H:%M")
        
        ttk.Label(c_frame, text="Start Time (YYYY-MM-DD HH:MM):").pack(anchor="w")
        s_entry = ttk.Entry(c_frame, width=30)
        s_entry.insert(0, start_def)
        s_entry.pack(fill=tk.X, pady=(5, 5))
        
        ttk.Label(c_frame, text="End Time (YYYY-MM-DD HH:MM):").pack(anchor="w")
        e_entry = ttk.Entry(c_frame, width=30)
        e_entry.insert(0, end_def)
        e_entry.pack(fill=tk.X, pady=(5, 20))
        
        def run_export():
            try:
                s_dt = datetime.strptime(s_entry.get(), "%Y-%m-%d %H:%M")
                e_dt = datetime.strptime(e_entry.get(), "%Y-%m-%d %H:%M")
                s_ts = s_dt.timestamp()
                e_ts = e_dt.timestamp()
                
                if s_ts >= e_ts:
                    messagebox.showerror("Error", "Start time must be before end time.")
                    return
                
                # Get list of selected channels (booleans)
                selected_channels = [v.get() for v in export_vars]

                self.process_export(s_dt, e_dt, selected_channels, top)
            except ValueError: 
                messagebox.showerror("Format Error", "Invalid Date Format.\nUse: YYYY-MM-DD HH:MM")
                
        ttk.Button(c_frame, text="Export Now", command=run_export).pack(fill=tk.X, pady=10)

    def process_export(self, start_dt, end_dt, channels_to_export, popup):
        try: 
            import pandas as pd
            import openpyxl
            from openpyxl.utils import get_column_letter
            from openpyxl.styles import PatternFill, Font
        except ImportError:
            messagebox.showerror("Error", "Install pandas and openpyxl: pip install pandas openpyxl")
            return
            
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"export.xlsx")
        if not out: return

        all_logs = glob.glob("log_*.csv")
        relevant_files = []
        req_start_date = start_dt.date()
        req_end_date = end_dt.date()

        for log in all_logs:
            try:
                date_str = log.replace("log_", "").replace(".csv", "")
                file_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                if req_start_date <= file_date <= req_end_date:
                    relevant_files.append(log)
            except ValueError:
                continue 

        relevant_files.sort()
        
        if not relevant_files:
            messagebox.showwarning("Empty", "No log files found for this date range.")
            return

        combined = []
        rows_found = 0
        start_ts = start_dt.timestamp()
        end_ts = end_dt.timestamp()

        popup_btn = None
        for child in popup.winfo_children():
            if isinstance(child, ttk.Button): popup_btn = child

        if popup_btn: popup_btn.config(text="Processing... Please Wait", state="disabled")
        popup.update()

        try:
            for log in relevant_files:
                try:
                    chunk_iter = pd.read_csv(log, comment='#', on_bad_lines='skip', chunksize=10000)
                    for chunk in chunk_iter:
                        if 'Timestamp_Unix' not in chunk.columns: continue
                        mask = (chunk['Timestamp_Unix'] >= start_ts) & (chunk['Timestamp_Unix'] <= end_ts)
                        filtered_chunk = chunk.loc[mask]
                        if not filtered_chunk.empty:
                            
                            # Just collect Raw Data
                            if 'Timestamp_ISO' in filtered_chunk.columns: 
                                final_chunk = filtered_chunk.drop(columns=['Timestamp_Unix'])
                            else:
                                filtered_chunk['Timestamp'] = pd.to_datetime(filtered_chunk['Timestamp_Unix'], unit='s')
                                final_chunk = filtered_chunk.drop(columns=['Timestamp_Unix'])
                            combined.append(final_chunk)
                            rows_found += len(final_chunk)
                except Exception as e: 
                    print(f"Skipping bad file/chunk in {log}: {e}")
                    continue

            if combined:
                full_df = pd.concat(combined, ignore_index=True)
                
                # 1. Rename Channels
                rename_map = {f"Ch_{i+1}": f"Raw_Ch_{i+1}" for i in range(MAX_CHANNELS)}
                full_df.rename(columns=rename_map, inplace=True)
                
                # --- FILTER CHANNELS BASED ON DIALOG SELECTION ---
                for i in range(MAX_CHANNELS):
                    # If user unchecked this channel in export dialog
                    if not channels_to_export[i]:
                        col_name = f"Raw_Ch_{i+1}"
                        if col_name in full_df.columns:
                            full_df.drop(columns=[col_name], inplace=True)

                # 2. Add empty columns for SELECTED calculations only
                for i in range(MAX_CHANNELS):
                    if channels_to_export[i]:
                        full_df[f"Cal_Ch_{i+1}"] = "" 

                # 3. Write Data starting at Row 4 (index 3)
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    full_df.to_excel(writer, index=False, startrow=3)
                
                # 4. Open to inject Metadata & Formulas
                wb = openpyxl.load_workbook(out)
                ws = wb.active
                
                # Labels
                ws['A2'] = "Factor"
                ws['A3'] = "Offset"
                ws['A2'].font = Font(bold=True)
                ws['A3'].font = Font(bold=True)

                # Map Headers to Column Letters
                header_row = 4
                col_map = {}
                for cell in ws[header_row]:
                    if cell.value:
                        col_map[cell.value] = get_column_letter(cell.column)
                        # Set auto width for all columns initially
                        ws.column_dimensions[get_column_letter(cell.column)].width = 18 
                
                max_row = ws.max_row
                
                for i in range(MAX_CHANNELS):
                    # Only process if selected for export
                    if not channels_to_export[i]:
                        continue
                    
                    raw_col = f"Raw_Ch_{i+1}"
                    cal_col = f"Cal_Ch_{i+1}"
                    
                    if raw_col in col_map and cal_col in col_map:
                        r_let = col_map[raw_col]
                        c_let = col_map[cal_col]
                        
                        # Get settings from UI
                        f_val = self.ch_vars['factor'][i].get()
                        o_val = self.ch_vars['offset'][i].get()
                        
                        # Set Header Parameters
                        ws[f"{c_let}2"] = f_val
                        ws[f"{c_let}3"] = o_val
                        
                        # Fill Formulas
                        for row in range(5, max_row + 1):
                            ws[f"{c_let}{row}"] = f"={r_let}{row}*{c_let}$2+{c_let}$3"
                            
                        # Apply Colors
                        hex_color = self.ch_vars['colors'][i].replace('#', '')
                        fill_color = "FF" + hex_color
                        header_cell = ws[f"{c_let}{header_row}"]
                        header_cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

                wb.save(out)
                messagebox.showinfo("Success", f"Export complete.\nSaved {rows_found} rows to {os.path.basename(out)}")
                popup.destroy()
            else: 
                messagebox.showwarning("Empty", "No data found inside the specific time range within the files.")
                popup.destroy()
                
        except Exception as e: 
            messagebox.showerror("Export Error", f"Critical Error:\n{str(e)}")
            if popup_btn: popup_btn.config(text="Export Failed", state="normal")

    # --- INI ---
    def load_settings(self):
        if not os.path.exists(INI_FILE): return
        config = configparser.ConfigParser()
        config.read(INI_FILE)
        try:
            if 'Graph' in config:
                self.window_size_var.set(config['Graph'].getint('window_size', DEFAULT_WINDOW_SIZE))
                self.y_min.set(config['Graph'].getfloat('y_min', -0.5))
                self.y_max.set(config['Graph'].getfloat('y_max', 4.5))
                iv = config['Graph'].get('interval', '1s')
                if iv in INTERVAL_OPTIONS:
                    self.interval_var.set(iv)
                    self.current_interval_sec = INTERVAL_OPTIONS[iv]
            for i in range(MAX_CHANNELS):
                sect = f'Channel_{i}'
                if sect in config:
                    self.ch_vars['active'][i].set(config[sect].getboolean('active', False))
                    self.ch_vars['factor'][i].set(config[sect].getfloat('factor', 1.0))
                    self.ch_vars['offset'][i].set(config[sect].getfloat('offset', 0.0))
        except: pass

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Graph'] = {
            'window_size': str(self.window_size_var.get()),
            'y_min': str(self.y_min.get()),
            'y_max': str(self.y_max.get()),
            'interval': self.interval_var.get()
        }
        for i in range(MAX_CHANNELS):
            config[f'Channel_{i}'] = {
                'active': str(self.ch_vars['active'][i].get()),
                'factor': str(self.ch_vars['factor'][i].get()),
                'offset': str(self.ch_vars['offset'][i].get())
            }
        with open(INI_FILE, 'w') as f: config.write(f)

    def on_close(self):
        self.is_running = False
        if self.data_source: self.data_source.disconnect()
        if self.log_file: self.log_file.close()
        if self.save_on_exit_var.get(): self.save_settings()
        self.root.destroy()

    def update_plot(self, frame):
        try:
            self.check_rollover()
            now_ts = time.time()
            time_diff = now_ts - self.last_capture_time

            if self.connected and self.is_running and time_diff >= self.current_interval_sec:
                raw_data = self.data_source.get_data()

                gap_threshold = self.current_interval_sec * 4.0
                if self.last_capture_time != 0 and time_diff > gap_threshold:
                    pad_ts = now_ts - (self.current_interval_sec / 2)
                    self.timestamps.append(pad_ts)
                    self.datetime_cache.append(datetime.fromtimestamp(pad_ts))
                    for i in range(MAX_CHANNELS): self.channel_data[i].append(float('nan'))

                self.last_capture_time = now_ts
                self.timestamps.append(now_ts)
                self.datetime_cache.append(datetime.fromtimestamp(now_ts))
                
                row = [now_ts, datetime.fromtimestamp(now_ts).strftime("%Y-%m-%d %H:%M:%S.%f")]
                for i in range(MAX_CHANNELS):
                    # --- RAW DATA CAPTURE ---
                    # We store the RAW value (no factor/offset) in memory and in CSV
                    val = raw_data[i] if i < len(raw_data) else 0.0
                    
                    self.channel_data[i].append(val) 
                    row.append(val)
                    
                    # Update Sidebar text (Calculated for display only)
                    f = self.ch_vars['factor'][i].get()
                    o = self.ch_vars['offset'][i].get()
                    calc_val = (val * f) + o
                    self.ch_vars['current_val_str'][i].set(f"{calc_val:.2f}")
                    
                if self.csv_writer:
                    self.csv_writer.writerow(row)
                    self.log_file.flush()

            if self.timestamps:
                limit = len(self.timestamps)
                if self.auto_scroll.get():
                     limit = int(self.window_size_var.get() * 1.5)
                
                slice_dt = self.datetime_cache[-limit:]
                
                for i, line in enumerate(self.lines):
                    if self.ch_vars['active'][i].get():
                        # --- DYNAMIC CALCULATION ---
                        # Get RAW slice
                        raw_slice = self.channel_data[i][-limit:]
                        # Get CURRENT Factor/Offset
                        f = self.ch_vars['factor'][i].get()
                        o = self.ch_vars['offset'][i].get()
                        # Apply calculation live to the slice
                        calc_slice = [(v * f) + o for v in raw_slice]
                        
                        line.set_data(slice_dt, calc_slice)
                        line.set_visible(True)
                    else: line.set_visible(False)
                
                window_pts = self.window_size_var.get()
                width_seconds = window_pts * self.current_interval_sec
                
                if self.auto_scroll.get():
                    if slice_dt:
                        self.ax.set_xlim(slice_dt[-1] - timedelta(seconds=width_seconds), slice_dt[-1])
                else:
                    curr_min, curr_max = self.ax.get_xlim()
                    center = (curr_min + curr_max) / 2.0
                    width_days = width_seconds / 86400.0 
                    self.ax.set_xlim(center - (width_days/2.0), center + (width_days/2.0))
            
            return self.lines
        except Exception as e:
            print(f"Plot Error: {e}")
            return self.lines

    def apply_scale(self):
        try:
            self.ax.set_ylim(self.y_min.get(), self.y_max.get())
            self.canvas.draw_idle()
        except ValueError: pass

    def toggle_capture(self):
        self.is_running = not self.is_running

if __name__ == "__main__":
    root = tk.Tk()
    app = SensorApp(root)
    root.mainloop()