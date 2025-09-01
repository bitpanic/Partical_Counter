import os
import sys
import threading
import time
import csv
import queue
import subprocess
import re
from collections import deque, namedtuple
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import ttk

# Third-party imports. These are expected to be installed per README instructions.
# Try multiple module names for SPS30 SHDLC device to support both PyPI and local archives.
try:
    from sensirion_shdlc_driver import ShdlcSerialPort, ShdlcConnection
except Exception:
    ShdlcSerialPort = None
    ShdlcConnection = None

# Optional modern UART SPS30 driver and adapters
Sps30UartDevice = None
ShdlcChannel = None
OutputFormat = None
try:
    from sensirion_driver_adapters.shdlc_adapter.shdlc_channel import ShdlcChannel as _ShdlcChannel  # type: ignore
    ShdlcChannel = _ShdlcChannel
    from sensirion_uart_sps30.device import Sps30Device as _Sps30UartDevice  # type: ignore
    Sps30UartDevice = _Sps30UartDevice
    try:
        from sensirion_uart_sps30.commands import OutputFormat as _OutputFormat  # type: ignore
        OutputFormat = _OutputFormat
    except Exception:
        OutputFormat = None
except Exception:
    Sps30UartDevice = None
    ShdlcChannel = None
    OutputFormat = None

Sps30ShdlcDevice = None
try:
    from sensirion_shdlc_sps import Sps30ShdlcDevice as _Sps30ShdlcDevice
    Sps30ShdlcDevice = _Sps30ShdlcDevice
except Exception:
    try:
        # Local archive variant name
        from shdlc_sps30 import Sps30ShdlcDevice as _Sps30ShdlcDevice
        Sps30ShdlcDevice = _Sps30ShdlcDevice
    except Exception:
        Sps30ShdlcDevice = None

try:
    import pystray
    from pystray import MenuItem as Item, Menu as TrayMenu
    from PIL import Image, ImageDraw
except Exception:
    pystray = None
    Item = None
    TrayMenu = None
    Image = None
    ImageDraw = None

try:
    import matplotlib
    matplotlib.use('TkAgg')
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    try:
        from matplotlib.ticker import FormatStrFormatter as _FormatStrFormatter
    except Exception:
        _FormatStrFormatter = None
except Exception:
    Figure = None
    FigureCanvasTkAgg = None
    _FormatStrFormatter = None


class CONFIG:
    # User-configurable settings
    uart_port = None  # e.g. "COM5"; None = auto-scan available COM ports
    sample_period_s = 5.0
    # Logs directory: when frozen (EXE), use %LOCALAPPDATA%\SPS30 Tray Logger\logs; otherwise ./logs
    if getattr(sys, 'frozen', False):
        _base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        log_dir = os.path.join(_base, 'StratusVision SFP', 'logs')
    else:
        log_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'logs')
    app_name = 'StratusVision SFP'


Sample = namedtuple('Sample', ['ts', 'pm1', 'pm25', 'pm4', 'pm10'])


def ensure_dir(path):
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def write_connection_log(message):
    try:
        ensure_dir(CONFIG.log_dir)
        path = os.path.join(CONFIG.log_dir, 'connection.log')
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(path, 'a', encoding='utf-8') as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass


def get_asset_path(*parts):
    try:
        base = getattr(sys, '_MEIPASS', None) or os.path.abspath(os.path.dirname(__file__))
        return os.path.join(base, 'assets', *parts)
    except Exception:
        return None


def load_logo_pil(desired_size=64):
    if Image is None:
        return None
    try:
        path = get_asset_path('logo.png')
        if path and os.path.isfile(path):
            img = Image.open(path).convert('RGBA')
            if desired_size:
                img = img.copy()
                img.thumbnail((desired_size, desired_size), Image.LANCZOS)
            return img
    except Exception:
        return None
    return None

def detect_serial_ports():
    """Return a list of available COM ports on Windows like ["COM3", "COM5"].

    Tries multiple strategies in order of preference:
    1) pyserial if available
    2) PowerShell: [System.IO.Ports.SerialPort]::GetPortNames()
    3) Fallback to a static scan of COM1..COM40 (may include non-existent ports)
    """
    # Attempt pyserial
    try:
        import serial.tools.list_ports as list_ports  # type: ignore
        ports = [p.device for p in list_ports.comports()]
    except Exception:
        ports = []

    # Attempt PowerShell if no ports found
    if not ports and sys.platform == 'win32':
        try:
            completed = subprocess.run([
                'powershell', '-NoProfile', '-Command',
                "[System.IO.Ports.SerialPort]::GetPortNames() -join '|'"
            ], capture_output=True, text=True, timeout=3)
            if completed.returncode == 0:
                out = (completed.stdout or '').strip()
                if out:
                    ports = [p for p in out.split('|') if p]
        except Exception:
            pass

    # Final fallback: static range (filter to numeric COM names only)
    if not ports:
        ports = [f'COM{i}' for i in range(1, 41)]

    # Normalize and sort by port number if possible
    def _key(name):
        m = re.match(r'COM(\d+)$', str(name).upper())
        return int(m.group(1)) if m else 9999

    unique = sorted({p.upper() for p in ports}, key=_key)
    return unique


class SPS30Reader:
    def __init__(self, uart_port, sample_period_s, on_sample_callback=None):
        self.uart_port = uart_port
        self.sample_period_s = max(0.5, float(sample_period_s))
        self.on_sample_callback = on_sample_callback
        self._stop_event = threading.Event()
        self._paused = threading.Event()
        self._paused.clear()
        self._thread = None
        self._device = None
        self._conn = None
        self._port = None
        self._consecutive_read_failures = 0
        # Allow many transient read failures before reconnecting to avoid UI toggling
        self._max_consecutive_failures = 50
        self._last_sample_time = None
        self._measurement_started = False
        self._warmup_deadline = 0.0
        self._debug_sample_counter = 0

    def start(self):
        if self._thread is not None:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, name='SPS30Reader', daemon=True)
        self._thread.start()

    def pause(self):
        self._paused.set()

    def resume(self):
        self._paused.clear()
        # Ensure measurement is running after resuming
        try:
            self._ensure_started()
        except Exception:
            pass

    def is_paused(self):
        return self._paused.is_set()

    def stop(self):
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=3.0)
            self._thread = None
        self._disconnect()

    def set_port(self, port_name):
        """Change the target COM port at runtime. None => auto-scan."""
        self.uart_port = port_name
        # Force reconnect on next loop
        self._disconnect()
        self._last_sample_time = None
        self._measurement_started = False

    def is_connected(self):
        # Consider connected only if we have a recent successful sample
        if self._device is None or self._last_sample_time is None:
            return False
        return (datetime.now() - self._last_sample_time) < timedelta(seconds=20)

    def _disconnect(self):
        try:
            if self._device is not None:
                try:
                    # Both legacy and modern drivers expose stop_measurement()
                    if hasattr(self._device, 'stop_measurement'):
                        self._device.stop_measurement()
                except Exception:
                    pass
        finally:
            self._device = None
        try:
            if self._conn is not None:
                self._conn = None
        finally:
            pass
        try:
            if self._port is not None:
                try:
                    self._port.close()
                except Exception:
                    pass
        finally:
            self._port = None

    def _connect(self):
        if ShdlcSerialPort is None or ShdlcConnection is None or Sps30ShdlcDevice is None:
            # If legacy driver isn't fully available, try modern UART driver instead
            if not (ShdlcSerialPort and ShdlcChannel and Sps30UartDevice):
                return False
        ports_to_try = []
        if self.uart_port:
            ports_to_try = [self.uart_port]
        else:
            # Use detected ports if available; fall back to a reasonable scan range
            try:
                detected = [p for p in detect_serial_ports() if p.upper().startswith('COM')]
            except Exception:
                detected = []
            ports_to_try = detected if detected else [f'COM{i}' for i in range(3, 41)]
        for port_name in ports_to_try:
            try:
                # Add a small response padding to improve robustness
                try:
                    port = ShdlcSerialPort(port_name, baudrate=115200, additional_response_time=0.20)
                except TypeError:
                    # Older drivers may not support additional_response_time
                    port = ShdlcSerialPort(port_name, baudrate=115200)
                # Prefer modern UART driver first (more consistent measurement API)
                if ShdlcChannel is not None and Sps30UartDevice is not None:
                    channel = ShdlcChannel(port)
                    dev = Sps30UartDevice(channel)
                    self._device = dev
                    self._port = port
                    self._conn = None
                    try:
                        write_connection_log(f"Using UART driver on {port_name}")
                    except Exception:
                        pass
                    try:
                        methods = [n for n in dir(self._device) if ('read' in n or 'data' in n) and not n.startswith('_')]
                        write_connection_log(f"UART device methods: {methods}")
                    except Exception:
                        pass
                    self._ensure_started()
                elif Sps30ShdlcDevice is not None:
                    conn = ShdlcConnection(port)
                    dev = Sps30ShdlcDevice(conn, slave_address=0)
                    self._device = dev
                    self._port = port
                    self._conn = conn
                    try:
                        write_connection_log(f"Using SHDLC driver on {port_name}")
                    except Exception:
                        pass
                    try:
                        methods = [n for n in dir(self._device) if ('read' in n or 'data' in n) and not n.startswith('_')]
                        write_connection_log(f"SHDLC device methods: {methods}")
                    except Exception:
                        pass
                    self._ensure_started()
                else:
                    raise RuntimeError('No compatible SPS30 driver available')

                # Initial warm-up before first reads
                self._warmup_deadline = time.time() + 5.0
                return True
            except Exception:
                # Try next port
                try:
                    port.close()
                except Exception:
                    pass
                self._port = None
                self._conn = None
                self._device = None
        return False

    def _ensure_started(self):
        """Ensure the device is in measurement mode for both driver types."""
        if self._device is None:
            return
        if self._measurement_started:
            return
        try:
            if hasattr(self._device, 'start_measurement'):
                if OutputFormat is not None:
                    try:
                        # Prefer modern API signature when available
                        self._device.start_measurement(OutputFormat(261))
                        self._measurement_started = True
                        try:
                            write_connection_log('Measurement started (UART, 261)')
                        except Exception:
                            pass
                        return
                    except Exception:
                        pass
                # Fallback: legacy start without args
                self._device.start_measurement()
                self._measurement_started = True
                try:
                    write_connection_log('Measurement started (SHDLC)')
                except Exception:
                    pass
        except Exception:
            pass

    def _read_measurement(self):
        if self._device is None:
            return None
        try:
            # Respect warm-up period after (re)connect
            if self._warmup_deadline and time.time() < self._warmup_deadline:
                time.sleep(0.2)
                return None
            # Legacy SHDLC driver path
            if hasattr(self._device, 'read_measured_values') or hasattr(self._device, 'read_measured_value'):
                for _ in range(10):
                    try:
                        if hasattr(self._device, 'is_data_ready') and self._device.is_data_ready():
                            break
                    except Exception:
                        break
                    time.sleep(0.1)

                try:
                    if hasattr(self._device, 'read_measured_values'):
                        m = self._device.read_measured_values()
                    else:
                        m = self._device.read_measured_value()
                    try:
                        self._debug_sample_counter += 1
                        if self._debug_sample_counter % 10 == 1:
                            write_connection_log(f"Legacy raw measurement: {m!r} ({type(m).__name__})")
                    except Exception:
                        pass
                except Exception as e:
                    try:
                        write_connection_log(f"Legacy read exception: {e!r}")
                    except Exception:
                        pass
                    return None

                # Extract PM values from various possible return formats
                def _to_float(x):
                    try:
                        return float(x)
                    except Exception:
                        return float('nan')

                pm1 = pm25 = pm4 = pm10 = float('nan')
                try:
                    # Object with attributes
                    if hasattr(m, 'pm1p0') or hasattr(m, 'pm2p5'):
                        pm1 = _to_float(getattr(m, 'pm1p0', float('nan')))
                        pm25 = _to_float(getattr(m, 'pm2p5', float('nan')))
                        pm4 = _to_float(getattr(m, 'pm4p0', float('nan')))
                        pm10 = _to_float(getattr(m, 'pm10p0', float('nan')))
                    # Dict-like
                    elif isinstance(m, dict):
                        pm1 = _to_float(m.get('pm1p0', float('nan')))
                        pm25 = _to_float(m.get('pm2p5', float('nan')))
                        pm4 = _to_float(m.get('pm4p0', float('nan')))
                        pm10 = _to_float(m.get('pm10p0', float('nan')))
                    # Tuple/list direct 4-tuple
                    elif isinstance(m, (list, tuple)) and len(m) >= 4 and not (
                        len(m) == 3 and isinstance(m[0], (list, tuple))
                    ):
                        pm1, pm25, pm4, pm10 = _to_float(m[0]), _to_float(m[1]), _to_float(m[2]), _to_float(m[3])
                    # SHDLC variant: ( (mc1, mc2.5, mc4, mc10), (nc...), typical_size )
                    elif isinstance(m, (list, tuple)) and len(m) >= 1 and isinstance(m[0], (list, tuple)) and len(m[0]) >= 4:
                        pm1, pm25, pm4, pm10 = _to_float(m[0][0]), _to_float(m[0][1]), _to_float(m[0][2]), _to_float(m[0][3])
                except Exception as e:
                    try:
                        write_connection_log(f"Legacy parse exception: {e!r}, value={m!r}")
                    except Exception:
                        pass

                try:
                    self._debug_sample_counter += 1
                    if self._debug_sample_counter % 10 == 1:
                        write_connection_log(f"Sample legacy parsed mc: pm1={pm1}, pm2.5={pm25}, pm4={pm4}, pm10={pm10}")
                except Exception:
                    pass

                return Sample(ts=datetime.now(), pm1=pm1, pm25=pm25, pm4=pm4, pm10=pm10)

            # Modern UART driver path
            if hasattr(self._device, 'read_measurement_values_uint16'):
                try:
                    # If device exposes data_ready or similar, poll briefly
                    for _ in range(20):
                        try:
                            if hasattr(self._device, 'is_data_ready') and self._device.is_data_ready():
                                break
                        except Exception:
                            break
                        time.sleep(0.1)
                    values = self._device.read_measurement_values_uint16()
                    mc_1p0, mc_2p5, mc_4p0, mc_10p0 = float(values[0]), float(values[1]), float(values[2]), float(values[3])
                    try:
                        # Heuristic: some UART firmwares return fixed-point values (e.g. 5000 for 0.5 µg/m³).
                        # If readings look unrealistically large, downscale by 1e5 to land in typical µg/m³ range.
                        max_val = max(mc_1p0, mc_2p5, mc_4p0, mc_10p0)
                        if max_val > 1000.0:
                            mc_1p0 /= 100000.0
                            mc_2p5 /= 100000.0
                            mc_4p0 /= 100000.0
                            mc_10p0 /= 100000.0
                    except Exception:
                        pass
                    # Debug log every ~10th sample to verify values
                    try:
                        self._debug_sample_counter += 1
                        if self._debug_sample_counter % 10 == 1:
                            write_connection_log(f"Sample UART mc_2p5={mc_2p5}")
                    except Exception:
                        pass
                    return Sample(
                        ts=datetime.now(),
                        pm1=mc_1p0,
                        pm25=mc_2p5,
                        pm4=mc_4p0,
                        pm10=mc_10p0,
                    )
                except Exception as e:
                    try:
                        write_connection_log(f"UART read exception: {e!r}")
                    except Exception:
                        pass
                    return None

            return None
        except Exception:
            return None

    def _run(self):
        backoff_s = 1.0
        while not self._stop_event.is_set():
            if self._device is None:
                ok = self._connect()
                if not ok:
                    time.sleep(min(10.0, backoff_s))
                    backoff_s = min(10.0, backoff_s * 1.5)
                    continue
                else:
                    backoff_s = 1.0
                    self._consecutive_read_failures = 0

            if self._paused.is_set():
                time.sleep(0.2)
                continue

            sample = self._read_measurement()
            if sample is None:
                # Allow a few failures before forcing reconnect
                self._consecutive_read_failures += 1
                if self._consecutive_read_failures >= self._max_consecutive_failures:
                    self._disconnect()
                    self._consecutive_read_failures = 0
                    self._measurement_started = False
                    time.sleep(1.0)
                    continue
                else:
                    # Try to ensure measurement is running and retry soon
                    try:
                        self._ensure_started()
                    except Exception:
                        pass
                    time.sleep(0.5)
                    continue

            if self.on_sample_callback:
                try:
                    self.on_sample_callback(sample)
                except Exception:
                    pass

            # Reset failure counter on success
            self._consecutive_read_failures = 0
            self._last_sample_time = datetime.now()

            # Sleep remaining time of the period
            end_time = time.time() + self.sample_period_s
            while time.time() < end_time and not self._stop_event.is_set() and not self._paused.is_set():
                time.sleep(0.05)


class CSVLogger:
    def __init__(self, log_dir):
        self.log_dir = log_dir
        ensure_dir(self.log_dir)

    def _file_for_today(self):
        day = datetime.now().strftime('%Y-%m-%d')
        return os.path.join(self.log_dir, f'sps30_{day}.csv')

    def append(self, sample):
        path = self._file_for_today()
        exists = os.path.isfile(path)
        try:
            with open(path, 'a', newline='') as f:
                writer = csv.writer(f)
                if not exists:
                    writer.writerow(['timestamp', 'pm1', 'pm2_5', 'pm4', 'pm10'])
                writer.writerow([
                    sample.ts.isoformat(timespec='seconds'),
                    f'{sample.pm1:.3f}', f'{sample.pm25:.3f}', f'{sample.pm4:.3f}', f'{sample.pm10:.3f}'
                ])
        except Exception:
            pass


class DataStore:
    def __init__(self):
        # Keep all samples for current session; UI filters to windows
        self.samples = deque()
        self.lock = threading.Lock()

    def add(self, sample):
        with self.lock:
            self.samples.append(sample)
            # Trim very old data (older than 7 days) to avoid unbounded growth
            cutoff = datetime.now() - timedelta(days=7)
            while self.samples and self.samples[0].ts < cutoff:
                self.samples.popleft()

    def get_window(self, window_td):
        now = datetime.now()
        with self.lock:
            if window_td is None:
                return list(self.samples)
            cutoff = now - window_td
            return [s for s in self.samples if s.ts >= cutoff]


class Dashboard(tk.Toplevel):
    def __init__(self, master, datastore, title, is_connected_fn):
        super().__init__(master)
        self.withdraw()  # start hidden
        self.title(title)
        self.protocol('WM_DELETE_WINDOW', self._on_close)
        self.resizable(True, True)
        self.datastore = datastore
        self.is_connected_fn = is_connected_fn

        self._build_ui()
        self._update_ui_scheduled = False

    def show(self):
        self.deiconify()
        self.lift()
        self.after(50, self._update_ui)

    def _on_close(self):
        self.withdraw()

    def _build_ui(self):
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True)

        # Header row: connection indicator and big current PM2.5 value
        header = ttk.Frame(container)
        header.pack(fill=tk.X, padx=8, pady=(8, 4))

        self.conn_label = tk.Label(header, text='● Disconnected', font=('Segoe UI', 10, 'bold'), fg='#e74c3c')
        self.conn_label.pack(side=tk.LEFT, padx=(0, 16))

        self.now_big_label = tk.Label(header, text='PM2.5: - µg/m³', font=('Segoe UI', 14, 'bold'))
        self.now_big_label.pack(side=tk.LEFT)

        self.tabs = ttk.Notebook(container)
        self.tabs.pack(fill=tk.BOTH, expand=True)

        self.windows = [
            ('Last 1h', timedelta(hours=1)),
            ('Last 3h', timedelta(hours=3)),
            ('Last 12h', timedelta(hours=12)),
            ('Last 24h', timedelta(hours=24)),
            ('All time', None),
        ]

        self.tab_frames = {}
        self.stats_labels = {}
        self.figures = {}
        self.axes = {}
        self.lines = {}

        for label, td in self.windows:
            frame = ttk.Frame(self.tabs)
            self.tabs.add(frame, text=label)
            self.tab_frames[label] = frame

            # Stats row
            stats_frame = ttk.Frame(frame)
            stats_frame.pack(fill=tk.X, padx=8, pady=6)
            self.stats_labels[label] = {
                'now': ttk.Label(stats_frame, text='Now: -  -  -  -'),
                'avg': ttk.Label(stats_frame, text='Avg: -  -  -  -'),
                'max': ttk.Label(stats_frame, text='Max: -  -  -  -'),
            }
            self.stats_labels[label]['now'].pack(side=tk.LEFT, padx=(0, 16))
            self.stats_labels[label]['avg'].pack(side=tk.LEFT, padx=(0, 16))
            self.stats_labels[label]['max'].pack(side=tk.LEFT)

            # Plot
            plot_frame = ttk.Frame(frame)
            plot_frame.pack(fill=tk.BOTH, expand=True)

            fig = Figure(figsize=(7, 3.5), dpi=100)
            ax = fig.add_subplot(111)
            ax.set_ylabel('µg/m³')
            ax.grid(True, linestyle='--', alpha=0.3)
            try:
                ax.set_ylim(bottom=0.0)
            except Exception:
                pass
            try:
                if _FormatStrFormatter is not None:
                    ax.yaxis.set_major_formatter(_FormatStrFormatter('%.4f'))
            except Exception:
                pass
            self.figures[label] = fig
            self.axes[label] = ax

            canvas = FigureCanvasTkAgg(fig, master=plot_frame)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill=tk.BOTH, expand=True)

            # Initialize empty lines
            line_pm1, = ax.plot([], [], label='PM1.0')
            line_pm25, = ax.plot([], [], label='PM2.5')
            line_pm4, = ax.plot([], [], label='PM4')
            line_pm10, = ax.plot([], [], label='PM10')
            ax.legend(loc='upper right')
            self.lines[label] = (line_pm1, line_pm25, line_pm4, line_pm10, canvas)

    def _format_stats(self, samples):
        def fmt(v):
            if v is None or (isinstance(v, float) and (v != v)):
                return '-'
            return f'{v:.4f}'

        if not samples:
            return 'Now: - / - / - / -', 'Avg: - / - / - / -', 'Max: - / - / - / -'
        last = samples[-1]
        pm1s = [s.pm1 for s in samples]
        pm25s = [s.pm25 for s in samples]
        pm4s = [s.pm4 for s in samples]
        pm10s = [s.pm10 for s in samples]
        now_str = f"Now: {fmt(last.pm1)} / {fmt(last.pm25)} / {fmt(last.pm4)} / {fmt(last.pm10)}"
        avg_str = f"Avg: {fmt(_avg(pm1s))} / {fmt(_avg(pm25s))} / {fmt(_avg(pm4s))} / {fmt(_avg(pm10s))}"
        max_str = f"Max: {fmt(max(pm1s))} / {fmt(max(pm25s))} / {fmt(max(pm4s))} / {fmt(max(pm10s))}"
        return now_str, avg_str, max_str

    def _update_tab(self, label, window_td):
        samples = self.datastore.get_window(window_td)
        ax = self.axes[label]
        line_pm1, line_pm25, line_pm4, line_pm10, canvas = self.lines[label]

        if not samples:
            for line in (line_pm1, line_pm25, line_pm4, line_pm10):
                line.set_data([], [])
            ax.relim()
            ax.autoscale_view()
            canvas.draw_idle()
            now_str, avg_str, max_str = self._format_stats(samples)
            self.stats_labels[label]['now'].configure(text=now_str)
            self.stats_labels[label]['avg'].configure(text=avg_str)
            self.stats_labels[label]['max'].configure(text=max_str)
            return

        xs = [s.ts for s in samples]
        ys1 = [s.pm1 for s in samples]
        ys25 = [s.pm25 for s in samples]
        ys4 = [s.pm4 for s in samples]
        ys10 = [s.pm10 for s in samples]

        line_pm1.set_data(xs, ys1)
        line_pm25.set_data(xs, ys25)
        line_pm4.set_data(xs, ys4)
        line_pm10.set_data(xs, ys10)

        ax.relim()
        ax.autoscale_view()
        ax.figure.autofmt_xdate()
        canvas.draw_idle()

        now_str, avg_str, max_str = self._format_stats(samples)
        self.stats_labels[label]['now'].configure(text=now_str)
        self.stats_labels[label]['avg'].configure(text=avg_str)
        self.stats_labels[label]['max'].configure(text=max_str)

    def _update_ui(self):
        # Update header connection state
        connected = False
        try:
            connected = bool(self.is_connected_fn())
        except Exception:
            connected = False
        if connected:
            self.conn_label.configure(text='● Connected', fg='#2ecc71')
        else:
            self.conn_label.configure(text='● Disconnected', fg='#e74c3c')

        # Update header current PM2.5 value
        all_samples = self.datastore.get_window(None)
        if all_samples:
            last = all_samples[-1]
            try:
                val = last.pm25
                if val == val:
                    self.now_big_label.configure(text=f'PM2.5: {val:.4f} µg/m³')
                else:
                    self.now_big_label.configure(text='PM2.5: - µg/m³')
            except Exception:
                self.now_big_label.configure(text='PM2.5: - µg/m³')
        else:
            self.now_big_label.configure(text='PM2.5: - µg/m³')

        for label, td in self.windows:
            self._update_tab(label, td)
        # Schedule next update
        self.after(1000, self._update_ui)


def _avg(values):
    vals = [v for v in values if v == v]
    if not vals:
        return float('nan')
    return sum(vals) / len(vals)


class App:
    def __init__(self):
        self.datastore = DataStore()
        self.csv_logger = CSVLogger(CONFIG.log_dir)
        self.sample_queue = queue.Queue()
        self.reader = SPS30Reader(CONFIG.uart_port, CONFIG.sample_period_s, self._on_sample)

        # Tk root (hidden owner for dashboard)
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title(CONFIG.app_name)
        self.root.protocol('WM_DELETE_WINDOW', self._on_root_close)

        # Dashboard window
        self.dashboard = Dashboard(self.root, self.datastore, title=f"{CONFIG.app_name} – Dashboard", is_connected_fn=self.reader.is_connected)

        # Tray icon
        self.icon = None
        self._build_tray_icon()

        # Periodic pump from sampling thread → UI
        self._pump_ui()

    def _on_sample(self, sample):
        try:
            self.sample_queue.put_nowait(sample)
        except Exception:
            pass

    def _pump_ui(self):
        drained = 0
        while True:
            try:
                s = self.sample_queue.get_nowait()
            except queue.Empty:
                break
            self.datastore.add(s)
            self.csv_logger.append(s)
            drained += 1
        # Schedule again quickly if there was new data; otherwise slower
        delay = 200 if drained else 500
        self.root.after(delay, self._pump_ui)

    def _build_tray_icon(self):
        if pystray is None:
            return
        image = load_logo_pil(desired_size=64) or create_tray_icon_image()

        def on_open(icon, item):
            self.open_dashboard()

        def on_pause_resume(icon, item):
            if self.reader.is_paused():
                self.reader.resume()
            else:
                self.reader.pause()
            # Update the menu text dynamically by recreating the menu
            self._rebuild_tray_menu()

        def on_quit(icon, item):
            self.quit()

        self.menu_open = Item('Open dashboard', on_open, default=True)
        self.menu_pause = Item(self._pause_menu_text(), on_pause_resume)
        self.menu_ports = Item('Ports', self._build_ports_menu())
        self.menu_test = Item('Connection test', self._on_connection_test)
        self.menu_quit = Item('Quit', on_quit)

        self.icon = pystray.Icon(
            'sps30_tray_logger',
            image,
            CONFIG.app_name,
            TrayMenu(
                self.menu_open,
                self.menu_pause,
                self.menu_ports,
                self.menu_test,
                Item('—', None, enabled=False),
                self.menu_quit,
            ),
        )

    def _rebuild_tray_menu(self):
        if self.icon is None:
            return
        image = self.icon.icon

        def on_open(icon, item):
            self.open_dashboard()

        def on_pause_resume(icon, item):
            if self.reader.is_paused():
                self.reader.resume()
            else:
                self.reader.pause()
            self._rebuild_tray_menu()

        def on_quit(icon, item):
            self.quit()

        self.menu_open = Item('Open dashboard', on_open, default=True)
        self.menu_pause = Item(self._pause_menu_text(), on_pause_resume)
        self.menu_ports = Item('Ports', self._build_ports_menu())
        self.menu_test = Item('Connection test', self._on_connection_test)
        self.menu_quit = Item('Quit', on_quit)

        self.icon.menu = TrayMenu(
            self.menu_open,
            self.menu_pause,
            self.menu_ports,
            self.menu_test,
            Item('—', None, enabled=False),
            self.menu_quit,
        )
        # Force redraw of tray icon context (workaround)
        self.icon.visible = True

    def _pause_menu_text(self):
        return 'Resume sampling' if self.reader.is_paused() else 'Pause sampling'

    def open_dashboard(self):
        self.dashboard.show()

    # ----- Tray: Ports submenu -----
    def _build_ports_menu(self):
        ports = detect_serial_ports()

        def is_selected(target):
            def _checked(_):
                cur = self.reader.uart_port
                return (cur is None and target is None) or (cur == target)
            return _checked

        def choose(target):
            def _handler(icon, item):
                self.reader.set_port(target)
                self._rebuild_tray_menu()
            return _handler

        def refresh(icon, item):
            self._rebuild_tray_menu()

        # Build dynamic items
        items = [
            Item('Auto (scan)', choose(None), checked=is_selected(None)),
            Item('Refresh list', refresh),
            Item('—', None, enabled=False),
        ]
        for p in ports:
            items.append(Item(p, choose(p), checked=is_selected(p)))

        return TrayMenu(*items)

    def _notify(self, text):
        try:
            if self.icon is not None and text:
                self.icon.title = f"{CONFIG.app_name} – {text}"
        except Exception:
            pass

    def _on_connection_test(self, icon=None, item=None):
        t = threading.Thread(target=self._run_connection_test, daemon=True)
        t.start()

    def _run_connection_test(self):
        write_connection_log('--- Connection test started ---')
        if ShdlcSerialPort is None or ShdlcConnection is None or Sps30ShdlcDevice is None:
            write_connection_log('Drivers not available: sensirion_shdlc_driver / sensirion_shdlc_sps')
            self._notify('Drivers missing')
            return

        # Temporarily release the serial port from the sampling thread
        was_paused = self.reader.is_paused()
        try:
            self.reader.pause()
            # Ensure port is closed so test can open it
            try:
                self.reader._disconnect()
            except Exception:
                pass

            ports_to_try = []
            if self.reader.uart_port:
                ports_to_try = [self.reader.uart_port]
            else:
                ports_to_try = [p for p in detect_serial_ports() if p.upper().startswith('COM')]
            write_connection_log(f'Ports to try: {ports_to_try}')
            for port_name in ports_to_try:
                port = None
                try:
                    write_connection_log(f'Trying {port_name} at 115200 8N1...')
                    port = ShdlcSerialPort(port_name, baudrate=115200)
                    conn = ShdlcConnection(port)
                    dev = Sps30ShdlcDevice(conn, slave_address=0)

                    # Try to start measurement if supported
                    try:
                        if hasattr(dev, 'start_measurement'):
                            dev.start_measurement()
                    except Exception:
                        pass

                    time.sleep(0.5)

                    # Prefer reading serial number; fall back to a measurement call if available
                    ok = False
                    serial_val = None
                    serial_methods = [
                        'device_information_serial_number',
                        'get_serial_number',
                        'serial_number',
                        'device_information',
                    ]
                    for mname in serial_methods:
                        try:
                            if hasattr(dev, mname):
                                val = getattr(dev, mname)()
                                serial_val = str(val)
                                ok = True
                                break
                        except Exception as e:
                            write_connection_log(f'{mname} threw: {e!r}')

                    if not ok and hasattr(dev, 'read_measured_values'):
                        try:
                            _ = dev.read_measured_values()
                            ok = True
                        except Exception as e:
                            write_connection_log(f'read_measured_values threw: {e!r}')

                    try:
                        port.close()
                    except Exception:
                        pass

                    if ok:
                        msg = f'OK on {port_name}' + (f' (SN: {serial_val})' if serial_val else '')
                        write_connection_log(msg)
                        self._notify(f'Connection OK on {port_name}')
                        return
                    else:
                        write_connection_log(f'No response on {port_name}')
                except Exception as e:
                    write_connection_log(f'Open/communicate failed for {port_name}: {e!r}')
                    try:
                        if port is not None:
                            port.close()
                    except Exception:
                        pass
        finally:
            # Resume sampling if it wasn't paused before
            if not was_paused:
                try:
                    self.reader.resume()
                except Exception:
                    pass

        self._notify('Connection failed')
        write_connection_log('No working port found')

    def run(self):
        # Start sampling
        self.reader.start()

        # Start tray icon in a dedicated thread so that Tk mainloop remains responsive
        if self.icon is not None:
            t = threading.Thread(target=self.icon.run, kwargs={'setup': self._tray_setup}, daemon=True)
            t.start()

        # Show dashboard on first run for discoverability
        self.root.after(200, self.open_dashboard)
        self.root.mainloop()

    def _tray_setup(self, icon):
        icon.visible = True

    def _on_root_close(self):
        # Root stays hidden; dashboard has its own close handler
        self.dashboard.withdraw()

    def quit(self):
        try:
            if self.icon is not None:
                self.icon.visible = False
                # pystray stops when visible False + stop called
                try:
                    self.icon.stop()
                except Exception:
                    pass
        finally:
            pass
        try:
            self.reader.stop()
        finally:
            pass
        try:
            self.dashboard.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        os._exit(0)


def create_tray_icon_image(size=64, color_fg=(46, 204, 113), color_bg=(33, 33, 33)):
    if Image is None:
        return None
    image = Image.new('RGBA', (size, size), color_bg + (255,))
    draw = ImageDraw.Draw(image)
    # Outer circle
    margin = int(size * 0.12)
    draw.ellipse([margin, margin, size - margin, size - margin], fill=color_fg + (255,), outline=(255, 255, 255, 200), width=max(1, size // 32))
    # Four small dots symbolizing PM sizes
    cx = cy = size // 2
    r = size // 14
    offsets = [(-r*2, 0), (0, -r*2), (r*2, 0), (0, r*2)]
    for dx, dy in offsets:
        draw.ellipse([cx + dx - r, cy + dy - r, cx + dx + r, cy + dy + r], fill=(255, 255, 255, 220))
    return image


def main():
    if sys.platform != 'win32':
        print('This script is intended for Windows. Proceeding anyway...')
    app = App()
    app.run()


if __name__ == '__main__':
    main()


