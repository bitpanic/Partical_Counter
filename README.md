## SPS30 Tray Logger (Windows + Dashboard)

A Windows tray app that logs Sensirion SPS30 particulate matter readings and provides a built-in dashboard. Double-clicking the tray icon (or choosing "Open dashboard") opens a clean view with tabs for Last 1h, 3h, 12h, 24h, and All time, plotting PM1/PM2.5/PM4/PM10 and showing quick stats (now/avg/max).

### Features
- Tray icon with quick menu: Open dashboard, Pause/Resume sampling, Quit
- Auto-scan COM ports (COM3..COM40) or set a fixed port
- CSV logging per day in `./logs` folder
- Tkinter dashboard with Matplotlib plots and quick stats across time windows
- Single-file EXE build (no console) via PyInstaller

### Install (Windows)
```powershell
py -m pip install --upgrade pip
py -m pip install pystray pillow sensirion-shdlc-driver sensirion-shdlc-sps matplotlib pyinstaller
```

### Configure
Edit the top of `sps30_tray_logger_win.py` if desired:
```python
CONFIG.uart_port = "COM5"  # or leave None to auto-scan COM3..COM40
CONFIG.sample_period_s = 5.0
```

### Run
```powershell
py sps30_tray_logger_win.py
```
- The tray icon appears. Double-click it or pick "Open dashboard".
- CSV files are created in `./logs` named `sps30_YYYY-MM-DD.csv`.

### Build single-file EXE (no console)
```powershell
py -m PyInstaller --onefile --noconsole --name sps30-tray-logger sps30_tray_logger_win.py
```
Result: `./dist/sps30-tray-logger.exe`. Double-click to run from the tray.

### Build script and installer (recommended)
Option A: One-liner build + optional installer
```powershell
./build_installer.ps1
```
- Produces the portable EXE at `dist\sps30-tray-logger.exe`.
- If [Inno Setup](https://jrsoftware.org/isinfo.php) is installed (ISCC on PATH), also produces an installer at `installer\Output\SPS30TrayLoggerSetup.exe`.

Option B: Manual (no installer)
- Use the PyInstaller command above.

### Auto-start on login (optional)
Option A: Startup folder
1. Press Win+R, type `shell:startup`, press Enter.
2. Copy `dist\sps30-tray-logger.exe` (or a shortcut) into that folder.

Option B: Task Scheduler
1. Open Task Scheduler → Create Task.
2. Triggers: At log on.
3. Actions: Start a program → browse to `sps30-tray-logger.exe`.
4. Conditions: Uncheck "Start the task only if the computer is on AC power" if desired.
5. Settings: Allow task to be run on demand.

### Notes
- If no `CONFIG.uart_port` is set, the app scans COM3..COM40 for the SPS30.
- Close the dashboard window to hide it; sampling continues in the tray.
- Use the tray menu to pause/resume sampling.

### Troubleshooting
- If imports show as missing in your editor but you installed them, ensure you’re using the same Python interpreter as `py` and restart your editor.
- If the SPS30 is not found, verify the COM port in Device Manager and set `CONFIG.uart_port` accordingly.
- If the dashboard doesn’t show, ensure `matplotlib` backend is available and `tkinter` is present (included with most Windows Python builds).

### License
MIT


