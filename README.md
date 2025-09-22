## ModPolling Tool

[![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows) [![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/) ![GUI-Tkinter](https://img.shields.io/badge/GUI-Tkinter-ff6f00) ![Modbus](https://img.shields.io/badge/Modbus-RTU%2FTCP-0a84ff) ![Status](https://img.shields.io/badge/Status-Active-success) ![License](https://img.shields.io/badge/License-TBD-lightgrey)

A modern Windows GUI for quickly polling Modbus devices using the `modpoll` utility. It streamlines serial and TCP polling, provides sensible equipment presets, shows a live command preview, parses and color-codes responses, and can optionally fetch unit metadata from a local IWMAC MySQL database.

### Highlights 🚀
- **⚡ One-click polling** over serial (RTU) or TCP
- **🔧 Equipment presets** auto-fill baud, parity, data/stop bits
- **🧪 Live command preview** with on-the-fly edits
- **🟢🟡🔴 Status indicator** with green/yellow/red blink feedback
- **🧠 Smart log parsing**: suppresses boilerplate, counts attempts, highlights issues
- **🔎 COM port discovery** via pySerial and Windows Registry
- **🗃️ Units tab (optional)**: fetches COM/IP/baud/parity from local MySQL
- **🔢 Auto-handling of COM10+** ports using `\\.\COMn` format

<p align="center">
    <img width="900" src="https://i.ibb.co/hst67zM/Skjermbilde-2024-11-18-142520.png" alt="Material Bread logo">
</p>


On first run, the app will ensure `modpoll.exe` is available:
- By default it will download to `C:\iwmac\bin\modpoll.exe` from a GitHub release.
- If the folder does not exist, download if from here **[Download modpoll.exe](https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/modpoll.exe)**

Note: If you prefer to ship `modpoll.exe` inside your project folder and avoid any system-installed copy, see the configuration section below.


## Quick Start ⚡
## 📥 Download

### Latest Release
Download the latest executable directly from our releases:

**[Download ModPollingTool.exe](https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/ModPollingTool.exe)**

- **File Size**: ~10.6 MB
- **Dependencies**: None (standalone executable)

## 🚀 Installation

1. **Download** the `ModPollingTool.exe` file from the link above
2. **Extract** or place the file in your desired directory
3. **Run** the executable by double-clicking `ModPollingTool.exe`
4. **No installation required** - it's a portable application
2. Press “Refresh” to list COM ports (or type one manually).
3. Choose an equipment preset (left pane) or type settings manually.
4. For RTU: pick `COM` port, set Address and Baud/Parity as needed.
5. For TCP: enter `host[:port]` in “Modbus TCP/IP (-m tcp)”.
6. Verify the live “Command” preview.
7. Click “Start Polling”. Use “Stop Polling” to end.


## Using the App 🧭

### Serial (RTU) polling 🔌
- Select a `COM` port. COM ports ≥ 10 are automatically formatted as `\\.\COMn`.
- Set Address (-a), Baudrate (-b), Parity (-p). Advanced (-d data bits, -s stop bits, -r start reference, -c count, -t type) are available under the Advanced tab.
- Click “Start Polling”. The log shows attempts, parsed data lines like `[100]: 70`, and highlights timeouts or port errors.

### Modbus TCP 🌐
- Enter `host[:port]` in “Modbus TCP/IP (-m tcp)”.
- Address (-a), registers (-r, -c) and type (-t) still apply.
- Click “Start Polling”.

### Equipment presets 🧩
- Search and select an equipment model; the app fills typical baud/parity/data/stop settings.
- You can still override any field afterward.

### Command preview and custom arguments 🧪
- The “Command” field shows exactly what will be passed to `modpoll`.
- Edit the command inline for advanced scenarios (e.g., adding flags not exposed in the UI). The app will use your edited command when starting polling.

### Status indicator 🟢🟡🔴
- The circular indicator blinks while parsing responses:
  - Green: valid responses or non-fatal Modbus exceptions (function/data address/value)
  - Yellow: checksum errors
  - Red: timeouts or port/socket errors

### Units tab (optional, IWMAC) 🗃️
- “Get units data” queries the local database for unit details, then auto-populates a table with Unit ID/Name, Driver, Address, IP/COM, Baudrate, Parity.
- A toggle filters to “Modbus-supported units” (those with a modbus value).


## What the app runs under the hood 🔍
The app constructs a `modpoll` command and runs it without spawning a console window. Examples:

```bash
# Serial RTU example
modpoll COM3 -b9600 -pnone -a1 -r100 -c1 -t3

# TCP example
modpoll -m tcp 192.168.1.50:502 -b9600 -peven -d8 -s1 -a1 -r100 -c10 -t3
```

The log suppresses `modpoll` headers and focuses on data, attempts, and actionable errors (timeouts, serial port already open, checksum errors, etc.).


## Configuration ⚙️

### Use a bundled modpoll.exe 📦
By default the app manages its own copy of `modpoll.exe` at `C:\iwmac\bin\modpoll.exe` (downloading it if missing).

Place `modpoll.exe` at that location. This avoids any dependency on a system-installed copy and keeps the binary under versioned distribution.

## Troubleshooting 🧰
- **modpoll.exe not found**: Ensure the app can write to the configured path, or place `modpoll.exe` there manually. If downloading fails, check network and proxy settings.
- **Serial port already open**: Stop any service using the same COM port (e.g., IWMAC plant server) and try again.
- **Reply time-out**: Inspect wiring, address, baudrate, parity. Many devices require exact settings; try vendor presets.
- **Checksum error**: Usually wiring/noise or parity/baud mismatch. Re-check cable shields and ground.
- **COM port missing**: Click Refresh, or type it manually; verify device manager; for virtual ports, ensure the vendor tool created the port.


## Development notes 🧑‍💻
- GUI built with `tkinter` and `ttk`
- Serial ports via `pyserial` and Windows Registry scan
- Subprocess runs `modpoll.exe` with `CREATE_NO_WINDOW`, streaming stdout/stderr in real-time
- Log filtering avoids `modpoll` banners and emphasizes actionable lines


## License 📄
Consider adding a license (e.g., MIT) to clarify usage and redistribution.

## ModPolling Tool (Windows) 🪟

Modern Windows GUI for quickly testing Modbus devices using the Modpoll utility. Built with Python and Tkinter, with quality-of-life features for serial/TCP, command preview, status indicator, and a Units tab that reads plant data from MySQL.

### Features
- **GUI for Modpoll**: Start/stop polling without the command line.
- **Equipment presets**: Auto-applies baudrate, parity, bits, etc.
- **Command preview**: Live `modpoll` command string updates as you change settings.
- **Status indicator**: Color feedback for responses and errors.
- **Units tab**: Pulls unit metadata (Unit ID/Name, Driver/Address, Regulator, IP, COM) from MySQL.

### Modpoll binary
- The app runs a known `modpoll.exe` rather than whatever may exist in the system PATH. It ensures the binary exists and downloads it if missing.
- Intention: always use the bundled/known Modpoll binary rather than any other system-installed one [[memory:5703358]].


### Usage
1. **Select equipment** on the left. Preset communication parameters are applied automatically.
2. In the **Basic** tab, pick a COM port or enter a **Modbus TCP/IP** host:port in the Advanced tab field.
3. Optionally adjust Advanced settings (data bits, stop bits, start reference, count, type).
4. Click **Start Polling** to run. Output appears in the log; the circular indicator shows status.
5. Use **Units → Get units data** to fetch from MySQL. Columns auto-size and support scrolling.


### Troubleshooting
- **Serial port already open**: Stop any service using the COM port (e.g., plant server) and try again.
- **Port or socket open error / Reply time-out**: Verify wiring/IP/port, parity/baudrate, address, and that the device is reachable.



