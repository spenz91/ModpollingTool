#!/usr/bin/env python
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial.tools.list_ports
import subprocess
import threading
import os
import queue
import webbrowser  # For opening URLs
import winreg     # For registry access
import re         # For regex matching

class ModpollingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ModPolling Tool")
        self.modpoll_process = None
        self.is_polling = False
        self.log_queue = queue.Queue()

        # Equipment settings
        self.equipment_settings = {
            "ADAM": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "AERMEC": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "AKCC250": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "AKCC350": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "AKCC55": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "AKCC550A": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "AKPC420": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "ANYBUS": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "ATLANTIUM": {"baudrate": "115200", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "AWD3": {"baudrate": "19200", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "BELIMO": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CAREL": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CIAT2": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CIRCUTOR": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "CLIMAVENTA": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CLIVET": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CORRIGO": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CORRIGO34": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CVM10": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CVM96": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "CVMC": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "DAIKIN": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "none"},  # '0A' treated as 'none'
            "DIXELL": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "DUPLINE": {"baudrate": "115200", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EDMK": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "EM100": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM21": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM210": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM23": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM24": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM24TCP": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "EM26": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM270": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM330": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM4": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "EM540": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "EW": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "FLAKTWOODS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "FLEXIT": {"baudrate": "9600", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "GREENCOOL": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "GRFOS": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "HECU": {"baudrate": "19200", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "IEM3250": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "INEPRO": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "INTESIS": {"baudrate": "9600", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "IR33PLUS": {"baudrate": "19200", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "IVPRODUKT": {"baudrate": "9600", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "IWT": {"baudrate": "57600", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "KAMSTRUP": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "LANDIS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "LDS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "LEMMENS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "LIEBHERR": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "MKD": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "MODBUS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "NEMO96": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "NETAVENT": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "NOVAGG": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "OJEXHAUST": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "PIIGAB": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "PMGOLD": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "POWERTAG": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "PR100T": {"baudrate": "19200", "stop_bits": "2", "data_bits": "8", "parity": "none"},
            "QALCOSONIC": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "REGIN": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "REGINRCF": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "SCHNEIDER": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "SLV": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "SOLARLOG": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "SWEGON": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "TROX": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "UH50": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "UNISAB3": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "VENT": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "VIESSMANN": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "WM14": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "WTRANS": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
        }

        # Flag to track equipment pane visibility
        self.equipment_pane_visible = False

        # For numbering specific log messages
        self.message_counts = {}

        # Initialize blinking variables
        self.blinking = False
        self.blink_job = None
        self.current_color = None
        self.blink_state = False  # To toggle between color and background
        self.blink_count = 0  # To track the number of blinks

        # Define colors
        self.GREEN_COLOR = "#00FF00"
        self.GREEN_DIM = "#66FF66"
        self.YELLOW_COLOR = "#FFFF00"
        self.YELLOW_DIM = "#FFFF66"
        self.RED_COLOR = "#FF0000"
        self.RED_DIM = "#FF6666"

        self.create_widgets()
        self.update_log()

    def create_widgets(self):
        # Configure grid layout for root
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        # Main Frame
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="NSEW")
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)  # Equipment
        main_frame.columnconfigure(1, weight=3)  # Settings
        main_frame.columnconfigure(2, weight=2)  # Log

        # Equipment Frame (Column 0)
        self.frame_equipment = ttk.LabelFrame(main_frame, text="Equipment")
        self.frame_equipment.grid(row=0, column=0, sticky="NSEW", padx=5, pady=5)
        self.frame_equipment.rowconfigure(0, weight=1)
        self.frame_equipment.columnconfigure(0, weight=1)

        # Create a frame to hold the Listbox and Scrollbar
        listbox_frame = ttk.Frame(self.frame_equipment)
        listbox_frame.grid(row=0, column=0, sticky="NSEW", padx=5, pady=5)

        # Equipment Listbox
        self.listbox_equipment = tk.Listbox(listbox_frame, height=20)
        self.listbox_equipment.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar for the Listbox
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox_equipment.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_equipment.config(yscrollcommand=scrollbar.set)

        # Populate equipment list in alphabetical order
        for equipment in sorted(self.equipment_settings.keys()):
            self.listbox_equipment.insert(tk.END, equipment)

        # Bind double-click event and select button
        self.listbox_equipment.bind('<Double-1>', self.select_equipment)
        btn_select = ttk.Button(self.frame_equipment, text="Select", command=self.select_equipment)
        btn_select.grid(row=1, column=0, pady=(0, 10))

        # Settings Frame with Notebook for Basic and Advanced Tabs (Column 1)
        self.frame_settings = ttk.LabelFrame(main_frame, text="Settings")
        self.frame_settings.grid(row=0, column=1, sticky="NSEW", padx=5, pady=5)
        self.frame_settings.rowconfigure(0, weight=1)
        self.frame_settings.columnconfigure(0, weight=1)

        # Notebook Widget for Tabs
        self.settings_notebook = ttk.Notebook(self.frame_settings)
        self.settings_notebook.pack(expand=True, fill='both')

        # Basic Tab
        self.basic_tab = ttk.Frame(self.settings_notebook)
        self.settings_notebook.add(self.basic_tab, text="Basic")

        # Advanced Tab
        self.advanced_tab = ttk.Frame(self.settings_notebook)
        self.settings_notebook.add(self.advanced_tab, text="Advanced")

        # ---------------- Basic Settings Widgets ----------------
        # COM Port
        ttk.Label(self.basic_tab, text="COM Port:").grid(column=0, row=0, sticky="W", padx=5, pady=5, ipadx=5)
        self.cmb_comport = ttk.Combobox(self.basic_tab, state="normal", width=15)
        self.cmb_comport.grid(column=1, row=0, padx=5, pady=5, sticky="W")
        self.btn_refresh = ttk.Button(self.basic_tab, text="Refresh", command=self.refresh_comports)
        self.btn_refresh.grid(column=2, row=0, padx=5, pady=5, sticky="W")

        # Baudrate
        ttk.Label(self.basic_tab, text="Baudrate (-b):").grid(column=0, row=1, sticky="W", padx=5, pady=5, ipadx=5)
        self.cmb_baudrate = ttk.Combobox(
            self.basic_tab,
            state="normal",  # Allow manual entry
            width=15,
            values=("2400", "4800", "9600", "19200", "38400", "57600", "115200")
        )
        self.cmb_baudrate.current(2)  # Default to "9600"
        self.cmb_baudrate.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Parity
        ttk.Label(self.basic_tab, text="Parity (-p):").grid(column=0, row=2, sticky="W", padx=5, pady=5, ipadx=5)
        self.cmb_parity = ttk.Combobox(self.basic_tab, state="readonly", width=15, values=("none", "even", "odd"))
        self.cmb_parity.current(0)  # Default to "none"
        self.cmb_parity.grid(column=1, row=2, padx=5, pady=5, sticky="W")

        # Address
        ttk.Label(self.basic_tab, text="Address (-a):").grid(column=0, row=3, sticky="W", padx=5, pady=5, ipadx=5)
        self.entry_slave_address = ttk.Entry(self.basic_tab, width=17)
        self.entry_slave_address.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_slave_address.insert(0, "1")

        # ---------------- Advanced Settings Widgets ----------------
        # Data Bits
        ttk.Label(self.advanced_tab, text="Data Bits (-d):").grid(column=0, row=0, sticky="W", padx=5, pady=5)
        self.cmb_databits = ttk.Combobox(self.advanced_tab, state="readonly", width=15, values=("7", "8"))
        self.cmb_databits.current(1)  # Default to "8"
        self.cmb_databits.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        # Stop Bits
        ttk.Label(self.advanced_tab, text="Stop Bits (-s):").grid(column=0, row=1, sticky="W", padx=5, pady=5)
        self.cmb_stopbits = ttk.Combobox(self.advanced_tab, state="readonly", width=15, values=("1", "2"))
        self.cmb_stopbits.current(0)  # Default to "1"
        self.cmb_stopbits.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Start Reference
        ttk.Label(self.advanced_tab, text="Start Reference (-r):").grid(column=0, row=2, sticky="W", padx=5, pady=5)
        self.entry_start_reference = ttk.Entry(self.advanced_tab, width=17)
        self.entry_start_reference.grid(column=1, row=2, padx=5, pady=5, sticky="W")
        self.entry_start_reference.insert(0, "100")

        # Number of Registers
        ttk.Label(self.advanced_tab, text="Number of Registers (-c):").grid(column=0, row=3, sticky="W", padx=5, pady=5)
        self.entry_num_values = ttk.Entry(self.advanced_tab, width=17)
        self.entry_num_values.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_num_values.insert(0, "1")

        # Register Data Type
        ttk.Label(self.advanced_tab, text="Register Data Type (-t):").grid(column=0, row=4, sticky="W", padx=5, pady=5)
        self.entry_register_data_type = ttk.Entry(self.advanced_tab, width=17)
        self.entry_register_data_type.grid(column=1, row=4, padx=5, pady=5, sticky="W")
        self.entry_register_data_type.insert(0, "3")  # Default is "3" unless overridden

        # Modbus/TCP
        ttk.Label(self.advanced_tab, text="Modbus/TCP (-m tcp):").grid(column=0, row=5, sticky="W", padx=5, pady=5)
        self.entry_modbus_tcp = ttk.Entry(self.advanced_tab, width=17)
        self.entry_modbus_tcp.grid(column=1, row=5, padx=5, pady=5, sticky="W")

        # Adjust column weights in Advanced Tab for better layout
        self.advanced_tab.columnconfigure(0, weight=1)
        self.advanced_tab.columnconfigure(1, weight=3)

        # ---------------- Polling Buttons and Status Indicator ----------------
        btn_frame = ttk.Frame(self.frame_settings)
        btn_frame.pack(pady=10, fill='x', expand=True)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        btn_frame.columnconfigure(3, weight=1)  # For status indicator

        # Start Polling Button
        self.btn_start = ttk.Button(btn_frame, text="Start Polling", command=self.start_polling)
        self.btn_start.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Stop Polling Button
        self.btn_stop = ttk.Button(btn_frame, text="Stop Polling", command=self.stop_polling, state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        # Advanced Button
        self.btn_advanced = ttk.Button(btn_frame, text="Advanced", command=self.select_advanced_tab)
        self.btn_advanced.grid(row=0, column=2, padx=5, pady=5, sticky="e")

        # ---------------- Status Indicator ----------------
        self.status_canvas_size = 20  # Size of the canvas
        self.status_canvas = tk.Canvas(btn_frame, width=self.status_canvas_size, height=self.status_canvas_size, highlightthickness=0)
        self.status_canvas.grid(row=0, column=3, padx=10, pady=5, sticky="w")
        self.circle = self.status_canvas.create_oval(2, 2, self.status_canvas_size-2, self.status_canvas_size-2, fill="grey")

        # ---------------- Log Frame (Column 2) ----------------
        self.frame_log = ttk.LabelFrame(main_frame, text="Log")
        self.frame_log.grid(row=0, column=2, sticky="NSEW", padx=5, pady=5)
        self.frame_log.rowconfigure(0, weight=1)
        self.frame_log.columnconfigure(0, weight=1)

        self.txt_log = scrolledtext.ScrolledText(self.frame_log, width=80, height=25, state='disabled',
                                                  bg="black", fg="white", font=("Consolas", 10))
        self.txt_log.grid(column=0, row=0, sticky="NSEW")
        self.txt_log.tag_configure('error', foreground='red')
        self.txt_log.tag_configure('info', foreground='green')

        # Refresh COM ports
        self.refresh_comports()

        # Footer
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=1, column=0, sticky="ew")
        # (Additional footer widgets can be added here if needed)

    def select_equipment(self, event=None):
        # Get the selected equipment name from the listbox
        equipment = self.listbox_equipment.get(tk.ACTIVE)
        settings = self.equipment_settings.get(equipment)
        if settings:
            self.cmb_baudrate.set(settings.get("baudrate"))
            self.cmb_stopbits.set(settings.get("stop_bits"))
            self.cmb_databits.set(settings.get("data_bits"))
            self.cmb_parity.set(settings.get("parity"))
        # Set register data type based on equipment preset:
        if equipment == "EW":
            self.entry_register_data_type.delete(0, tk.END)
            self.entry_register_data_type.insert(0, "4")
        else:
            self.entry_register_data_type.delete(0, tk.END)
            self.entry_register_data_type.insert(0, "3")

    def refresh_comports(self):
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.cmb_comport['values'] = port_list
        if port_list:
            self.cmb_comport.current(0)

    def start_polling(self):
        # Build modpoll command based on the settings and start the polling process.
        # (Implementation details depend on your modpoll executable and parameters)
        messagebox.showinfo("Start Polling", "Polling started (functionality not fully implemented)")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")

    def stop_polling(self):
        # Stop the polling process.
        messagebox.showinfo("Stop Polling", "Polling stopped (functionality not fully implemented)")
        self.btn_stop.config(state="disabled")
        self.btn_start.config(state="normal")

    def update_log(self):
        # Update log text from the queue periodically.
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.txt_log.config(state="normal")
                self.txt_log.insert(tk.END, message + "\n")
                self.txt_log.see(tk.END)
                self.txt_log.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.update_log)

    def select_advanced_tab(self):
        self.settings_notebook.select(self.advanced_tab)

def main():
    root = tk.Tk()
    app = ModpollingTool(root)
    root.mainloop()

if __name__ == '__main__':
    main()
