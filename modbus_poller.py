import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import serial.tools.list_ports
import subprocess
import threading
import os
import queue
import webbrowser  # For opening URLs
import winreg  # For registry access
import re  # For regex matching
import urllib.request  # For downloading files
import urllib.error  # For handling download errors
import shlex  # For parsing command line arguments

class ModpollingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ModPolling Tool")
        self.modpoll_process = None
        self.is_polling = False
        self.log_queue = queue.Queue()
        
        # Default modpoll path
        self.modpoll_path = r"C:\iwmac\bin\modpoll.exe"
        
        # Ensure modpoll.exe exists, download if needed
        self.ensure_modpoll_exists()

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
        
        # Poll attempt tracking
        self.poll_attempt_counter = 0
        self.current_start_reference = None

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

    def ensure_modpoll_exists(self):
        r"""
        Check if modpoll.exe exists in C:\iwmac\bin, if not download it automatically.
        """
        if not os.path.exists(self.modpoll_path):
            # Create directory if it doesn't exist
            modpoll_dir = os.path.dirname(self.modpoll_path)
            if not os.path.exists(modpoll_dir):
                try:
                    os.makedirs(modpoll_dir, exist_ok=True)
                except Exception as e:
                    self.log_queue.put(('error', f"Failed to create directory {modpoll_dir}: {e}"))
                    return
            
            # Download modpoll.exe
            self.download_modpoll()

    def download_modpoll(self):
        """
        Download modpoll.exe from the specified URL.
        """
        url = "https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/modpoll.exe"
        
        def download_thread():
            try:
                self.log_queue.put(('info', "Downloading modpoll.exe..."))
                
                # Download the file
                urllib.request.urlretrieve(url, self.modpoll_path)
                
                self.log_queue.put(('info', f"Successfully downloaded modpoll.exe to {self.modpoll_path}"))
                    
            except urllib.error.URLError as e:
                error_msg = f"Failed to download modpoll.exe: {e}"
                self.log_queue.put(('error', error_msg))
                messagebox.showerror("Download Error", error_msg)
            except Exception as e:
                error_msg = f"Unexpected error downloading modpoll.exe: {e}"
                self.log_queue.put(('error', error_msg))
                messagebox.showerror("Download Error", error_msg)
        
        # Run download in a separate thread to avoid blocking the UI
        threading.Thread(target=download_thread, daemon=True).start()

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
        self.frame_equipment.rowconfigure(1, weight=1)  # Changed from 0 to 1 to make room for search
        self.frame_equipment.columnconfigure(0, weight=1)

        # Search Frame (Row 0)
        search_frame = ttk.Frame(self.frame_equipment)
        search_frame.grid(row=0, column=0, sticky="EW", padx=5, pady=(5, 0))
        search_frame.columnconfigure(0, weight=1)

        # Search Label and Entry
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_equipment)  # Bind to search changes
        self.entry_search = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        self.entry_search.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Clear search button
        btn_clear_search = ttk.Button(search_frame, text="X", width=3, command=self.clear_search)
        btn_clear_search.pack(side=tk.RIGHT)

        # Create a frame to hold the Listbox and Scrollbar (Row 1)
        listbox_frame = ttk.Frame(self.frame_equipment)
        listbox_frame.grid(row=1, column=0, sticky="NSEW", padx=5, pady=5)

        # Equipment Listbox
        self.listbox_equipment = tk.Listbox(listbox_frame, height=18)  # Reduced height to make room for search
        self.listbox_equipment.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar for the Listbox
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox_equipment.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure the Listbox to work with the Scrollbar
        self.listbox_equipment.config(yscrollcommand=scrollbar.set)

        # Store original equipment list for filtering
        self.all_equipment = sorted(self.equipment_settings.keys())
        
        # Populate equipment list in alphabetical order
        for equipment in self.all_equipment:
            self.listbox_equipment.insert(tk.END, equipment)

        # Bind single-click event
        self.listbox_equipment.bind('<Button-1>', self.select_equipment)
        
        # Bind Enter key to select equipment
        self.listbox_equipment.bind('<Return>', self.select_equipment)
        
        # Bind Enter key in search field to select first item
        self.entry_search.bind('<Return>', self.select_first_equipment)
        
        # Bind Ctrl+F to focus search field
        self.root.bind('<Control-f>', self.focus_search)
        
        # Focus search field on startup
        self.entry_search.focus_set()

        # Select Button below the Listbox
        btn_select = ttk.Button(self.frame_equipment, text="Select", command=self.select_equipment)
        btn_select.grid(row=2, column=0, pady=(0, 10))  # Changed from row=1 to row=2

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
        ttk.Label(self.basic_tab, text="COM Port:").grid(
            column=0, row=0, sticky="W", padx=5, pady=5, ipadx=5
        )
        self.cmb_comport = ttk.Combobox(self.basic_tab, state="normal", width=15)
        self.cmb_comport.grid(column=1, row=0, padx=5, pady=5, sticky="W")
        self.btn_refresh = ttk.Button(
            self.basic_tab, text="Refresh", command=self.refresh_comports
        )
        self.btn_refresh.grid(column=2, row=0, padx=5, pady=5, sticky="W")

        # Baudrate
        ttk.Label(self.basic_tab, text="Baudrate (-b):").grid(
            column=0, row=1, sticky="W", padx=5, pady=5, ipadx=5
        )
        self.cmb_baudrate = ttk.Combobox(
            self.basic_tab,
            state="normal",  # Changed from "readonly" to "normal" to allow manual entry
            width=15,
            values=("2400", "4800", "9600", "19200", "38400", "57600", "115200"),  # Removed "Custom"
        )
        self.cmb_baudrate.current(2)  # Default to "9600"
        self.cmb_baudrate.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Parity
        ttk.Label(self.basic_tab, text="Parity (-p):").grid(
            column=0, row=2, sticky="W", padx=5, pady=5, ipadx=5
        )
        self.cmb_parity = ttk.Combobox(
            self.basic_tab, state="readonly", width=15, values=("none", "even", "odd")
        )
        self.cmb_parity.current(0)  # Default to "none"
        self.cmb_parity.grid(column=1, row=2, padx=5, pady=5, sticky="W")

        # Address
        ttk.Label(self.basic_tab, text="Address (-a):").grid(
            column=0, row=3, sticky="W", padx=5, pady=5, ipadx=5
        )
        self.entry_slave_address = ttk.Entry(self.basic_tab, width=17)
        self.entry_slave_address.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_slave_address.insert(0, "1")

        # ---------------- Advanced Settings Widgets ----------------

        # Data Bits
        ttk.Label(self.advanced_tab, text="Data Bits (-d):").grid(
            column=0, row=0, sticky="W", padx=5, pady=5
        )
        self.cmb_databits = ttk.Combobox(
            self.advanced_tab, state="readonly", width=15, values=("7", "8")
        )
        self.cmb_databits.current(1)  # Default to "8"
        self.cmb_databits.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        # Stop Bits
        ttk.Label(self.advanced_tab, text="Stop Bits (-s):").grid(
            column=0, row=1, sticky="W", padx=5, pady=5
        )
        self.cmb_stopbits = ttk.Combobox(
            self.advanced_tab, state="readonly", width=15, values=("1", "2")
        )
        self.cmb_stopbits.current(0)  # Default to "1"
        self.cmb_stopbits.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Start Reference
        ttk.Label(self.advanced_tab, text="Start Reference (-r):").grid(
            column=0, row=2, sticky="W", padx=5, pady=5
        )
        self.entry_start_reference = ttk.Entry(self.advanced_tab, width=17)
        self.entry_start_reference.grid(column=1, row=2, padx=5, pady=5, sticky="W")
        self.entry_start_reference.insert(0, "100")

        # Number of Registers
        ttk.Label(self.advanced_tab, text="Number of Registers (-c):").grid(
            column=0, row=3, sticky="W", padx=5, pady=5
        )
        self.entry_num_values = ttk.Entry(self.advanced_tab, width=17)
        self.entry_num_values.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_num_values.insert(0, "1")

        # Register Data Type
        ttk.Label(self.advanced_tab, text="Register Data Type (-t):").grid(
            column=0, row=4, sticky="W", padx=5, pady=5
        )
        self.entry_register_data_type = ttk.Entry(self.advanced_tab, width=17)
        self.entry_register_data_type.grid(column=1, row=4, padx=5, pady=5, sticky="W")
        self.entry_register_data_type.insert(0, "3")

        # Modbus TCP/IP
        ttk.Label(self.advanced_tab, text="Modbus TCP/IP (-m tcp):").grid(
            column=0, row=5, sticky="W", padx=5, pady=5
        )
        self.entry_modbus_tcp = ttk.Entry(self.advanced_tab, width=17)
        self.entry_modbus_tcp.grid(column=1, row=5, padx=5, pady=5, sticky="W")

        # Adjust column weights in Advanced Tab for better layout
        self.advanced_tab.columnconfigure(0, weight=1)
        self.advanced_tab.columnconfigure(1, weight=3)

        # ---------------- Polling Buttons and Status Indicator ----------------
        # Polling Buttons and Status Indicator are placed in the same frame using grid
        btn_frame = ttk.Frame(self.frame_settings)
        btn_frame.pack(pady=10, fill='x', expand=True)

        # Configure grid in btn_frame
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        btn_frame.columnconfigure(3, weight=1)
        btn_frame.columnconfigure(4, weight=1)  # For status indicator
        
        # Store custom command arguments
        self.custom_arguments = None

        # Start Polling Button
        self.btn_start = ttk.Button(
            btn_frame, text="Start Polling", command=self.start_polling
        )
        self.btn_start.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Stop Polling Button
        self.btn_stop = ttk.Button(
            btn_frame, text="Stop Polling", command=self.stop_polling, state="disabled"
        )
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        # Advanced Button
        self.btn_advanced = ttk.Button(
            btn_frame, text="Advanced", command=self.select_advanced_tab
        )
        self.btn_advanced.grid(row=0, column=2, padx=5, pady=5, sticky="e")

        # ---------------- Status Indicator ----------------
        # Create a Canvas for the circular indicator
        self.status_canvas_size = 30  # Increased size from 20 to 30
        self.status_canvas = tk.Canvas(btn_frame, width=self.status_canvas_size, height=self.status_canvas_size, highlightthickness=0)
        self.status_canvas.grid(row=0, column=4, padx=(5, 10), pady=5, sticky="n")  # Centered in the column

        # Draw a circle on the canvas
        self.circle = self.status_canvas.create_oval(
            2, 2, self.status_canvas_size-2, self.status_canvas_size-2, fill="grey"
        )

        # ---------------- Log Frame (Column 2) ----------------
        self.frame_log = ttk.LabelFrame(main_frame, text="Log")
        self.frame_log.grid(row=0, column=2, sticky="NSEW", padx=5, pady=5)
        self.frame_log.rowconfigure(1, weight=1)  # Changed from 0 to 1 to make room for command line
        self.frame_log.columnconfigure(0, weight=1)
        
        # Command Line Frame (Row 0)
        cmd_frame = ttk.Frame(self.frame_log)
        cmd_frame.grid(row=0, column=0, sticky="EW", padx=5, pady=(5, 0))
        cmd_frame.columnconfigure(0, weight=1)
        
        # Command Line Label and Entry
        ttk.Label(cmd_frame, text="Command:").pack(side=tk.LEFT, padx=(0, 5))
        self.cmd_var = tk.StringVar()
        self.entry_cmd = ttk.Entry(cmd_frame, textvariable=self.cmd_var, font=("Consolas", 9))
        self.entry_cmd.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Apply Command Button
        btn_apply_cmd = ttk.Button(cmd_frame, text="Apply", command=self.apply_custom_command, width=8)
        btn_apply_cmd.pack(side=tk.RIGHT)
        
        # Log Text (Row 1)
        self.txt_log = scrolledtext.ScrolledText(
            self.frame_log,
            width=80,
            height=24,  # Reduced height to make room for command line
            state='disabled',
            bg="black",
            fg="white",
            font=("Consolas", 10),
        )
        self.txt_log.grid(column=0, row=1, sticky="NSEW", padx=5, pady=5)

        # Log text tags
        self.txt_log.tag_configure('error', foreground='red')
        self.txt_log.tag_configure('info', foreground='green')

        # Refresh COM ports
        self.refresh_comports()

        # Footer
        footer_frame = ttk.Frame(self.root)
        footer_frame.grid(row=1, column=0, sticky="EW")
        self.root.rowconfigure(1, weight=0)
        self.root.columnconfigure(0, weight=1)

        # Modpoll guide link
        footer_label_left = ttk.Label(
            footer_frame, text="Modpoll guide", foreground="blue", cursor="hand2"
        )
        footer_label_left.pack(side='left', padx=5, pady=5)
        footer_label_left.bind(
            "<Button-1>",
            lambda e: webbrowser.open(
                "https://iwmac.zendesk.com/hc/en-gb/articles/13020280416796-Installasjon-Modpoll-Guide"
            ),
        )

        # Copyright label
        footer_label_right = ttk.Label(footer_frame, text="©TKH")
        footer_label_right.pack(side='right', padx=5, pady=5)

    def select_advanced_tab(self):
        """Switch to the Advanced tab in the settings notebook."""
        self.settings_notebook.select(self.advanced_tab)



    def refresh_comports(self):
        """
        Refresh the list of available COM ports by combining:
        1. COM ports detected by pySerial.
        2. COM ports detected via registry scan.
        """
        ports = serial.tools.list_ports.comports()
        com_ports_serial = [port.device for port in ports]

        com_ports_registry = self.get_com_ports_from_registry()

        # Combine both lists and remove duplicates
        combined_com_ports = sorted(list(set(com_ports_serial + com_ports_registry)))

        # Update the combobox values
        self.cmb_comport['values'] = combined_com_ports

        # If there are available COM ports, select the first one
        if combined_com_ports:
            self.cmb_comport.current(0)
        else:
            self.cmb_comport.set('')

    def get_com_ports_from_registry(self):
        """
        Scan the Windows Registry to find all COM ports.
        This includes COM ports managed by external tools like NPort Administrator.
        """
        com_ports = []
        try:
            reg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
            key = winreg.OpenKey(reg, r"HARDWARE\DEVICEMAP\SERIALCOMM")
            for i in range(0, winreg.QueryInfoKey(key)[1]):
                val = winreg.EnumValue(key, i)
                com_ports.append(val[1])
        except Exception as e:
            self.log_queue.put(('error', f"Error accessing registry for COM ports: {e}"))
        return com_ports

    def format_com_port(self, com_port):
        try:
            port_number = int(com_port.replace("COM", "").replace("com", "").strip())
            if port_number >= 10:
                return f"\\\\.\\{com_port.upper()}"
            else:
                return com_port.upper()
        except ValueError:
            return com_port.upper()

    def start_polling(self):
        if self.is_polling:
            self.log_queue.put(('info', "Polling is already running."))
            return

        # Check if modpoll.exe exists before starting
        if not os.path.exists(self.modpoll_path):
            self.log_queue.put(('error', f"modpoll.exe not found at {self.modpoll_path}. Please wait for download to complete or check the path."))
            messagebox.showwarning("Modpoll Not Found", f"modpoll.exe not found at {self.modpoll_path}. Please wait for download to complete or check the path.")
            return

        com_port = self.cmb_comport.get().strip()
        modbus_tcp = self.entry_modbus_tcp.get().strip()
        use_tcp = bool(modbus_tcp)

        if not com_port and not use_tcp:
            messagebox.showwarning(
                "Missing COM Port or TCP", "Please select a COM port or specify Modbus TCP/IP."
            )
            return

        # Collect parameters
        baudrate = self.cmb_baudrate.get().strip()
        if not baudrate.isdigit():
            messagebox.showwarning("Invalid Baudrate", "Please enter a valid numeric baudrate.")
            return

        parity = self.cmb_parity.get()
        databits = self.cmb_databits.get()
        stopbits = self.cmb_stopbits.get()
        adresse = self.entry_slave_address.get().strip()
        start_reference = self.entry_start_reference.get().strip()
        num_registers = self.entry_num_values.get().strip()
        register_data_type = self.entry_register_data_type.get().strip()

        # Validate numeric entries
        if not all(
            [
                adresse.isdigit(),
                start_reference.isdigit(),
                num_registers.isdigit(),
                register_data_type.isdigit(),
            ]
        ):
            messagebox.showwarning("Invalid Input", "Please enter valid numeric values.")
            return

        # Use custom arguments if available, otherwise build default arguments
        if self.custom_arguments:
            arguments = self.custom_arguments
            self.log_queue.put(('info', f"Using custom command arguments: {' '.join(arguments)}"))
        else:
            # Build command arguments using the same method as equipment selection
            self.build_and_display_command()
            # Get the command from the display
            cmd_text = self.cmd_var.get().strip()
            if cmd_text:
                arguments = cmd_text.split()
            else:
                # Fallback to building arguments directly
                modbus_tcp_cmd = ["-m", "tcp", modbus_tcp] if use_tcp else [self.format_com_port(com_port)]
                arguments = modbus_tcp_cmd + [
                    f"-b{baudrate}",
                    f"-p{parity}",
                    f"-d{databits}",
                    f"-s{stopbits}",
                    f"-a{adresse}",
                    f"-r{start_reference}",
                    f"-c{num_registers}",
                    f"-t{register_data_type}",
                ]

        # Reset attempt counter for a new polling session and store first reference
        self.poll_attempt_counter = 0
        self.current_start_reference = start_reference

        # Start polling thread
        threading.Thread(target=self.run_modpoll, args=(arguments, com_port, baudrate, parity, databits, stopbits, adresse, start_reference, num_registers, register_data_type), daemon=True).start()

    def run_modpoll(self, arguments, com_port, baudrate, parity, databits, stopbits, adresse, start_reference, num_registers, register_data_type):
        self.is_polling = True
        self.update_buttons()
        self.log_queue.put(('info', "Polling started..."))

        try:
            # Use the hardcoded modpoll path
            modpoll_path = self.modpoll_path
            if not os.path.exists(modpoll_path):
                self.log_queue.put(('error', f"Error: modpoll.exe not found at {modpoll_path}"))
                self.stop_polling()
                return

            cmd = [modpoll_path] + arguments

            # Log the command (mask the full path to show just 'modpoll')
            masked_cmd = ['modpoll'] + arguments
            self.log_queue.put(('info', f"Running command: {' '.join(masked_cmd)}"))
            self.log_queue.put(('info', f"Parameters: COM={com_port}, Baud={baudrate}, Parity={parity}, DataBits={databits}, StopBits={stopbits}, Addr={adresse}, Ref={start_reference}, Count={num_registers}, Type={register_data_type}"))

            # Prevent a new window from opening
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

            self.modpoll_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )

            # Start threads to read stdout and stderr
            stdout_thread = threading.Thread(
                target=self.read_stream, args=(self.modpoll_process.stdout,), daemon=True
            )
            stderr_thread = threading.Thread(
                target=self.read_stream, args=(self.modpoll_process.stderr,), daemon=True
            )
            stdout_thread.start()
            stderr_thread.start()

            # Wait for the process to finish
            self.modpoll_process.wait()

            # Wait for both threads to finish
            stdout_thread.join()
            stderr_thread.join()

        except Exception as e:
            self.log_queue.put(('error', f"Error running Modpoll: {str(e)}"))
        finally:
            self.is_polling = False
            self.update_buttons()
            self.log_queue.put(('info', "Polling finished."))

    def _increment_attempt(self):
        self.poll_attempt_counter += 1
        self.log_queue.put(('info', f"Attempt {self.poll_attempt_counter}"))

    def read_stream(self, stream):
        for line in iter(stream.readline, ''):
            line = line.strip()
            if line:
                lower_line = line.lower()

                # Skip the specific unwanted lines
                if "polling slave (ctrl-c to stop) ..." in lower_line:
                    continue  # Do not log this line
                
                # Skip modpoll header and copyright information
                if any(header_text in line for header_text in [
                    "modpoll - FieldTalk(tm) Modbus(R) Polling Utility",
                    "Copyright (c) 2002-2004 FOCUS Software Engineering Pty Ltd",
                    "Getopt Library Copyright (C) 1987-1997	Free Software Foundation, Inc."
                ]):
                    continue  # Do not log these lines

                # Handle "Reply time-out!"
                if "reply time-out!" in lower_line:
                    self._increment_attempt()
                    combined_message = "Reply time-out!\nEquipment does not respond."
                    self.log_queue.put(('error', combined_message))
                    self.trigger_status_indicator('red')

                # Handle "serial port already open"
                elif "serial port already open" in lower_line:
                    combined_message = "Serial port already open!\nRemember to stop plant server in IWMAC escape!"
                    self.log_queue.put(('error', combined_message))
                    self.trigger_status_indicator('red')

                # Handle "port or socket open error!"
                elif "port or socket open error!" in lower_line:
                    combined_message = "Port or socket open error!\nRemember to stop plant server in IWMAC escape!"
                    self.log_queue.put(('error', combined_message))
                    self.trigger_status_indicator('red')

                # Handle "checksum error"
                elif "checksum error" in lower_line:
                    base_message = "Checksum Error Detected!"
                    count = self.message_counts.get(base_message, 0) + 1
                    self.message_counts[base_message] = count
                    numbered_message = f"{base_message} [{count}]"
                    self.log_queue.put(('error', numbered_message))
                    self.trigger_status_indicator('yellow')

                # Handle "illegal function exception response!"
                elif "illegal function exception response!" in lower_line:
                    self._increment_attempt()
                    base_message = "Illegal Function Exception Response!"
                    self.log_queue.put(('info', base_message))
                    self.trigger_status_indicator('green')

                # Handle "illegal data address exception response!"
                elif "illegal data address exception response!" in lower_line:
                    self._increment_attempt()
                    base_message = "Illegal Data Address Exception Response!"
                    self.log_queue.put(('info', base_message))
                    self.trigger_status_indicator('green')

                # Handle "illegal data value exception response!"
                elif "illegal data value exception response!" in lower_line:
                    self._increment_attempt()
                    base_message = "Illegal Data Value Exception Response!"
                    self.log_queue.put(('info', base_message))
                    self.trigger_status_indicator('green')

                # Handle lines like "[100]: 0"
                elif re.match(r'\[\d+\]', line):
                    # Detect start of a new polling cycle using the first reference
                    if self.current_start_reference and line.startswith(f"[{self.current_start_reference}]"):
                        self._increment_attempt()
                    base_message = line
                    count = self.message_counts.get(base_message, 0) + 1
                    self.message_counts[base_message] = count
                    numbered_message = f"{base_message} [{count}]"
                    self.log_queue.put(('info', numbered_message))
                    self.trigger_status_indicator('green')

                else:
                    self.log_queue.put(('normal', line))
        stream.close()

    def stop_polling(self):
        if self.modpoll_process and self.modpoll_process.poll() is None:
            try:
                self.modpoll_process.terminate()
                self.modpoll_process.wait(timeout=5)
                self.log_queue.put(('info', "Polling stopped."))
            except subprocess.TimeoutExpired:
                self.modpoll_process.kill()
                self.log_queue.put(('info', "Polling forcefully stopped."))
            except Exception as e:
                self.log_queue.put(('error', f"Error stopping polling: {str(e)}"))
            finally:
                self.modpoll_process = None
                self.is_polling = False
                self.update_buttons()
        else:
            self.log_queue.put(('info', "No polling process to stop."))

    def update_buttons(self):
        if self.is_polling:
            self.btn_start.config(state="disabled")
            self.btn_stop.config(state="normal")
            self.btn_advanced.config(state="disabled")  # Disable Advanced button during polling
        else:
            self.btn_start.config(state="normal")
            self.btn_stop.config(state="disabled")
            self.btn_advanced.config(state="normal")  # Enable Advanced button when not polling

    def append_log(self, message, tag=None):
        self.txt_log.config(state='normal')
        if tag == 'error':
            self.txt_log.insert(tk.END, f"{message}\n", 'error')
        elif tag == 'info':
            self.txt_log.insert(tk.END, f"{message}\n", 'info')
        else:
            self.txt_log.insert(tk.END, f"{message}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def update_log(self):
        while not self.log_queue.empty():
            item = self.log_queue.get()
            if isinstance(item, tuple):
                tag, message = item
                self.append_log(message, tag)
            else:
                self.append_log(item)
        self.root.after(100, self.update_log)

    def on_closing(self):
        if self.is_polling:
            if messagebox.askokcancel("Exit", "Polling is running. Do you want to exit?"):
                self.stop_polling()
                self.root.destroy()
        else:
            self.root.destroy()

    def select_equipment(self, event=None):
        selected_indices = self.listbox_equipment.curselection()
        if not selected_indices:
            messagebox.showwarning(
                "Selection Error", "Please select an equipment before pressing 'Select'."
            )
            return

        selected_equipment = self.listbox_equipment.get(selected_indices[0])
        
        # Check if "No results found" is selected
        if selected_equipment == "No results found":
            return
            
        # Clear custom command when selecting equipment preset
        if self.custom_arguments:
            self.custom_arguments = None
            self.log_queue.put(('info', f"Custom command cleared - using {selected_equipment} preset"))
            
        settings = self.equipment_settings.get(selected_equipment, {})
        baudrate = settings.get("baudrate", "9600")
        parity = settings.get("parity", "none")
        stop_bits = settings.get("stop_bits", "1")
        data_bits = settings.get("data_bits", "8")

        # Update Baudrate
        if baudrate in self.cmb_baudrate['values']:
            self.cmb_baudrate.set(baudrate)
        else:
            self.cmb_baudrate.set(baudrate)

        # Update Parity
        if parity in self.cmb_parity['values']:
            self.cmb_parity.set(parity)
        else:
            self.cmb_parity.set("none")

        # Update Stop Bits
        if stop_bits in self.cmb_stopbits['values']:
            self.cmb_stopbits.set(stop_bits)
        else:
            self.cmb_stopbits.set("1")  # Default to "1" if invalid

        # Update Data Bits
        if data_bits in self.cmb_databits['values']:
            self.cmb_databits.set(data_bits)
        else:
            self.cmb_databits.set("8")  # Default to "8" if invalid

        # Build and display the full command
        self.build_and_display_command()

        # Switch to Basic tab after selection
        self.settings_notebook.select(self.basic_tab)

    def filter_equipment(self, *args):
        """Filter equipment list based on search term"""
        search_term = self.search_var.get().lower().strip()
        
        # Clear current listbox
        self.listbox_equipment.delete(0, tk.END)
        
        # Filter and populate equipment list
        if search_term:
            filtered_equipment = [eq for eq in self.all_equipment if search_term in eq.lower()]
        else:
            filtered_equipment = self.all_equipment
        
        # Populate filtered list
        for equipment in filtered_equipment:
            self.listbox_equipment.insert(tk.END, equipment)
        
        # If only one result, select it automatically
        if len(filtered_equipment) == 1:
            self.listbox_equipment.selection_set(0)
        elif len(filtered_equipment) == 0 and search_term:
            # Show "No results found" message
            self.listbox_equipment.insert(tk.END, "No results found")
            self.listbox_equipment.itemconfig(tk.END, fg='gray')

    def clear_search(self):
        """Clear search field and show all equipment"""
        self.search_var.set("")
        self.filter_equipment()

    def select_first_equipment(self, event=None):
        """Select the first equipment in the filtered list"""
        if self.listbox_equipment.size() > 0:
            self.listbox_equipment.selection_clear(0, tk.END)
            self.listbox_equipment.selection_set(0)
            self.select_equipment()

    def focus_search(self, event=None):
        """Focus on search field"""
        self.entry_search.focus_set()
        self.entry_search.select_range(0, tk.END)

    def build_and_display_command(self):
        """Build and display the full command based on current settings"""
        # Get current settings
        com_port = self.cmb_comport.get().strip()
        modbus_tcp = self.entry_modbus_tcp.get().strip()
        use_tcp = bool(modbus_tcp)
        
        baudrate = self.cmb_baudrate.get().strip()
        parity = self.cmb_parity.get()
        databits = self.cmb_databits.get()
        stopbits = self.cmb_stopbits.get()
        adresse = self.entry_slave_address.get().strip()
        start_reference = self.entry_start_reference.get().strip()
        num_registers = self.entry_num_values.get().strip()
        register_data_type = self.entry_register_data_type.get().strip()
        
        # Build command arguments
        if use_tcp:
            modbus_tcp_cmd = ["-m", "tcp", modbus_tcp]
            arguments = modbus_tcp_cmd + [
                f"-b{baudrate}",
                f"-p{parity}",
                f"-d{databits}",
                f"-s{stopbits}",
                f"-a{adresse}",
                f"-r{start_reference}",
                f"-c{num_registers}",
                f"-t{register_data_type}",
            ]
        else:
            # Use COM port if available
            if com_port:
                # Ensure COM port is properly formatted (e.g., "COM1" not just "1")
                if not com_port.upper().startswith('COM'):
                    formatted_com_port = f"COM{com_port}"
                else:
                    formatted_com_port = com_port.upper()
                
                # Apply the same formatting logic as format_com_port method
                try:
                    port_number = int(formatted_com_port.replace("COM", "").replace("com", "").strip())
                    if port_number >= 10:
                        formatted_com_port = f"\\\\.\\{formatted_com_port}"
                except ValueError:
                    pass
                
                arguments = [formatted_com_port] + [
                    f"-b{baudrate}",
                    f"-p{parity}",
                    f"-d{databits}",
                    f"-s{stopbits}",
                    f"-a{adresse}",
                    f"-r{start_reference}",
                    f"-c{num_registers}",
                    f"-t{register_data_type}",
                ]
            else:
                # Show placeholder if no COM port selected
                arguments = ["[COM_PORT]"] + [
                    f"-b{baudrate}",
                    f"-p{parity}",
                    f"-d{databits}",
                    f"-s{stopbits}",
                    f"-a{adresse}",
                    f"-r{start_reference}",
                    f"-c{num_registers}",
                    f"-t{register_data_type}",
                ]
        
        # Update command line display
        self.cmd_var.set(' '.join(arguments))

    def apply_custom_command(self):
        """Apply custom command from the command line entry"""
        custom_cmd = self.cmd_var.get().strip()
        if custom_cmd:
            # Parse the custom command into arguments
            try:
                # Split the command into arguments, handling quotes properly
                import shlex
                self.custom_arguments = shlex.split(custom_cmd)
                self.log_queue.put(('info', f"Custom command applied: {custom_cmd}"))
            except Exception as e:
                self.log_queue.put(('error', f"Invalid command format: {e}"))
                self.custom_arguments = None
        else:
            self.custom_arguments = None
            self.log_queue.put(('info', "Custom command cleared, using default settings"))

    def toggle_equipment_pane(self):
        if not self.equipment_pane_visible:
            self.show_equipment_pane()
        else:
            self.hide_equipment_pane()

    def show_equipment_pane(self):
        # Since the equipment pane is already visible in the grid layout,
        # toggling its visibility can be done by showing or hiding it.
        self.frame_equipment.grid()
        self.equipment_pane_visible = True

    def hide_equipment_pane(self):
        self.frame_equipment.grid_remove()
        self.equipment_pane_visible = False

    # ---------------- Status Indicator Methods ----------------

    def trigger_status_indicator(self, response_type):
        """
        Trigger the status indicator based on the response type.
        response_type: 'green', 'yellow', 'red'
        """
        color_map = {
            'green': (self.GREEN_COLOR, self.GREEN_DIM),
            'yellow': (self.YELLOW_COLOR, self.YELLOW_DIM),
            'red': (self.RED_COLOR, self.RED_DIM)
        }

        if response_type not in color_map:
            return

        self.current_color, self.dim_color = color_map[response_type]

        # If already blinking, cancel the previous job
        if self.blink_job:
            self.root.after_cancel(self.blink_job)
            self.blink_job = None
            self.blinking = False
            self.blink_count = 0

        # Start blinking
        self.blink_state = False
        self.blinking = True
        self.blink_count = 0
        self.blink_indicator()

    def blink_indicator(self):
        if not self.blinking:
            return

        if self.blink_state:
            # Set to dim color
            self.status_canvas.itemconfig(self.circle, fill=self.dim_color)
        else:
            # Set to main color
            self.status_canvas.itemconfig(self.circle, fill=self.current_color)

        self.blink_state = not self.blink_state
        self.blink_count += 1

        # Schedule the next blink
        self.blink_job = self.root.after(500, self.blink_indicator)  # Blink every 500ms

        # Stop blinking after 10 toggles (5 seconds)
        if self.blink_count >= 10:
            self.blinking = False
            self.blink_count = 0
            self.status_canvas.itemconfig(self.circle, fill="grey")
            if self.blink_job:
                self.root.after_cancel(self.blink_job)
                self.blink_job = None

    # ---------------- End of Status Indicator Methods ----------------

if __name__ == "__main__":
    root = tk.Tk()
    tool = ModpollingTool(root)
    root.protocol("WM_DELETE_WINDOW", tool.on_closing)
    root.mainloop()
