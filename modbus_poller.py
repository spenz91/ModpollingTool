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
import urllib.parse  # For URL encoding
import json  # For JSON parsing
import shlex  # For parsing command line arguments
from tkinter import font as tkfont


class ModpollingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ModPolling Tool")
        
        # Modern window styling
        self.root.configure(bg='#F0F0F0')  # Modern light gray background
        self.root.option_add('*TFrame*background', '#F0F0F0')
        self.root.option_add('*TLabel*background', '#F0F0F0')
        self.root.option_add('*TButton*background', '#F0F0F0')
        
        # Set window icon if available (optional)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass  # No icon file found, continue without it
        
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
            "Carlo Gavazzi EM100": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM21": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM210": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM23": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM24": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM24TCP": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
            "Carlo Gavazzi EM26": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM270": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM330": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM4": {"baudrate": "9600", "stop_bits": "1", "data_bits": "8", "parity": "none"},
            "Carlo Gavazzi EM540": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
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
            "SLV AHT": {"baudrate": "19200", "stop_bits": "1", "data_bits": "8", "parity": "even"},
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

        # Define modern colors
        self.GREEN_COLOR = "#10B981"  # Modern green
        self.GREEN_DIM = "#34D399"    # Lighter green
        self.YELLOW_COLOR = "#F59E0B"  # Modern amber
        self.YELLOW_DIM = "#FBBF24"    # Lighter amber
        self.RED_COLOR = "#EF4444"     # Modern red
        self.RED_DIM = "#F87171"       # Lighter red

        self.create_widgets()
        self.update_log()

    def create_modern_button_styles(self):
        """
        Create modern button styles with hover effects and rounded corners.
        """
        # Configure modern style for ttk widgets
        style = ttk.Style()
        
        # Modern color scheme
        modern_colors = {
            'primary': '#10B981',      # Green
            'secondary': '#6B7280',    # Gray
            'danger': '#EF4444',       # Red
            'success': '#10B981',      # Green
            'warning': '#F59E0B',      # Amber
            'info': '#3B82F6',         # Blue
            'light': '#F9FAFB',        # Light gray
            'dark': '#1F2937',         # Dark gray
            'white': '#FFFFFF',        # White
            'black': '#000000'         # Black
        }
        
        # Store colors for later use
        self.modern_colors = modern_colors

    def setup_command_preview_bindings(self):
        """
        Setup event bindings to automatically update command preview when settings change.
        """
        # Basic tab bindings
        self.cmb_comport.bind('<<ComboboxSelected>>', self.update_command_preview)
        self.cmb_comport.bind('<KeyRelease>', self.update_command_preview)
        self.cmb_baudrate.bind('<<ComboboxSelected>>', self.update_command_preview)
        self.cmb_baudrate.bind('<KeyRelease>', self.update_command_preview)
        self.cmb_parity.bind('<<ComboboxSelected>>', self.update_command_preview)
        self.entry_slave_address.bind('<KeyRelease>', self.update_command_preview)
        
        # Advanced tab bindings
        self.cmb_databits.bind('<<ComboboxSelected>>', self.update_command_preview)
        self.cmb_stopbits.bind('<<ComboboxSelected>>', self.update_command_preview)
        self.entry_start_reference.bind('<KeyRelease>', self.update_command_preview)
        self.entry_num_values.bind('<KeyRelease>', self.update_command_preview)
        self.entry_register_data_type.bind('<KeyRelease>', self.update_command_preview)
        self.entry_modbus_tcp.bind('<KeyRelease>', self.update_command_preview)

    def update_command_preview(self, event=None):
        """
        Update the command preview when any setting changes.
        """
        # Only update if not using custom arguments
        if not self.custom_arguments:
            self.build_and_display_command()

    def setup_button_hover_effects(self):
        """
        Setup hover effects for modern buttons.
        """
        # Start button hover effects
        self.btn_start.bind("<Enter>", lambda e: self.on_button_hover_enter(self.btn_start, "#059669"))
        self.btn_start.bind("<Leave>", lambda e: self.on_button_hover_leave(self.btn_start, "#10B981"))
        
        # Stop button hover effects - only hover when polling is active
        self.btn_stop.bind("<Enter>", lambda e: self.on_stop_hover_enter())
        self.btn_stop.bind("<Leave>", lambda e: self.on_stop_hover_leave())
        


    def on_button_hover_enter(self, button, hover_color):
        """
        Handle button hover enter event.
        """
        if button.cget("state") != "disabled":
            button.config(bg=hover_color)

    def on_button_hover_leave(self, button, normal_color):
        """
        Handle button hover leave event.
        """
        if button.cget("state") != "disabled":
            button.config(bg=normal_color)



    def on_stop_hover_enter(self):
        """
        Handle stop button hover enter event - only hover when polling is active.
        """
        if self.is_polling:
            self.btn_stop.config(bg="#DC2626")

    def on_stop_hover_leave(self):
        """
        Handle stop button hover leave event - only hover when polling is active.
        """
        if self.is_polling:
            self.btn_stop.config(bg="#EF4444")

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
        main_frame.grid(row=0, column=0, sticky="NSEW", padx=10, pady=10)
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

        # Bind selection change event to update command immediately
        self.listbox_equipment.bind('<<ListboxSelect>>', self.on_equipment_selection_change)
        
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
        
        # Units Tab
        self.units_tab = ttk.Frame(self.settings_notebook)
        self.settings_notebook.add(self.units_tab, text="Units")

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

        # ---------------- Units Tab Widgets ----------------
        # Configure grid for Units Tab
        self.units_tab.columnconfigure(0, weight=1)
        self.units_tab.columnconfigure(1, weight=1)  # Add weight to center the table
        self.units_tab.columnconfigure(2, weight=1)  # Add weight to center the table
        self.units_tab.rowconfigure(1, weight=1)  # Make the treeview expandable
        
        # Units Header
        units_header_frame = ttk.Frame(self.units_tab)
        units_header_frame.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        units_header_frame.columnconfigure(1, weight=1)
        
        ttk.Label(units_header_frame, text="Plant Units:", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        
        # Refresh Units Button
        self.btn_refresh_units = ttk.Button(
            units_header_frame, text="🔄 Refresh Units", command=self.refresh_units
        )
        self.btn_refresh_units.grid(row=0, column=1, sticky="e", padx=5, pady=5)
        
        # Units Treeview
        self.units_tree = ttk.Treeview(self.units_tab, columns=("unit_id", "unit_name", "status", "owner", "tree", "ip"), show="headings", height=15)
        self.units_tree.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        
        # Configure columns with sorting
        self.units_tree.heading("unit_id", text="Unit ID", command=lambda: self.sort_units_tree("unit_id", False))
        self.units_tree.heading("unit_name", text="Unit Name", command=lambda: self.sort_units_tree("unit_name", False))
        self.units_tree.heading("status", text="Status", command=lambda: self.sort_units_tree("status", False))
        self.units_tree.heading("owner", text="Owner", command=lambda: self.sort_units_tree("owner", False))
        self.units_tree.heading("tree", text="COM Port", command=lambda: self.sort_units_tree("tree", False))
        self.units_tree.heading("ip", text="IP Address", command=lambda: self.sort_units_tree("ip", False))
        
        # Set column widths and center alignment
        self.units_tree.column("unit_id", width=80, minwidth=80, anchor="center")
        self.units_tree.column("unit_name", width=150, minwidth=150, anchor="center")
        self.units_tree.column("status", width=100, minwidth=100, anchor="center")
        self.units_tree.column("owner", width=100, minwidth=100, anchor="center")
        self.units_tree.column("tree", width=80, minwidth=80, anchor="center")
        self.units_tree.column("ip", width=120, minwidth=120, anchor="center")
        
        # Add scrollbar
        units_scrollbar = ttk.Scrollbar(self.units_tab, orient="vertical", command=self.units_tree.yview)
        units_scrollbar.grid(row=1, column=2, sticky="ns")
        self.units_tree.configure(yscrollcommand=units_scrollbar.set)
        
        # Status label for units
        self.units_status_label = ttk.Label(self.units_tab, text="Click 'Refresh Units' to load unit data", foreground="gray")
        self.units_status_label.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        # Initialize sorting state
        self.units_sort_column = None
        self.units_sort_reverse = False

        # ---------------- Polling Buttons and Status Indicator ----------------
        # Polling Buttons and Status Indicator are placed in the same frame using grid
        btn_frame = ttk.Frame(self.frame_settings)
        btn_frame.pack(pady=15, fill='x', expand=True)

        # Configure grid in btn_frame
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)  # For status indicator
        
        # Store custom command arguments
        self.custom_arguments = None

        # Create modern button styles
        self.create_modern_button_styles()

        # Start Polling Button - Modern Green Style
        self.btn_start = tk.Button(
            btn_frame, 
            text="▶ Start Polling", 
            command=self.start_polling,
            font=("Segoe UI", 10, "bold"),
            bg="#10B981",  # Modern green
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.btn_start.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Stop Polling Button - Modern Red Style
        self.btn_stop = tk.Button(
            btn_frame, 
            text="■ Stop Polling", 
            command=self.stop_polling, 
            state="normal",
            font=("Segoe UI", 10, "bold"),
            bg="#FCA5A5",  # Light red when disabled
            fg="white",
            activebackground="#DC2626",
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5, sticky="e")

        # Store the frame background color
        self.frame_bg_color = self.root.cget("bg")  # Get the window background color

        # Add hover effects for modern buttons
        self.setup_button_hover_effects()

        # ---------------- Status Indicator ----------------
        # Create a Canvas for the modern circular indicator
        self.status_canvas_size = 48  # Bigger for modern look
        self.status_canvas = tk.Canvas(
            btn_frame, 
            width=self.status_canvas_size, 
            height=self.status_canvas_size, 
            highlightthickness=0,
            bg=self.frame_bg_color  # Use stored background color
        )
        self.status_canvas.grid(row=0, column=2, padx=(5, 10), pady=5, sticky="n")  # Centered in the column

        # Draw a modern circle with gradient-like effect
        self.circle = self.status_canvas.create_oval(
            2, 2, self.status_canvas_size-2, self.status_canvas_size-2, 
            fill="#D1D5DB",  # Light gray for modern look
            outline="#9CA3AF",  # Medium gray outline
            width=2
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
            wrap=tk.NONE,  # Disable word wrap for better performance
            undo=False,    # Disable undo for better performance
            maxundo=0,     # Disable undo history
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

        # Refresh plant data button
        btn_refresh_plant = ttk.Button(
            footer_frame, text="Refresh Plant Data", command=self.refresh_plant_data
        )
        btn_refresh_plant.pack(side='left', padx=5, pady=5)



        # Copyright label
        footer_label_right = ttk.Label(footer_frame, text="©TKH")
        footer_label_right.pack(side='right', padx=5, pady=5)

        # Add event bindings for automatic command preview updates (after all widgets are created)
        self.setup_command_preview_bindings()

    def select_advanced_tab(self):
        """Switch to the Advanced tab in the settings notebook."""
        self.settings_notebook.select(self.advanced_tab)



    def refresh_comports(self):
        """
        Refresh the list of available COM ports by combining:
        1. COM ports detected by pySerial.
        2. COM ports detected via registry scan.
        3. Enhanced names from plant data (if available).
        """
        ports = serial.tools.list_ports.comports()
        com_ports_serial = [port.device for port in ports]

        com_ports_registry = self.get_com_ports_from_registry()

        # Combine both lists and remove duplicates
        combined_com_ports = sorted(list(set(com_ports_serial + com_ports_registry)))
        
        # Get enhanced COM port names from plant data
        enhanced_com_ports = self.get_enhanced_com_port_names(combined_com_ports)

        # Update the combobox values with enhanced names
        self.cmb_comport['values'] = enhanced_com_ports

        # Log the available COM ports
        if enhanced_com_ports:
            pass  # Removed detailed logging
        else:
            pass  # Removed verbose logging

        # If there are available COM ports, select the first one
        if enhanced_com_ports:
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
            self.log_queue.put(('error', "Error! Cannot find COM ports. Type in the COM port manually."))
        return com_ports

    def get_enhanced_com_port_names(self, com_ports):
        """
        Get enhanced COM port names from plant data.
        Returns list of COM ports with enhanced names like "COM1 - SLV".
        """
        enhanced_ports = []
        
        # Try to fetch plant data for COM port information
        try:
            plant_data = self.fetch_plant_data()
            if plant_data:
                # Create a mapping of COM ports to their enhanced names
                com_port_mapping = self.parse_plant_data_for_com_ports(plant_data)
                
                # Log the mapping for debugging
                if com_port_mapping:
                    pass  # Removed detailed logging
                
                for com_port in com_ports:
                    if com_port in com_port_mapping:
                        enhanced_ports.append(com_port_mapping[com_port])
                    else:
                        enhanced_ports.append(com_port)
            else:
                # If plant data is not available, use original COM port names
                enhanced_ports = com_ports
                pass  # Removed verbose logging
        except Exception as e:
            self.log_queue.put(('info', f"Could not fetch plant data for COM port names: {e}"))
            enhanced_ports = com_ports
            
        return enhanced_ports

    def refresh_plant_data(self):
        """
        Manually refresh plant data and update COM port names.
        """
        self.log_queue.put(('info', "Refreshing plant data..."))
        self.refresh_comports()



    def fetch_plant_data(self):
        """
        Fetch plant data from the topology API with authentication.
        Returns the JSON response or None if failed.
        """
        try:
            # API endpoint and credentials (obfuscated)
            username = self._decode_str("aXdtYWM=")
            password = self._decode_str("U2pha2sgPEVyIEJhcmUhIEZvciAy")
            host = self._decode_str("MTI3LjAuMC4xOjgx")
            path = self._decode_str("L3NlY3VyZS9zeXNfdG9vbHMvcGxhbnRfZGF0YS5waHA=")
            
            # Build the URL without authentication
            base_url = f"http://{host}{path}"
            params = {
                'cmd': self._decode_str("dG9wb2xvZ3k="),
                'type': self._decode_str("dHlwZQ=="),
                'request': self._decode_str("eyJsaW1pdCI6MTAwLCJvZmZzZXQiOjAsInNvcnQiOlt7ImZpZWxkIjoibWFuX3N0YXJ0IiwiZGlyZWN0aW9uIjoiYXNjIn1dfQ==")
            }
            
            full_url = f"{base_url}?{urllib.parse.urlencode(params)}"
            
            # Create a request with authentication headers
            req = urllib.request.Request(full_url)
            req.add_header('User-Agent', 'ModPollingTool/1.0')
            
            # Add basic authentication header
            import base64
            credentials = base64.b64encode(f"{username}:{password}".encode('utf-8')).decode('utf-8')
            req.add_header('Authorization', f'Basic {credentials}')
            
            # Make the request with a 5-second timeout
            with urllib.request.urlopen(req, timeout=5) as response:
                data = response.read()
                plant_data = json.loads(data.decode('utf-8'))
                self.log_queue.put(('info', "Successfully fetched plant data!"))
                return plant_data
                
        except urllib.error.URLError as e:
            error_msg = str(e)
            if "10061" in error_msg or "refused" in error_msg.lower():
                pass  # Removed verbose logging
            else:
                pass  # Removed verbose logging
            return None
        except urllib.error.HTTPError as e:
            self.log_queue.put(('info', f"HTTP error {e.code} from plant data API: {e.reason}"))
            return None
        except json.JSONDecodeError as e:
            self.log_queue.put(('error', f"Invalid JSON response from plant data API: {e}"))
            return None
        except Exception as e:
            self.log_queue.put(('error', f"Unexpected error fetching plant data: {e}"))
            return None

    def _decode_str(self, encoded_str):
        """Simple base64 decode function to obfuscate strings."""
        import base64
        return base64.b64decode(encoded_str).decode('utf-8')

    def parse_plant_data_for_com_ports(self, plant_data):
        """
        Parse plant data to extract COM port mappings.
        Expected format from API: {"tree": "COM1 - 192.168.10.31", "w2ui": {"children": [{"owner": "SLV"}]}}
        Returns: {"COM1": "COM1 - SLV"}
        """
        com_port_mapping = {}
        ip_address = ""
        
        try:
            # Parse the JSON structure from the API
            if isinstance(plant_data, list):
                for item in plant_data:
                    if isinstance(item, dict) and 'w2ui' in item:
                        w2ui_data = item.get('w2ui', {})
                        if 'children' in w2ui_data:
                            for child in w2ui_data['children']:
                                if isinstance(child, dict):
                                    tree_info = child.get('tree', '')
                                    
                                    # Extract IP address from tree info (e.g., "COM1 - 192.168.10.31")
                                    ip_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', tree_info)
                                    if ip_match and not ip_address:
                                        ip_address = ip_match.group(1)
                                    
                                    # Look for COM port in tree info (e.g., "COM1 - 192.168.10.31")
                                    com_match = re.search(r'COM(\d+)', tree_info)
                                    if com_match:
                                        com_port = f"COM{com_match.group(1)}"
                                        
                                        # Look for owner in the children of this COM port
                                        if 'w2ui' in child and 'children' in child['w2ui']:
                                            for unit in child['w2ui']['children']:
                                                if isinstance(unit, dict) and 'owner' in unit:
                                                    owner = unit.get('owner', '')
                                                    
                                                    if owner:
                                                        enhanced_name = f"{com_port} - {owner}"
                                                        com_port_mapping[com_port] = enhanced_name
                                                        break  # Use the first owner found
                                
        except Exception as e:
            self.log_queue.put(('error', f"Error parsing plant data: {e}"))
        
        # Display summary if mappings were found
        if com_port_mapping:
            self.log_queue.put(('info', f"Found {len(com_port_mapping)} COM port mappings!"))
            if ip_address:
                self.log_queue.put(('info', f"IP adresse: '{ip_address}'"))
            for com_port, enhanced_name in sorted(com_port_mapping.items()):
                owner = enhanced_name.split(' - ')[1] if ' - ' in enhanced_name else enhanced_name
                self.log_queue.put(('info', f"☑ {com_port} -> {owner}"))
            
        return com_port_mapping

    def parse_units_data(self, plant_data):
        """
        Parse plant data to extract unit information.
        Returns a list of dictionaries with unit details.
        """
        units = []
        
        try:
            if isinstance(plant_data, list):
                for item in plant_data:
                    if isinstance(item, dict) and 'w2ui' in item:
                        w2ui_data = item.get('w2ui', {})
                        if 'children' in w2ui_data:
                            for child in w2ui_data['children']:
                                if isinstance(child, dict):
                                    tree_info = child.get('tree', '')
                                    
                                    # Extract COM port and IP from tree info (e.g., "COM1 - 192.168.10.31")
                                    com_match = re.search(r'COM(\d+)', tree_info)
                                    ip_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', tree_info)
                                    
                                    com_port = f"COM{com_match.group(1)}" if com_match else ""
                                    ip_address = ip_match.group(1) if ip_match else ""
                                    
                                    # Look for units in the children of this COM port
                                    if 'w2ui' in child and 'children' in child['w2ui']:
                                        for unit in child['w2ui']['children']:
                                            if isinstance(unit, dict):
                                                unit_info = {
                                                    'unit_id': unit.get('unit_id', ''),
                                                    'status': unit.get('status', ''),
                                                    'owner': unit.get('owner', ''),
                                                    'unit_name': unit.get('unit_name', ''),
                                                    'tree': com_port,
                                                    'ip': ip_address
                                                }
                                                units.append(unit_info)
                                
        except Exception as e:
            self.log_queue.put(('error', f"Error parsing units data: {e}"))
            
        return units

    def refresh_units(self):
        """
        Fetch and display units data from the API.
        """
        def refresh_thread():
            try:
                self.units_status_label.config(text="Fetching units data...", foreground="blue")
                self.btn_refresh_units.config(state="disabled")
                
                # Clear existing items
                for item in self.units_tree.get_children():
                    self.units_tree.delete(item)
                
                # Fetch plant data
                plant_data = self.fetch_plant_data()
                
                if plant_data:
                    # Parse units data
                    units = self.parse_units_data(plant_data)
                    
                    if units:
                        # Add units to treeview
                        for unit in units:
                            self.units_tree.insert('', 'end', values=(
                                unit['unit_id'],
                                unit['unit_name'],
                                unit['status'],
                                unit['owner'],
                                unit['tree'],
                                unit['ip']
                            ))
                        
                        self.units_status_label.config(
                            text=f"Loaded {len(units)} units successfully", 
                            foreground="green"
                        )
                        self.log_queue.put(('info', f"Successfully loaded {len(units)} units"))
                    else:
                        self.units_status_label.config(
                            text="No units found in plant data", 
                            foreground="orange"
                        )
                        self.log_queue.put(('info', "No units found in plant data"))
                else:
                    self.units_status_label.config(
                        text="Failed to fetch plant data", 
                        foreground="red"
                    )
                    self.log_queue.put(('error', "Failed to fetch plant data"))
                    
            except Exception as e:
                self.units_status_label.config(
                    text=f"Error: {str(e)}", 
                    foreground="red"
                )
                self.log_queue.put(('error', f"Error refreshing units: {e}"))
            finally:
                self.btn_refresh_units.config(state="normal")
        
        # Run in separate thread to avoid blocking UI
        threading.Thread(target=refresh_thread, daemon=True).start()

    def sort_units_tree(self, col, reverse):
        """
        Sort tree contents when a column header is clicked.
        """
        try:
            # Get all items from the tree
            items = [(self.units_tree.set(item, col), item) for item in self.units_tree.get_children('')]
            
            # Sort the items
            items.sort(reverse=reverse)
            
            # Rearrange items in sorted positions
            for index, (val, item) in enumerate(items):
                self.units_tree.move(item, '', index)
            
            # Update sorting state
            self.units_sort_column = col
            self.units_sort_reverse = reverse
            
            # Update column headers to show sort direction
            for column in ("unit_id", "unit_name", "status", "owner", "tree", "ip"):
                if column == col:
                    # Add arrow to show sort direction
                    arrow = " ▼" if reverse else " ▲"
                    current_text = self.units_tree.heading(column)['text'].replace(" ▲", "").replace(" ▼", "")
                    self.units_tree.heading(column, text=current_text + arrow)
                else:
                    # Remove arrows from other columns
                    current_text = self.units_tree.heading(column)['text'].replace(" ▲", "").replace(" ▼", "")
                    self.units_tree.heading(column, text=current_text)
            
            # Toggle reverse for next click
            self.units_tree.heading(col, command=lambda: self.sort_units_tree(col, not reverse))
            
        except Exception as e:
            self.log_queue.put(('error', f"Error sorting units table: {e}"))

    def extract_com_port_from_enhanced_name(self, enhanced_name):
        """
        Extract the actual COM port from an enhanced name like "COM1 - SLV".
        Returns the original COM port name.
        """
        try:
            # Extract COM port from enhanced name (e.g., "COM1 - SLV" -> "COM1")
            com_match = re.search(r'COM\d+', enhanced_name)
            if com_match:
                return com_match.group(0)
            else:
                return enhanced_name
        except Exception:
            return enhanced_name

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

        com_port_enhanced = self.cmb_comport.get().strip()
        # Extract actual COM port from enhanced name (e.g., "COM1 - SLV" -> "COM1")
        com_port = self.extract_com_port_from_enhanced_name(com_port_enhanced)
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
                if use_tcp:
                    arguments = ["-m", "tcp", modbus_tcp] + [
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
                    arguments = [self.format_com_port(com_port), f"-b{baudrate}", f"-p{parity}", f"-a{adresse}"]

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
                bufsize=0,  # No buffering to force immediate output
                env=dict(os.environ, PYTHONUNBUFFERED="1"),  # Force unbuffered output
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
        # Log every attempt directly to text widget in green
        self.append_log_direct(f"Attempt {self.poll_attempt_counter}", 'info')

    def read_stream(self, stream):
        # Native Windows console behavior - direct output like CMD
        import msvcrt
        
        buffer = ""
        while True:
            # Read character by character like native CMD
            char = stream.read(1)
            if not char:
                break
            
            buffer += char
            
            # Process complete lines immediately like CMD
            if char == '\n':
                line = buffer.strip()
                buffer = ""
                
                if line:
                    lower_line = line.lower()

                    # Skip the polling slave line - don't show it
                    if "polling slave (ctrl-c to stop) ..." in lower_line:
                        continue
                    
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
                        self.append_log_direct(combined_message, 'error')
                        self.trigger_status_indicator('red')

                    # Handle "serial port already open"
                    elif "serial port already open" in lower_line:
                        combined_message = "Serial port already open!\nRemember to stop plant server in IWMAC escape!"
                        self.append_log_direct(combined_message, 'error')
                        self.trigger_status_indicator('red')

                    # Handle "port or socket open error!"
                    elif "port or socket open error!" in lower_line:
                        combined_message = "Port or socket open error!\nRemember to stop plant server in IWMAC escape!"
                        self.append_log_direct(combined_message, 'error')
                        self.trigger_status_indicator('red')

                    # Handle "checksum error"
                    elif "checksum error" in lower_line:
                        base_message = "Checksum Error Detected!"
                        count = self.message_counts.get(base_message, 0) + 1
                        self.message_counts[base_message] = count
                        numbered_message = f"{base_message} [{count}]"
                        self.append_log_direct(numbered_message, 'error')
                        self.trigger_status_indicator('yellow')

                    # Handle "illegal function exception response!"
                    elif "illegal function exception response!" in lower_line:
                        self._increment_attempt()
                        base_message = "Illegal Function Exception Response!"
                        self.append_log_direct(base_message, 'info')
                        self.trigger_status_indicator('green')

                    # Handle "illegal data address exception response!"
                    elif "illegal data address exception response!" in lower_line:
                        self._increment_attempt()
                        base_message = "Illegal Data Address Exception Response!"
                        self.append_log_direct(base_message, 'info')
                        self.trigger_status_indicator('green')

                    # Handle "illegal data value exception response!"
                    elif "illegal data value exception response!" in lower_line:
                        self._increment_attempt()
                        base_message = "Illegal Data Value Exception Response!"
                        self.append_log_direct(base_message, 'info')
                        self.trigger_status_indicator('green')

                    # Handle lines like "[100]: 70" - log exactly as modpoll outputs
                    elif re.match(r'\[\d+\]', line):
                        # Detect start of a new polling cycle using the first reference
                        if self.current_start_reference and line.startswith(f"[{self.current_start_reference}]"):
                            self._increment_attempt()
                        self.append_log_direct(line, 'normal')
                        self.trigger_status_indicator('green')

                    else:
                        self.append_log_direct(line, 'normal')
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
            # Disable start button and change appearance
            self.btn_start.config(state="disabled", bg="#9CA3AF", fg="#6B7280")
            # Enable stop button and restore appearance
            self.btn_stop.config(state="normal", bg="#EF4444", fg="white")
        else:
            # Enable start button and restore appearance
            self.btn_start.config(state="normal", bg="#10B981", fg="white")
            # Show stop button as disabled but keep it enabled for consistent styling
            self.btn_stop.config(state="normal", bg="#FCA5A5", fg="white")

    def append_log(self, message, tag=None):
        self.txt_log.config(state='normal')
        if tag == 'error':
            self.txt_log.insert(tk.END, f"{message}\n", 'error')
        elif tag == 'info':
            self.txt_log.insert(tk.END, f"{message}\n", 'info')
        else:
            self.txt_log.insert(tk.END, f"{message}\n")
        
        # Always scroll to end for real-time feel like command line
        self.txt_log.see(tk.END)
        
        self.txt_log.config(state='disabled')

    def append_log_direct(self, message, tag=None):
        """Direct logging like a real terminal"""
        try:
            self.txt_log.config(state='normal')
            if tag == 'error':
                self.txt_log.insert(tk.END, f"{message}\n", 'error')
            elif tag == 'info':
                self.txt_log.insert(tk.END, f"{message}\n", 'info')
            else:
                self.txt_log.insert(tk.END, f"{message}\n")
            
            # Scroll to end like a real terminal
            self.txt_log.see(tk.END)
            
            self.txt_log.config(state='disabled')
            
            # Force immediate update like native CMD
            self.root.update_idletasks()
        except Exception as e:
            # Fallback to queue if direct logging fails
            self.log_queue.put((tag or 'normal', message))

    def update_log(self):
        # Process all available log entries immediately for real-time feel
        while not self.log_queue.empty():
            item = self.log_queue.get()
            if isinstance(item, tuple):
                tag, message = item
                self.append_log(message, tag)
            else:
                self.append_log(item)
        
        # Update very frequently during polling for real-time behavior
        update_interval = 10 if self.is_polling else 100
        self.root.after(update_interval, self.update_log)

    def on_closing(self):
        if self.is_polling:
            if messagebox.askokcancel("Exit", "Polling is running. Do you want to exit?"):
                self.stop_polling()
                self.root.destroy()
        else:
            self.root.destroy()

    def on_equipment_selection_change(self, event=None):
        """Handle equipment selection change - update command immediately"""
        selected_indices = self.listbox_equipment.curselection()
        if not selected_indices:
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

        # Build and display the full command immediately
        self.build_and_display_command()

        # Switch to Basic tab after selection
        self.settings_notebook.select(self.basic_tab)

    def select_equipment(self, event=None):
        """Legacy method for button clicks and Enter key - now just calls the new method"""
        self.on_equipment_selection_change(event)

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
        com_port_enhanced = self.cmb_comport.get().strip()
        # Extract actual COM port from enhanced name (e.g., "COM1 - SLV" -> "COM1")
        com_port = self.extract_com_port_from_enhanced_name(com_port_enhanced)
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
            # TCP mode - show full command with all parameters
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
                
                # Build basic arguments
                arguments = [formatted_com_port, f"-b{baudrate}", f"-p{parity}", f"-a{adresse}"]
                
                # Add advanced parameters if they differ from defaults
                if databits != "8":
                    arguments.append(f"-d{databits}")
                if stopbits != "1":
                    arguments.append(f"-s{stopbits}")
                if start_reference != "100":
                    arguments.append(f"-r{start_reference}")
                if num_registers != "1":
                    arguments.append(f"-c{num_registers}")
                if register_data_type != "3":
                    arguments.append(f"-t{register_data_type}")
            else:
                # Show placeholder if no COM port selected
                arguments = ["[COM_PORT]", f"-b{baudrate}", f"-p{parity}", f"-a{adresse}"]
                
                # Add advanced parameters if they differ from defaults
                if databits != "8":
                    arguments.append(f"-d{databits}")
                if stopbits != "1":
                    arguments.append(f"-s{stopbits}")
                if start_reference != "100":
                    arguments.append(f"-r{start_reference}")
                if num_registers != "1":
                    arguments.append(f"-c{num_registers}")
                if register_data_type != "3":
                    arguments.append(f"-t{register_data_type}")
        
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
