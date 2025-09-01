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

import shlex  # For parsing command line arguments
from tkinter import font as tkfont
import datetime
import urllib
import urllib.request
import urllib.error


class ModpollingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ModPolling Tool")
        
        # Set smaller default window size
        self.root.geometry("1195x628")  # Exact size requested by user
        self.root.minsize(1000, 600)   # Minimum size to ensure usability
        
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
        # Default MySQL client path
        self.mysql_exe_path = r"C:\iwmac\mysql\bin\mysql.exe"
        
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
        main_frame.columnconfigure(0, weight=1)

        # Create a PanedWindow for resizable sections
        self.paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned_window.grid(row=0, column=0, sticky="NSEW")

        # Equipment Frame (Column 0) - Left pane
        self.frame_equipment = ttk.LabelFrame(self.paned_window, text="Equipment")
        self.paned_window.add(self.frame_equipment, weight=10)  # Fixed size for Equipment (smaller)
        self.frame_equipment.rowconfigure(1, weight=1)  # Changed from 0 to 1 to make room for search
        self.frame_equipment.columnconfigure(0, weight=1)

        # Search Frame (Row 0)
        search_frame = ttk.Frame(self.frame_equipment)
        search_frame.grid(row=0, column=0, sticky="EW", padx=5, pady=(5, 0))
        search_frame.columnconfigure(0, weight=1)
        search_frame.columnconfigure(1, weight=3)
        search_frame.columnconfigure(2, weight=0)

        # Search Label and Entry
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, sticky="W", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_equipment)  # Bind to search changes
        self.entry_search = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        self.entry_search.grid(row=0, column=1, sticky="EW", padx=(0, 5))

        # Clear search button
        btn_clear_search = ttk.Button(search_frame, text="X", width=3, command=self.clear_search)
        btn_clear_search.grid(row=0, column=2, sticky="E")

        # Create a frame to hold the Listbox and Scrollbar (Row 1)
        listbox_frame = ttk.Frame(self.frame_equipment)
        listbox_frame.grid(row=1, column=0, sticky="NSEW", padx=5, pady=5)
        listbox_frame.rowconfigure(0, weight=1)
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.columnconfigure(1, weight=0)

        # Equipment Listbox
        self.listbox_equipment = tk.Listbox(listbox_frame, height=18)  # Reduced height to make room for search
        self.listbox_equipment.grid(row=0, column=0, sticky="NSEW")

        # Scrollbar for the Listbox
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox_equipment.yview)
        scrollbar.grid(row=0, column=1, sticky="NS")

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
        self.frame_settings = ttk.LabelFrame(self.paned_window, text="Settings")
        self.paned_window.add(self.frame_settings, weight=40)  # Bigger weight for Settings
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
            text="Start Polling", 
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
            text="Stop Polling", 
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
        
        # Units Tab
        self.units_tab = ttk.Frame(self.settings_notebook)
        self.settings_notebook.add(self.units_tab, text="Units")
        
        # ---------------- Units Tab Widgets ----------------
        # Controls frame (row 0)
        units_controls = ttk.Frame(self.units_tab)
        units_controls.grid(row=0, column=0, sticky="EW", padx=5, pady=5)
        
        self.btn_get_units = ttk.Button(units_controls, text="Get units data", command=self.handle_get_units)
        self.btn_get_units.pack(side=tk.LEFT)
        
        # Table (Treeview) with scrollbar (row 1)
        units_table_frame = ttk.Frame(self.units_tab)
        units_table_frame.grid(row=1, column=0, sticky="NSEW", padx=5, pady=5)
        self.units_tab.rowconfigure(1, weight=1)
        self.units_tab.columnconfigure(0, weight=1)
        
        # Configure the table frame to be expandable
        units_table_frame.rowconfigure(0, weight=1)
        units_table_frame.columnconfigure(0, weight=1)
        
        columns = ("unit_id", "unit_name", "driver_type", "driver_addr", "regulator_type", "ip_address", "com_port")
        self.units_tree = ttk.Treeview(units_table_frame, columns=columns, show="headings")
        # Define user-friendly column headers
        column_headers = {
            "unit_id": "Unit ID",
            "unit_name": "Unit Name", 
            "driver_type": "Driver Type",
            "driver_addr": "Driver Address",
            "regulator_type": "Regulator Type",
            "ip_address": "IP Address",
            "com_port": "COM Port"
        }
        
        # Set column widths - make them very wide to ensure ALL text is visible with no cutoff
        column_widths = {
            "unit_id": 200,
            "unit_name": 300, 
            "driver_type": 250,
            "driver_addr": 250,
            "regulator_type": 300,
            "ip_address": 800,  # IP Address column (now before COM Port)
            "com_port": 200     # COM Port column (now after IP Address)
        }
        
        for col in columns:
            self.units_tree.heading(col, text=column_headers.get(col, col))
            # Set columns to auto-size to content with minimum width
            self.units_tree.column(col, anchor="w", minwidth=100, stretch=True)
        
        # Create scrollbars - both vertical and horizontal
        units_v_scrollbar = ttk.Scrollbar(units_table_frame, orient="vertical", command=self.units_tree.yview)
        units_h_scrollbar = ttk.Scrollbar(units_table_frame, orient="horizontal", command=self.units_tree.xview)
        
        # Configure the treeview to work with both scrollbars
        self.units_tree.configure(yscrollcommand=units_v_scrollbar.set, xscrollcommand=units_h_scrollbar.set)
        
        # Use grid layout for proper resizing behavior
        self.units_tree.grid(row=0, column=0, sticky="NSEW")
        units_v_scrollbar.grid(row=0, column=1, sticky="NS")
        units_h_scrollbar.grid(row=1, column=0, sticky="EW")
        
        # Configure grid weights for the table frame
        units_table_frame.columnconfigure(0, weight=1)
        units_table_frame.rowconfigure(0, weight=1)
        units_table_frame.rowconfigure(1, weight=0)  # Horizontal scrollbar row doesn't expand
        
        # Ensure the treeview can expand to show all content
        def on_treeview_resize(event):
            # Update column widths when treeview is resized
            tree_width = event.width
            col_count = len(columns)
            if col_count > 0:
                # Distribute width evenly among columns with minimum width
                col_width = max(200, tree_width // col_count)
                for col in columns:
                    # Get the minimum width for this column
                    min_width = column_widths.get(col, 200)
                    # Use the larger of calculated width or minimum width
                    final_width = max(col_width, min_width)
                    self.units_tree.column(col, width=final_width)
        
        self.units_tree.bind("<Configure>", on_treeview_resize)
        
        # Function to auto-size columns to fit content exactly
        def auto_size_columns():
            """Automatically size columns to fit the biggest word in each column with no extra space."""
            try:
                # Get all items in the treeview
                items = self.units_tree.get_children()
                if not items:
                    return
                
                # Calculate optimal width for each column
                for col in columns:
                    # Start with header width
                    header_text = column_headers.get(col, col)
                    header_width = len(header_text) * 8  # Approximate character width
                    max_width = header_width
                    
                    # Check all values in this column
                    for item in items:
                        value = self.units_tree.set(item, col)
                        if value:
                            # Calculate width needed for this value
                            value_width = len(str(value)) * 8  # Approximate character width
                            max_width = max(max_width, value_width)
                    
                    # Set column width to exactly fit content (minimal padding)
                    optimal_width = max_width + 5  # Minimal padding (reduced from 10 to 5)
                    
                    # Get minimum width for this column (especially important for IP Address)
                    min_width = column_widths.get(col, 200)
                    
                    # Set column width to the larger of: content width or minimum width
                    optimal_width = max(optimal_width, min_width)
                    
                    self.units_tree.column(col, width=optimal_width)
                    
            except Exception as e:
                print(f"Error in auto_size_columns: {e}")
        
        # Store the auto-size function for later use
        self.auto_size_columns = auto_size_columns
        
        # ---------------- Log Frame (Column 2) ----------------
        self.frame_log = ttk.LabelFrame(self.paned_window, text="Log")
        self.paned_window.add(self.frame_log, weight=50)  # Smaller weight for Log
        self.frame_log.rowconfigure(1, weight=1)  # Changed from 0 to 1 to make room for command line
        self.frame_log.columnconfigure(0, weight=1)
        
        # Set better default pane sizes after the window is fully loaded
        def set_default_pane_sizes():
            try:
                # Wait for the window to be fully loaded
                self.root.update_idletasks()
                
                # Get the total width of the paned window
                total_width = self.paned_window.winfo_width()
                if total_width > 200:  # Only set if window has loaded
                    # Calculate proportions based on new weights: Equipment=10, Settings=40, Log=50
                    # Total weight = 10 + 40 + 50 = 100
                    # Equipment: 10/100 = 10% (fixed size)
                    # Settings: 40/100 = 40% (bigger)
                    # Log: 50/100 = 50% (smaller than before)
                    
                    # Set Equipment pane to 10% of total width (fixed size)
                    equipment_width = int(total_width * 0.10)
                    # Set Log pane to 50% of total width (smaller)
                    log_width = int(total_width * 0.50)
                    # Settings pane will get the remaining 40%
                    
                    # Set the sash positions (borders between panes)
                    self.paned_window.sashpos(0, equipment_width)
                    self.paned_window.sashpos(1, total_width - log_width)
                    
            except Exception as e:
                print(f"Error setting pane sizes: {e}")
        
        # Schedule setting pane sizes after the window is fully loaded
        self.root.after(200, set_default_pane_sizes)
        
        # Command Line Frame (Row 0)
        cmd_frame = ttk.Frame(self.frame_log)
        cmd_frame.grid(row=0, column=0, sticky="EW", padx=5, pady=(5, 0))
        cmd_frame.columnconfigure(0, weight=0)
        cmd_frame.columnconfigure(1, weight=1)
        cmd_frame.columnconfigure(2, weight=0)
        
        # Command Line Label and Entry
        ttk.Label(cmd_frame, text="Command:").grid(row=0, column=0, sticky="W", padx=(0, 5))
        self.cmd_var = tk.StringVar()
        self.entry_cmd = ttk.Entry(cmd_frame, textvariable=self.cmd_var, font=("Consolas", 9))
        self.entry_cmd.grid(row=0, column=1, sticky="EW", padx=(0, 5))
        
        # Apply Command Button
        btn_apply_cmd = ttk.Button(cmd_frame, text="Apply", command=self.apply_custom_command, width=8)
        btn_apply_cmd.grid(row=0, column=2, sticky="E")
        
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
            self.log_queue.put(('error', "Error! Cannot find COM ports. Type in the COM port manually."))
        return com_ports











    def _decode_str(self, encoded_str):
        """Simple base64 decode function for obfuscation."""
        import base64
        try:
            return base64.b64decode(encoded_str).decode('utf-8')
        except Exception:
            return encoded_str  # Return original if decode fails





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

    def extract_ip_address(self, text):
        """Extract IP address from text using regex pattern matching."""
        if not text:
            return ''
        
        import re
        # Pattern to match IPv4 addresses
        ip_pattern = r'[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}'
        
        # Find the first IP address in the text
        match = re.search(ip_pattern, text)
        if match:
            ip_address = match.group(0)
            # Always hide 127.0.0.1 addresses (localhost)
            if ip_address.startswith('127.0.0.1'):
                return ''
            return ip_address
        return ''

    def handle_get_units(self):
        """Fetch units data from MySQL and populate the Units table."""
        try:
            if not os.path.exists(self.mysql_exe_path):
                messagebox.showerror("MySQL Not Found", f"mysql.exe not found at {self.mysql_exe_path}")
                return
            
            # Now I understand the relationship:
            # owner in iw_sys_plant_settings matches driver_type in iw_sys_plant_units
            # We need to look for COM port and IP address settings in the value field
            # For IP address, we'll automatically detect any value that matches IPv4 pattern
            # For COM port, we'll look for reception_port setting
            
            query = (
                "SELECT u.unit_id, u.unit_name, u.driver_type, u.driver_addr, u.regulator_type, "
                "COALESCE(com_port_setting.value, '') as com_port, "
                "COALESCE(ip_setting.value, '') as ip_address "
                "FROM iw_sys_plant_units u "
                "LEFT JOIN iw_sys_plant_settings com_port_setting ON u.driver_type = com_port_setting.owner AND com_port_setting.setting = 'reception_port' "
                "LEFT JOIN iw_sys_plant_settings ip_setting ON u.driver_type = ip_setting.owner AND "
                "   (ip_setting.value REGEXP '^[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}$' OR "
                "    ip_setting.value REGEXP 'https?://[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}' OR "
                "    ip_setting.value REGEXP '[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}') "
                "ORDER BY u.unit_id"
            )
            
            cmd = [
                self.mysql_exe_path,
                "-h", "127.0.0.1",
                "-u", "iwmac",
                "-D", "iw_plant_server3",
                "-N", "-B",
                "-e", query,
            ]
            
            def run_query():
                try:
                    # Prevent window popup
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    creationflags = subprocess.CREATE_NO_WINDOW
                    
                    result = subprocess.run(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        shell=False,
                        startupinfo=startupinfo,
                        creationflags=creationflags,
                    )
                    if result.returncode != 0:
                        err = result.stderr.strip() or "Unknown MySQL error"
                        self.log_queue.put(('error', f"MySQL query failed: {err}"))
                        messagebox.showerror("MySQL Error", err)
                        return
                    
                    rows = []
                    for line in result.stdout.splitlines():
                        if not line.strip():
                            continue
                        parts = line.split('\t')
                        # Ensure we have exactly 7 columns (unit_id, unit_name, driver_type, driver_addr, regulator_type, com_port, ip_address)
                        if len(parts) >= 7:
                            # Extract only the IP address from the full value
                            full_ip_value = parts[6]  # The full value from MySQL
                            clean_ip_address = self.extract_ip_address(full_ip_value)
                            
                            # Reorder data to match new column order: unit_id, unit_name, driver_type, driver_addr, regulator_type, ip_address, com_port
                            reordered_row = [
                                parts[0],  # unit_id
                                parts[1],  # unit_name
                                parts[2],  # driver_type
                                parts[3],  # driver_addr
                                parts[4],  # regulator_type
                                clean_ip_address,  # Clean IP address (only the IP, no extra text)
                                parts[5]   # com_port
                            ]
                            rows.append(reordered_row)
                        elif len(parts) >= 5:
                            # If we only get 5 columns, add empty values for ip_address and com_port
                            reordered_row = [
                                parts[0],  # unit_id
                                parts[1],  # unit_name
                                parts[2],  # driver_type
                                parts[3],  # driver_addr
                                parts[4],  # regulator_type
                                '',        # ip_address (empty)
                                ''         # com_port (empty)
                            ]
                            rows.append(reordered_row)
                    
                    def update_table():
                        # Clear existing
                        for item in self.units_tree.get_children():
                            self.units_tree.delete(item)
                        # Insert new
                        for r in rows:
                            self.units_tree.insert('', 'end', values=r)
                        
                        # Auto-size columns to fit content after data is loaded
                        self.root.after(100, self.auto_size_columns)
                    
                    self.root.after(0, update_table)
                    self.log_queue.put(('info', f"Successfully loaded {len(rows)} units with COM port and IP address information!"))
                    
                except Exception as e:
                    self.log_queue.put(('error', f"Error fetching units: {e}"))
                    messagebox.showerror("Error", str(e))
            
            threading.Thread(target=run_query, daemon=True).start()
            
        except Exception as e:
            self.log_queue.put(('error', f"Error starting units fetch: {e}"))
            messagebox.showerror("Error", str(e))

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
