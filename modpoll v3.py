import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font as tkfont
import serial.tools.list_ports
import subprocess
import threading
import os
import queue
import webbrowser  # For opening URLs
import winreg  # For registry access
import re  # For regex matching

import shlex  # For parsing command line arguments
import datetime
import urllib
import urllib.request
import urllib.error

# Set CustomTkinter appearance and smooth animations
ctk.set_appearance_mode("dark")  # Modes: "System" (default), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

# Enable smooth animations - set widget scaling for better performance
ctk.set_widget_scaling(1.0)  # Perfect scaling for smooth animations

# Deactivate automatic dpi awareness for smoother animations (Windows-specific)
try:
    ctk.deactivate_automatic_dpi_awareness()
except:
    pass  # Not critical if it fails


class ModpollingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("ModPolling Tool")
        
        # Set default window size
        self.root.geometry("1400x800")
        self.root.minsize(1200, 700)
        
        # Modern glassmorphism color scheme inspired by the reference design
        self.bg_primary = '#1A1D2E'      # Deep dark navy background
        self.bg_secondary = '#22283A'    # Card backgrounds
        self.bg_tertiary = '#2D3348'     # Input backgrounds
        self.bg_card = '#252A3D'         # Card hover state
        self.text_primary = '#FFFFFF'    # Pure white text
        self.text_secondary = '#9CA3AF'  # Muted gray text
        self.accent_primary = '#8B5CF6'  # Purple accent (main)
        self.accent_secondary = '#6366F1' # Indigo accent
        self.accent_success = '#10B981'  # Green success
        self.accent_warning = '#F59E0B'  # Amber warning
        self.accent_error = '#EF4444'    # Red error
        self.accent_glow = '#A78BFA'     # Lighter purple for glow effects
        
        # Configure CustomTkinter window
        self.root._set_appearance_mode("dark")
        
        # Set window icon if available (optional)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass  # No icon file found, continue without it
        
        self.modpoll_process = None
        self.is_polling = False
        self.log_queue = queue.Queue()
        # Batched terminal writes (thread-safe queue for CMD-like real-time output)
        self.terminal_write_queue = queue.Queue()
        self.terminal_flush_scheduled = False
        self.last_status_update = None  # Throttle status indicator updates
        self.last_status_time = 0  # Time of last status update
        
        # Units data and filter state
        self.units_rows = []
        self.hide_no_baudrate = False
        
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
        
        # Debug toggle for MySQL fetch logs
        self.debug_mysql_logs = False
        
        # Poll attempt tracking
        self.poll_attempt_counter = 0
        self.current_start_reference = None
        self.last_seen_reference = None  # Track last seen [ref]: to detect cycle repeats

        # Initialize blinking variables
        self.blinking = False
        self.blink_job = None
        self.current_color = None
        self.blink_state = False  # To toggle between color and background
        self.blink_count = 0  # To track the number of blinks

        # Define modern status colors - matching new theme
        self.GREEN_COLOR = "#10B981"  # Modern green
        self.GREEN_DIM = "#059669"    # Dimmer green
        self.YELLOW_COLOR = "#F59E0B"  # Modern amber
        self.YELLOW_DIM = "#D97706"    # Dimmer amber
        self.RED_COLOR = "#EF4444"     # Modern red
        self.RED_DIM = "#DC2626"       # Dimmer red

        # Configure ttk styles for modern look
        self.setup_ttk_styles()
        
        self.create_widgets()
        self.update_log()

    def setup_ttk_styles(self):
        """Setup CustomTkinter appearance (no longer using TTK styles)"""
        # CustomTkinter handles styling automatically
        pass
        
    def create_modern_button_styles(self):
        """
        Create modern button styles with hover effects and rounded corners.
        """
        # This method is now integrated into setup_ttk_styles
        pass

    def setup_smooth_combobox_effects(self):
        """
        Add smooth hover and focus effects to all ComboBoxes and Entry fields for better animations.
        """
        # ComboBoxes
        comboboxes = [
            self.cmb_comport,
            self.cmb_baudrate,
            self.cmb_parity,
            self.cmb_databits,
            self.cmb_stopbits
        ]
        
        for combo in comboboxes:
            # Optimize dropdown animations by setting internal properties
            try:
                # Access the dropdown menu and configure for smooth animations
                if hasattr(combo, '_dropdown_menu'):
                    combo._dropdown_menu.configure(
                        tearoff=0,
                        relief='flat',
                        borderwidth=0,
                        activeborderwidth=0
                    )
            except:
                pass  # Not critical
            
            # Add hover effect with smooth border color transition
            combo.bind("<Enter>", lambda e, c=combo: self.on_combobox_hover(c, True))
            combo.bind("<Leave>", lambda e, c=combo: self.on_combobox_hover(c, False))
            
            # Add focus effect
            combo.bind("<FocusIn>", lambda e, c=combo: self.on_combobox_focus(c, True))
            combo.bind("<FocusOut>", lambda e, c=combo: self.on_combobox_focus(c, False))
            
            # Add smooth click animation
            combo.bind("<Button-1>", lambda e, c=combo: self.on_combobox_click(c))
        
        # Entry fields
        entries = [
            self.entry_slave_address,
            self.entry_start_reference,
            self.entry_num_values,
            self.entry_register_data_type,
            self.entry_modbus_tcp,
            self.entry_search,
            self.entry_cmd
        ]
        
        for entry in entries:
            # Add hover effect
            entry.bind("<Enter>", lambda e, ent=entry: self.on_entry_hover(ent, True))
            entry.bind("<Leave>", lambda e, ent=entry: self.on_entry_hover(ent, False))
            
            # Add focus effect
            entry.bind("<FocusIn>", lambda e, ent=entry: self.on_entry_focus(ent, True))
            entry.bind("<FocusOut>", lambda e, ent=entry: self.on_entry_focus(ent, False))
    
    def on_combobox_hover(self, combobox, is_hovering):
        """Smooth hover effect for ComboBoxes"""
        if is_hovering:
            combobox.configure(border_color=self.accent_primary)
        else:
            # Only reset if not focused
            if combobox.cget("border_color") != self.accent_glow:
                combobox.configure(border_color=self.bg_tertiary)
    
    def on_combobox_focus(self, combobox, is_focused):
        """Smooth focus effect for ComboBoxes"""
        if is_focused:
            combobox.configure(
                border_color=self.accent_glow,
                fg_color=self.bg_card
            )
        else:
            combobox.configure(
                border_color=self.bg_tertiary,
                fg_color=self.bg_tertiary
            )
    
    def on_combobox_click(self, combobox):
        """Smooth click animation for ComboBox dropdown"""
        # Immediate visual feedback on click
        combobox.configure(
            border_color=self.accent_secondary,
            button_color=self.accent_secondary
        )
        
        # Restore after short delay for smooth transition
        self.root.after(150, lambda: combobox.configure(
            border_color=self.accent_glow,
            button_color=self.accent_primary
        ))
    
    def wrap_dropdown_with_animation(self, combobox):
        """Wrap combobox with smooth dropdown animation and proper styling"""
        try:
            # Store reference to track this combobox
            if not hasattr(self, '_animated_combos'):
                self._animated_combos = []
                # Start monitoring for new toplevels
                self.monitor_toplevels()
            
            self._animated_combos.append(combobox)
            
            # Style the dropdown menu properly
            self.style_dropdown_menu(combobox)
            
            # Also try direct method override
            if hasattr(combobox, '_open_dropdown_menu'):
                original_open = combobox._open_dropdown_menu
                
                def animated_open():
                    original_open()
                    # Style and animate
                    for delay in [1, 5, 10, 20]:
                        self.root.after(delay, lambda: self.style_and_animate_dropdown())
                
                combobox._open_dropdown_menu = animated_open
        except Exception as e:
            print(f"Wrap error: {e}")
    
    def style_dropdown_menu(self, combobox):
        """Apply dark theme styling to dropdown menu"""
        try:
            if hasattr(combobox, '_dropdown_menu'):
                menu = combobox._dropdown_menu
                # Style the menu with dark colors
                menu.configure(
                    bg=self.bg_secondary,
                    fg=self.text_primary,
                    activebackground=self.accent_primary,
                    activeforeground=self.text_primary,
                    relief='flat',
                    borderwidth=2,
                    activeborderwidth=0,
                    font=("Segoe UI", 11),
                    selectcolor=self.accent_primary  # For radio/check indicators
                )
        except Exception as e:
            pass  # Silently fail if styling not supported
    
    def monitor_toplevels(self):
        """Monitor for new Toplevel windows and animate them"""
        if not hasattr(self, '_known_toplevels'):
            self._known_toplevels = set()
        
        try:
            # Check for new toplevels
            current_toplevels = set()
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    current_toplevels.add(widget)
                    
                    # If this is a new toplevel, it might be a dropdown
                    if widget not in self._known_toplevels:
                        # Animate it
                        self.root.after(1, lambda w=widget: self.animate_new_toplevel(w))
            
            self._known_toplevels = current_toplevels
            
        except:
            pass
        
        # Keep monitoring
        self.root.after(50, self.monitor_toplevels)
    
    def animate_new_toplevel(self, toplevel):
        """Animate and style a newly created toplevel (likely a dropdown)"""
        try:
            if toplevel.winfo_exists() and toplevel.winfo_ismapped():
                # Get geometry
                x = toplevel.winfo_x()
                y = toplevel.winfo_y()
                
                # Check if it looks like a dropdown (small height, positioned near a widget)
                height = toplevel.winfo_height()
                if height < 400:  # Likely a dropdown
                    # Apply dark theme styling
                    self.apply_dark_dropdown_styling(toplevel)
                    
                    # Start animation
                    start_y = y - 15
                    toplevel.geometry(f"+{x}+{start_y}")
                    toplevel.attributes('-alpha', 0.0)
                    self.slide_and_fade_dropdown(toplevel, x, y, start_y)
        except:
            pass
    
    def style_and_animate_dropdown(self):
        """Find dropdown, apply styling, and animate it"""
        try:
            # Look for toplevel windows
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Toplevel) and widget.winfo_ismapped():
                    # Check if it's small enough to be a dropdown
                    try:
                        height = widget.winfo_height()
                        if 20 < height < 400 and widget.winfo_exists():
                            # Only process if not already done
                            if not hasattr(widget, '_dropdown_styled'):
                                widget._dropdown_styled = True
                                
                                # Apply dark styling to the toplevel
                                self.apply_dark_dropdown_styling(widget)
                                
                                # Animate it
                                x = widget.winfo_x()
                                y = widget.winfo_y()
                                start_y = y - 15
                                widget.geometry(f"+{x}+{start_y}")
                                widget.attributes('-alpha', 0.0)
                                self.slide_and_fade_dropdown(widget, x, y, start_y)
                                return
                    except:
                        pass
        except:
            pass
    
    def apply_dark_dropdown_styling(self, toplevel):
        """Apply comprehensive dark theme to dropdown toplevel"""
        try:
            # Style the toplevel window with dark background
            toplevel.configure(bg=self.bg_secondary)
            
            # Remove window decorations for a cleaner look
            try:
                toplevel.overrideredirect(False)  # Keep it manageable
            except:
                pass
            
            # Find and style all child widgets (menu items)
            def style_children(widget):
                try:
                    widget_class = widget.winfo_class()
                    
                    # Style frames
                    if widget_class in ['Frame', 'CTkFrame']:
                        try:
                            widget.configure(bg=self.bg_secondary, fg_color=self.bg_secondary)
                        except:
                            try:
                                widget.configure(bg=self.bg_secondary)
                            except:
                                pass
                    
                    # Style labels (menu items)
                    elif widget_class in ['Label', 'CTkLabel']:
                        try:
                            widget.configure(
                                bg=self.bg_secondary,
                                fg=self.text_primary,
                                font=("Segoe UI", 11)
                            )
                        except:
                            pass
                    
                    # Style buttons if any
                    elif widget_class in ['Button', 'CTkButton']:
                        try:
                            widget.configure(
                                bg=self.bg_secondary,
                                fg=self.text_primary,
                                activebackground=self.accent_primary,
                                activeforeground=self.text_primary
                            )
                        except:
                            pass
                    
                    # Recursively style children
                    for child in widget.winfo_children():
                        style_children(child)
                except:
                    pass
            
            style_children(toplevel)
        except:
            pass
    
    def slide_and_fade_dropdown(self, window, target_x, target_y, start_y):
        """Smooth slide down and fade in animation"""
        steps = 15
        duration = 180  # Total duration in ms
        delay = duration // steps
        
        def animation_step(step):
            try:
                if step <= steps and window.winfo_exists():
                    # Easing function (ease-out cubic)
                    progress = step / steps
                    eased = 1 - pow(1 - progress, 3)
                    
                    # Calculate new position
                    current_y = int(start_y + (target_y - start_y) * eased)
                    
                    # Calculate opacity
                    alpha = progress
                    
                    # Apply transformations
                    window.geometry(f"+{target_x}+{current_y}")
                    window.attributes('-alpha', alpha)
                    
                    # Next step
                    if step < steps:
                        self.root.after(delay, lambda: animation_step(step + 1))
                    else:
                        # Ensure final state
                        window.attributes('-alpha', 1.0)
            except:
                pass
        
        # Start animation
        animation_step(0)
    
    def fade_in_widget(self, widget, duration=200, steps=20):
        """Simple fade-in animation for any widget"""
        delay = duration // steps
        
        def fade_step(step):
            try:
                if step <= steps and widget.winfo_exists():
                    alpha = step / steps
                    widget.attributes('-alpha', alpha)
                    
                    if step < steps:
                        self.root.after(delay, lambda: fade_step(step + 1))
            except:
                pass
        
        fade_step(0)
    
    def on_entry_hover(self, entry, is_hovering):
        """Smooth hover effect for Entry fields"""
        if is_hovering:
            entry.configure(border_color=self.accent_primary)
        else:
            # Only reset if not focused
            if entry.cget("border_color") != self.accent_glow:
                entry.configure(border_color=self.bg_tertiary)
    
    def on_entry_focus(self, entry, is_focused):
        """Smooth focus effect for Entry fields"""
        if is_focused:
            entry.configure(
                border_color=self.accent_glow,
                fg_color=self.bg_card
            )
        else:
            entry.configure(
                border_color=self.bg_tertiary,
                fg_color=self.bg_tertiary
            )
    
    def setup_command_preview_bindings(self):
        """
        Setup event bindings to automatically update command preview when settings change.
        CustomTkinter comboboxes use command parameter, entry fields use KeyRelease.
        """
        # CustomTkinter ComboBoxes - use configure to set command
        self.cmb_comport.configure(command=lambda choice: self.update_command_preview())
        self.cmb_baudrate.configure(command=lambda choice: self.update_command_preview())
        self.cmb_parity.configure(command=lambda choice: self.update_command_preview())
        self.cmb_databits.configure(command=lambda choice: self.update_command_preview())
        self.cmb_stopbits.configure(command=lambda choice: self.update_command_preview())
        
        # Entry fields - use KeyRelease for real-time updates
        self.entry_slave_address.bind('<KeyRelease>', self.update_command_preview)
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
        CustomTkinter handles hover effects automatically.
        This method is kept for compatibility but no longer needed.
        """
        pass
        


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
            self.btn_stop.config(bg="#FF3333")

    def on_stop_hover_leave(self):
        """
        Handle stop button hover leave event - only hover when polling is active.
        """
        if self.is_polling:
            self.btn_stop.config(bg=self.accent_error)

    def ensure_modpoll_exists(self):
        r"""
        Ensure the modpoll directory exists and report if modpoll.exe is missing.
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
            # Inform user that modpoll.exe is missing
            self.log_queue.put(('info', f"modpoll.exe not found. Download it and save to {self.modpoll_path}"))

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
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Main Frame - CustomTkinter Frame
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        # Keep a reference so we can adjust layout per tab
        self.main_frame = main_frame
        main_frame.grid(row=0, column=0, sticky="NSEW", padx=10, pady=10)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_columnconfigure(2, weight=2)

        # Equipment Frame (Column 0) - Modern card style
        self.frame_equipment = ctk.CTkFrame(
            main_frame,
            corner_radius=20,
            fg_color=self.bg_secondary,
            border_width=1,
            border_color=self.bg_tertiary
        )
        self.frame_equipment.grid(row=0, column=0, sticky="NSEW", padx=(0, 10))
        self.frame_equipment.grid_rowconfigure(1, weight=1)
        self.frame_equipment.grid_columnconfigure(0, weight=1)
        
        # Equipment header frame with better styling
        header_frame = ctk.CTkFrame(
            self.frame_equipment,
            fg_color="transparent"
        )
        header_frame.grid(row=0, column=0, pady=(20, 5), padx=20, sticky="EW")
        
        # Equipment header label
        equipment_header = ctk.CTkLabel(
            header_frame,
            text="Equipment Presets",
            font=("Segoe UI", 18, "bold"),
            text_color=self.accent_glow
        )
        equipment_header.pack(side="left")

        # Search Frame - CustomTkinter with enhanced styling
        search_frame = ctk.CTkFrame(
            self.frame_equipment,
            fg_color=self.bg_primary,
            corner_radius=15,
            border_width=2,
            border_color=self.bg_tertiary
        )
        search_frame.grid(row=1, column=0, sticky="EW", padx=20, pady=(10, 15))
        search_frame.grid_columnconfigure(0, weight=1)

        # Search Entry - CustomTkinter with placeholder
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_equipment)  # Bind to search changes
        self.entry_search = ctk.CTkEntry(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="Search equipment...",
            height=45,
            corner_radius=13,
            border_width=0,
            fg_color="transparent",
            text_color=self.text_primary,
            font=("Segoe UI", 12),
            placeholder_text_color=self.text_secondary
        )
        self.entry_search.grid(row=0, column=0, sticky="EW", padx=5, pady=5)

        # Listbox container frame with border
        listbox_container = ctk.CTkFrame(
            self.frame_equipment,
            fg_color=self.bg_primary,
            corner_radius=15,
            border_width=2,
            border_color=self.bg_tertiary
        )
        listbox_container.grid(row=2, column=0, sticky="NSEW", padx=20, pady=(0, 15))
        listbox_container.grid_rowconfigure(0, weight=1)
        listbox_container.grid_columnconfigure(0, weight=1)

        # Equipment Listbox - Enhanced styling
        self.listbox_equipment = tk.Listbox(
            listbox_container,
            height=22,
            bg=self.bg_primary,
            fg=self.text_primary,
            selectbackground=self.accent_primary,
            selectforeground="white",
            font=("Segoe UI", 11),
            borderwidth=0,
            highlightthickness=0,
            activestyle='none',
            relief='flat',
            selectborderwidth=0,
            cursor="hand2"  # Show pointer cursor on hover
        )
        self.listbox_equipment.grid(row=0, column=0, sticky="NSEW", padx=8, pady=8)
        
        # Add scrollbar with custom styling
        scrollbar = ctk.CTkScrollbar(
            listbox_container,
            command=self.listbox_equipment.yview,
            fg_color=self.bg_primary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary
        )
        scrollbar.grid(row=0, column=1, sticky="NS", padx=(0, 8), pady=8)
        self.listbox_equipment.configure(yscrollcommand=scrollbar.set)

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
        btn_select = ctk.CTkButton(
            self.frame_equipment,
            text="Select Equipment",
            command=self.select_equipment,
            height=50,
            corner_radius=15,
            fg_color=self.accent_primary,
            hover_color=self.accent_secondary,
            font=("Segoe UI", 13, "bold"),
            border_width=2,
            border_color=self.accent_glow,
            cursor="hand2"
        )
        btn_select.grid(row=3, column=0, pady=(0, 20), padx=20, sticky="EW")

        # Settings Frame (Column 1) - Modern card style
        self.frame_settings = ctk.CTkFrame(
            main_frame,
            corner_radius=20,
            fg_color=self.bg_secondary,
            border_width=1,
            border_color=self.bg_tertiary
        )
        self.frame_settings.grid(row=0, column=1, sticky="NSEW", padx=(0, 10))
        self.frame_settings.grid_rowconfigure(1, weight=1)
        self.frame_settings.grid_columnconfigure(0, weight=1)
        
        # Settings header - simplified for better performance
        settings_header = ctk.CTkLabel(
            self.frame_settings,
            text="Configuration",
            font=("Segoe UI", 18, "bold"),
            text_color=self.accent_glow
        )
        settings_header.grid(row=0, column=0, pady=(20, 10), padx=20, sticky="W")

        # CustomTkinter Tabview for Settings - Simplified for performance
        self.settings_notebook = ctk.CTkTabview(
            self.frame_settings,
            corner_radius=12,
            fg_color=self.bg_secondary,
            segmented_button_fg_color=self.bg_tertiary,
            segmented_button_selected_color=self.accent_primary,
            segmented_button_selected_hover_color=self.accent_secondary,
            segmented_button_unselected_color=self.bg_tertiary,
            segmented_button_unselected_hover_color=self.bg_card,
            text_color=self.text_primary,
            text_color_disabled=self.text_secondary
        )
        self.settings_notebook.grid(row=1, column=0, sticky="NSEW", padx=15, pady=(0, 15))

        # Add tabs
        self.settings_notebook.add("Basic")
        self.settings_notebook.add("Advanced")
        self.settings_notebook.add("Units")
        
        # Get tab frames
        self.basic_tab = self.settings_notebook.tab("Basic")
        self.advanced_tab = self.settings_notebook.tab("Advanced")
        self.units_tab = self.settings_notebook.tab("Units")

        # ---------------- Basic Settings Widgets (CustomTkinter) ----------------
        # Add some top padding to the basic tab
        self.basic_tab.grid_columnconfigure(1, weight=1)
        
        # COM Port label
        ctk.CTkLabel(
            self.basic_tab,
            text="COM Port:",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=0, sticky="W", padx=15, pady=(15, 10))
        self.cmb_comport = ctk.CTkComboBox(
            self.basic_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            dropdown_fg_color=self.bg_secondary,  # Darker dropdown background
            dropdown_text_color=self.text_primary,
            dropdown_hover_color=self.accent_primary,
            dropdown_font=("Segoe UI", 11),
            state="normal",
            justify="left"
        )
        self.cmb_comport.grid(column=1, row=0, padx=5, pady=5, sticky="W")
        self.btn_refresh = ctk.CTkButton(
            self.basic_tab,
            text="Refresh",
            command=self.refresh_comports,
            width=110,
            height=40,
            corner_radius=10,
            fg_color=self.accent_primary,
            hover_color=self.accent_secondary,
            font=("Segoe UI", 11, "bold"),
            border_width=0
        )
        self.btn_refresh.grid(column=2, row=0, padx=5, pady=5, sticky="W")

        # Baudrate label
        ctk.CTkLabel(
            self.basic_tab,
            text="Baudrate (-b):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=1, sticky="W", padx=15, pady=(10, 10))
        self.cmb_baudrate = ctk.CTkComboBox(
            self.basic_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            values=["2400", "4800", "9600", "19200", "38400", "57600", "115200"],
            dropdown_fg_color=self.bg_secondary,
            dropdown_text_color=self.text_primary,
            dropdown_hover_color=self.accent_primary,
            dropdown_font=("Segoe UI", 11),
            state="normal",
            justify="left"
        )
        self.cmb_baudrate.set("9600")  # Default to "9600"
        self.cmb_baudrate.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Parity label
        ctk.CTkLabel(
            self.basic_tab,
            text="Parity (-p):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=2, sticky="W", padx=15, pady=(10, 10))
        self.cmb_parity = ctk.CTkComboBox(
            self.basic_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            values=["none", "even", "odd"],
            state="readonly",
            dropdown_fg_color=self.bg_secondary,
            dropdown_text_color=self.text_primary,
            dropdown_hover_color=self.accent_primary,
            dropdown_font=("Segoe UI", 11),
            justify="left"
        )
        self.cmb_parity.set("none")  # Default to "none"
        self.cmb_parity.grid(column=1, row=2, padx=5, pady=5, sticky="W")

        # Address label
        ctk.CTkLabel(
            self.basic_tab,
            text="Address (-a):", 
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=3, sticky="W", padx=15, pady=(10, 10))
        self.entry_slave_address = ctk.CTkEntry(
            self.basic_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            placeholder_text="Enter address..."
        )
        self.entry_slave_address.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_slave_address.insert(0, "1")

        # ---------------- Advanced Settings Widgets (CustomTkinter) ----------------
        # Add column configuration for advanced tab
        self.advanced_tab.grid_columnconfigure(1, weight=1)
        
        # Data Bits label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Data Bits (-d):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=0, sticky="W", padx=15, pady=(15, 10))
        self.cmb_databits = ctk.CTkComboBox(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            values=["7", "8"],
            state="readonly",
            dropdown_fg_color=self.bg_secondary,
            dropdown_text_color=self.text_primary,
            dropdown_hover_color=self.accent_primary,
            dropdown_font=("Segoe UI", 11),
            justify="left"
        )
        self.cmb_databits.set("8")  # Default to "8"
        self.cmb_databits.grid(column=1, row=0, padx=5, pady=5, sticky="W")

        # Stop Bits label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Stop Bits (-s):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=1, sticky="W", padx=15, pady=(10, 10))
        self.cmb_stopbits = ctk.CTkComboBox(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            button_color=self.accent_primary,
            button_hover_color=self.accent_secondary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            values=["1", "2"],
            state="readonly",
            dropdown_fg_color=self.bg_secondary,
            dropdown_text_color=self.text_primary,
            dropdown_hover_color=self.accent_primary,
            dropdown_font=("Segoe UI", 11),
            justify="left"
        )
        self.cmb_stopbits.set("1")  # Default to "1"
        self.cmb_stopbits.grid(column=1, row=1, padx=5, pady=5, sticky="W")

        # Start Reference label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Start Reference (-r):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=2, sticky="W", padx=15, pady=(10, 10))
        self.entry_start_reference = ctk.CTkEntry(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            placeholder_text="Enter reference..."
        )
        self.entry_start_reference.grid(column=1, row=2, padx=5, pady=5, sticky="W")
        self.entry_start_reference.insert(0, "100")

        # Number of Registers label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Number of Registers (-c):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=3, sticky="W", padx=15, pady=(10, 10))
        self.entry_num_values = ctk.CTkEntry(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            placeholder_text="Number of registers..."
        )
        self.entry_num_values.grid(column=1, row=3, padx=5, pady=5, sticky="W")
        self.entry_num_values.insert(0, "1")

        # Register Data Type label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Register Data Type (-t):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=4, sticky="W", padx=15, pady=(10, 10))
        self.entry_register_data_type = ctk.CTkEntry(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            placeholder_text="Data type..."
        )
        self.entry_register_data_type.grid(column=1, row=4, padx=5, pady=5, sticky="W")
        self.entry_register_data_type.insert(0, "3")

        # Modbus TCP/IP label
        ctk.CTkLabel(
            self.advanced_tab,
            text="Modbus TCP/IP (-m tcp):",
            text_color=self.text_primary,
            font=("Segoe UI", 11)
        ).grid(column=0, row=5, sticky="W", padx=15, pady=(10, 10))
        self.entry_modbus_tcp = ctk.CTkEntry(
            self.advanced_tab,
            width=220,
            height=40,
            corner_radius=10,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.text_primary,
            font=("Segoe UI", 11),
            placeholder_text="192.168.1.100"
        )
        self.entry_modbus_tcp.grid(column=1, row=5, padx=5, pady=5, sticky="W")

        # Adjust column weights in Advanced Tab for better layout
        self.advanced_tab.columnconfigure(0, weight=1)
        self.advanced_tab.columnconfigure(1, weight=3)

        # ---------------- Polling Buttons and Status Indicator ----------------
        # Button frame - CustomTkinter
        btn_frame = ctk.CTkFrame(self.frame_settings, fg_color="transparent")
        # Keep reference so we can hide it on Units tab
        self.poll_controls_frame = btn_frame
        btn_frame.grid(row=2, column=0, pady=15, padx=15, sticky="EW")

        # Configure grid in btn_frame
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=0)  # For status indicator
        
        # Store custom command arguments
        self.custom_arguments = None
        # Track if command entry has been manually edited by the user
        self.cmd_dirty = False
        # Cache the last generated argument list (avoid re-parsing display string)
        self._last_built_arguments = None

        # Create modern button styles
        self.create_modern_button_styles()

        # Start Polling Button - Modern gradient style
        self.btn_start = ctk.CTkButton(
            btn_frame,
            text="▶  START POLLING",
            command=self.start_polling,
            font=("Segoe UI", 14, "bold"),
            fg_color=self.accent_primary,
            hover_color=self.accent_secondary,
            text_color="white",
            corner_radius=12,
            height=50,
            width=200,
            border_width=2,
            border_color=self.accent_glow
        )
        self.btn_start.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        # Stop Polling Button - Modern muted style
        self.btn_stop = ctk.CTkButton(
            btn_frame,
            text="■  STOP POLLING",
            command=self.stop_polling,
            font=("Segoe UI", 14, "bold"),
            fg_color=self.bg_tertiary,
            hover_color=self.accent_error,
            text_color=self.text_secondary,
            corner_radius=12,
            height=50,
            width=200,
            border_width=1,
            border_color=self.bg_tertiary,
            state="normal"
        )
        self.btn_stop.grid(row=0, column=1, padx=10, pady=5, sticky="e")

        # Store the frame background color
        self.frame_bg_color = self.bg_primary

        # Add hover effects for modern buttons
        self.setup_button_hover_effects()

        # ---------------- Status Indicator ----------------
        # Create a Canvas for the modern circular indicator with glow effect
        self.status_canvas_size = 56
        self.status_canvas = tk.Canvas(
            btn_frame, 
            width=self.status_canvas_size, 
            height=self.status_canvas_size, 
            highlightthickness=0,
            bg=self.bg_secondary
        )
        self.status_canvas.grid(row=0, column=2, padx=(15, 10), pady=5, sticky="n")

        # Draw outer glow circle (will be updated when status changes)
        self.glow_circle = self.status_canvas.create_oval(
            4, 4, self.status_canvas_size-4, self.status_canvas_size-4,
            fill="", outline="#3A3A4E", width=3
        )
        
        # Draw main status circle
        self.circle = self.status_canvas.create_oval(
            12, 12, self.status_canvas_size-12, self.status_canvas_size-12, 
            fill="#3A3A4E",
            outline="#5A5A6E",
            width=2
        )
        
        # ---------------- Units Tab Widgets (CustomTkinter) ----------------
        # Controls frame (row 0)
        units_controls = ctk.CTkFrame(self.units_tab, fg_color="transparent")
        units_controls.grid(row=0, column=0, sticky="EW", padx=15, pady=(15, 10))
        
        self.btn_get_units = ctk.CTkButton(
            units_controls,
            text="Get Units Data",
            command=self.handle_get_units,
            height=45,
            width=170,
            corner_radius=12,
            fg_color=self.accent_primary,
            hover_color=self.accent_secondary,
            text_color="white",
            font=("Segoe UI", 11, "bold"),
            border_width=2,
            border_color=self.accent_glow
        )
        self.btn_get_units.pack(side=tk.LEFT, padx=(0, 10))
        
        # Toggle button to hide/show units without baudrate
        self.btn_toggle_baud_filter = ctk.CTkButton(
            units_controls,
            text="🔍  Show All Units",
            command=self.toggle_hide_no_baudrate,
            width=180,
            height=45,
            corner_radius=12,
            fg_color=self.bg_tertiary,
            hover_color=self.accent_primary,
            text_color=self.text_primary,
            font=("Segoe UI", 11, "bold"),
            border_width=2,
            border_color=self.bg_card
        )
        self.btn_toggle_baud_filter.pack(side=tk.LEFT)

        # Apply preset from selected unit (COM / -b / -p)
        self.btn_apply_unit_preset = ctk.CTkButton(
            units_controls,
            text="Apply Preset",
            command=self.apply_selected_unit_preset,
            width=160,
            height=45,
            corner_radius=12,
            fg_color=self.bg_tertiary,
            hover_color=self.accent_primary,
            text_color=self.text_primary,
            font=("Segoe UI", 11, "bold"),
            border_width=2,
            border_color=self.bg_card
        )
        self.btn_apply_unit_preset.pack(side=tk.LEFT, padx=(10, 0))
        
        # Table (Treeview) with scrollbar (row 1)
        units_table_frame = ctk.CTkFrame(self.units_tab, fg_color=self.bg_primary, corner_radius=10)
        units_table_frame.grid(row=1, column=0, sticky="NSEW", padx=10, pady=(0, 10))
        self.units_tab.grid_rowconfigure(1, weight=1)
        self.units_tab.grid_columnconfigure(0, weight=1)
        
        # Configure the table frame to be expandable
        units_table_frame.rowconfigure(0, weight=1)
        units_table_frame.columnconfigure(0, weight=1)
        
        # Column order: move IP Address to the end (after Parity)
        columns = ("unit_id", "unit_name", "driver_type", "driver_addr", "regulator_type", "com_port", "baudrate", "parity", "ip_address")
        
        # Configure Treeview style for modern dark theme - with proper dark colors
        tree_style = ttk.Style()
        tree_style.theme_use('clam')  # Use clam as base for better styling
        
        tree_style.configure("Modern.Treeview",
                            background=self.bg_primary,
                            foreground=self.text_primary,
                            fieldbackground=self.bg_primary,
                            bordercolor=self.bg_tertiary,
                            borderwidth=1,
                            relief='solid',
                            font=("Segoe UI", 10))
        tree_style.configure("Modern.Treeview.Heading",
                            background=self.bg_tertiary,
                            foreground=self.text_primary,
                            borderwidth=1,
                            bordercolor=self.bg_tertiary,
                            relief='solid',
                            font=("Segoe UI", 10, "bold"))
        tree_style.map("Modern.Treeview",
                      background=[('selected', self.accent_primary)],
                      foreground=[('selected', self.bg_primary)])
        tree_style.map("Modern.Treeview.Heading",
                      background=[('active', self.bg_secondary)],
                      foreground=[('active', self.text_primary)])
        
        # Apply row striping for better readability + clearer separation
        tree_style.configure("Modern.Treeview", rowheight=30)
        tree_style.configure("Modern.Treeview.Item", bordercolor=self.bg_tertiary, borderwidth=1, relief="solid")
        
        self.units_tree = ttk.Treeview(units_table_frame, columns=columns, show="headings", style="Modern.Treeview")
        
        # Configure row tags for striping
        self.units_tree.tag_configure('oddrow', background=self.bg_primary)
        self.units_tree.tag_configure('evenrow', background=self.bg_tertiary)
        # Define user-friendly column headers
        column_headers = {
            "unit_id": "Unit ID",
            "unit_name": "Unit Name", 
            "driver_type": "Driver Type",
            "driver_addr": "Driver Address",
            "regulator_type": "Regulator Type",
            "ip_address": "IP Address",
            "com_port": "COM Port",
            "baudrate": "Baudrate",
            "parity": "Parity"
        }
        
        # Set column widths (minimums). Keep IP Address compact.
        column_widths = {
            "unit_id": 200,
            "unit_name": 300, 
            "driver_type": 250,
            "driver_addr": 250,
            "regulator_type": 300,
            "ip_address": 170,  # IP Address column (keep compact)
            "com_port": 200,
            "baudrate": 160,
            "parity": 140
        }
        
        for col in columns:
            self.units_tree.heading(col, text=column_headers.get(col, col), anchor="center")
            # Set columns to auto-size to content with minimum width and center alignment
            if col == "ip_address":
                # IP should not expand; keep it close to actual IP length
                self.units_tree.column(
                    col,
                    anchor="center",
                    minwidth=120,
                    width=column_widths.get(col, 170),
                    stretch=False
                )
            else:
                self.units_tree.column(col, anchor="center", minwidth=100, stretch=True)
        
        # Create scrollbars - both vertical and horizontal
        units_v_scrollbar = ttk.Scrollbar(units_table_frame, orient="vertical", command=self.units_tree.yview)
        units_h_scrollbar = ttk.Scrollbar(units_table_frame, orient="horizontal", command=self.units_tree.xview)
        
        # ---- Column separator lines (vertical dividers) ----
        # Draw vertical lines between columns using the same white as the frame.
        # (Slightly inset so they match the table border.)
        self.units_column_separators = []
        for _ in range(len(columns) - 1):
            sep = tk.Frame(units_table_frame, bg=self.text_primary, width=1)
            sep.place_forget()
            self.units_column_separators.append(sep)

        def update_units_column_separators():
            """Position vertical separators at current column boundaries (handles horizontal scroll)."""
            try:
                tree = self.units_tree
                # Avoid heavy work when tab isn't visible yet
                if not tree.winfo_ismapped():
                    return

                tree_x = tree.winfo_x()
                tree_y = tree.winfo_y()
                tree_w = tree.winfo_width()
                tree_h = tree.winfo_height()

                # Separator vertical span: match the Treeview widget area (no padding/overlap)
                y_start = tree_y + 1
                y_end = tree_y + tree_h - 1
                sep_h = max(0, y_end - y_start)

                # Convert xview fraction to pixel scroll offset
                total_cols_width = 0
                col_widths_now = []
                for c in columns:
                    w = int(tree.column(c, "width"))
                    col_widths_now.append(w)
                    total_cols_width += w

                x0 = tree.xview()[0] if tree.xview() else 0.0
                scroll_px = int(total_cols_width * x0)

                running = 0
                for i in range(len(columns) - 1):
                    running += col_widths_now[i]
                    boundary_x = tree_x + (running - scroll_px)

                    sep = self.units_column_separators[i]
                    # Only show if within visible tree area
                    if boundary_x <= tree_x or boundary_x >= (tree_x + tree_w):
                        sep.place_forget()
                    else:
                        # Start/stop exactly within the Treeview area
                        sep.place(x=boundary_x, y=y_start, height=sep_h)
            except Exception:
                # Best-effort only
                pass

        self.update_units_column_separators = update_units_column_separators
        # One fast update after the Units tab is rendered
        self.root.after_idle(update_units_column_separators)

        def _xscroll_proxy(first, last):
            units_h_scrollbar.set(first, last)
            update_units_column_separators()

        # Configure the treeview to work with both scrollbars (+ keep separators synced)
        self.units_tree.configure(yscrollcommand=units_v_scrollbar.set, xscrollcommand=_xscroll_proxy)
        
        # Use grid layout for proper resizing behavior with padding
        self.units_tree.grid(row=0, column=0, sticky="NSEW", padx=10, pady=10)
        units_v_scrollbar.grid(row=0, column=1, sticky="NS", pady=10, padx=(0, 10))
        units_h_scrollbar.grid(row=1, column=0, sticky="EW", padx=10, pady=(0, 10))
        
        # Configure grid weights for the table frame
        units_table_frame.columnconfigure(0, weight=1)
        units_table_frame.columnconfigure(1, weight=0)
        units_table_frame.rowconfigure(0, weight=1)
        units_table_frame.rowconfigure(1, weight=0)  # Horizontal scrollbar row doesn't expand
        
        # Ensure the treeview can expand to show all content
        def on_treeview_resize(event):
            # Keep separators aligned after resize (lightweight)
            update_units_column_separators()
        
        self.units_tree.bind("<Configure>", on_treeview_resize)
        # Fill command inputs when selecting a unit row
        self.units_tree.bind('<<TreeviewSelect>>', self.on_units_selection_from_table)
        
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
                # Keep separators aligned after auto-size
                update_units_column_separators()
                    
            except Exception as e:
                print(f"Error in auto_size_columns: {e}")
        
        # Store the auto-size function for later use
        self.auto_size_columns = auto_size_columns

        # ---------------- Log Frame (Column 2) - Modern card style ----------------
        self.frame_log = ctk.CTkFrame(
            main_frame,
            corner_radius=20,
            fg_color=self.bg_secondary,
            border_width=1,
            border_color=self.bg_tertiary
        )
        self.frame_log.grid(row=0, column=2, sticky="NSEW")
        self.frame_log.grid_rowconfigure(2, weight=1)
        self.frame_log.grid_columnconfigure(0, weight=1)
        
        # Log header - with modern styling
        log_header = ctk.CTkLabel(
            self.frame_log,
            text="📟  Terminal Output",
            font=("Segoe UI", 16, "bold"),
            text_color=self.accent_glow
        )
        log_header.grid(row=0, column=0, pady=(20, 10), padx=20, sticky="W")
        
        # Command Line Frame - CustomTkinter
        cmd_frame = ctk.CTkFrame(self.frame_log, fg_color="transparent")
        cmd_frame.grid(row=1, column=0, sticky="EW", padx=15, pady=(0, 10))
        cmd_frame.grid_columnconfigure(0, weight=1)
        
        # Command Entry - CustomTkinter
        self.cmd_var = tk.StringVar()
        mono_font = self.get_best_monospace_font()
        
        self.entry_cmd = ctk.CTkEntry(
            cmd_frame,
            textvariable=self.cmd_var,
            placeholder_text="Command preview...",
            height=40,
            corner_radius=12,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_tertiary,
            text_color=self.accent_glow,
            font=(mono_font, 10),
            placeholder_text_color=self.text_secondary
        )
        self.entry_cmd.grid(row=0, column=0, sticky="EW")
        # Mark command as dirty when user types
        self.entry_cmd.bind('<KeyRelease>', self.on_cmd_entry_changed)
        
        # (Apply button removed; Start Polling now uses edited command directly)
        
        # Determine best monospace font available
        mono_font = self.get_best_monospace_font()
        
        # Log Text - Modern terminal textbox
        self.txt_log = ctk.CTkTextbox(
            self.frame_log,
            width=600,
            height=500,
            corner_radius=12,
            border_width=2,
            border_color=self.bg_tertiary,
            fg_color=self.bg_primary,
            text_color=self.accent_glow,
            font=(mono_font, 10),
            wrap="none",
            activate_scrollbars=True,
            scrollbar_button_color=self.accent_primary,
            scrollbar_button_hover_color=self.accent_secondary
        )
        self.txt_log.grid(column=0, row=2, sticky="NSEW", padx=15, pady=(0, 15))
        
        # Store the internal text widget for tag configuration
        try:
            # CustomTkinter's textbox internal widget needs special handling
            self.txt_log._textbox.tag_configure('error', foreground='#FF4444', font=(mono_font, 10))
            self.txt_log._textbox.tag_configure('info', foreground='#10B981', font=(mono_font, 10))
            self.txt_log._textbox.tag_configure('normal', foreground=self.text_primary)
            # Attempts: smaller + purple (same vibe as Configuration header)
            self.txt_log._textbox.tag_configure('attempt', foreground=self.accent_glow, font=(mono_font, 9))
            # Accent: purple (used for welcome/equipment/config lines)
            self.txt_log._textbox.tag_configure('accent', foreground=self.accent_glow, font=(mono_font, 10))
            # Successful device values, e.g. "[100]: 43"
            self.txt_log._textbox.tag_configure('response_ok', foreground=self.accent_success, font=(mono_font, 10))
        except:
            pass

        # Refresh COM ports
        self.refresh_comports()

        # Footer - Modern gradient style
        footer_frame = ctk.CTkFrame(
            self.root,
            height=60,
            corner_radius=0,
            fg_color=self.bg_secondary,
            border_width=1,
            border_color=self.bg_tertiary
        )
        footer_frame.grid(row=1, column=0, sticky="EW")
        self.root.grid_rowconfigure(1, weight=0)

        # Modpoll guide link - Modern button
        footer_label_left = ctk.CTkButton(
            footer_frame,
            text="📖  Modpoll Guide",
            command=lambda: webbrowser.open(
                "https://iwmac.zendesk.com/hc/en-gb/articles/13020280416796-Installasjon-Modpoll-Guide"
            ),
            fg_color="transparent",
            hover_color=self.bg_card,
            text_color=self.accent_glow,
            font=("Segoe UI", 11),
            height=35,
            corner_radius=10,
            anchor="w"
        )
        footer_label_left.pack(side='left', padx=20, pady=12)

        # Download link - Modern button
        footer_label_center = ctk.CTkButton(
            footer_frame,
            text="⬇️  Download modpoll.exe",
            command=lambda: webbrowser.open(
                "https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/modpoll.exe"
            ),
            fg_color="transparent",
            hover_color=self.bg_card,
            text_color=self.accent_glow,
            font=("Segoe UI", 11),
            height=35,
            corner_radius=10,
            anchor="w"
        )
        footer_label_center.pack(side='left', padx=20, pady=12)

        # (Removed) Copyright label text per request

        # Add event bindings for automatic command preview updates (after all widgets are created)
        self.setup_command_preview_bindings()

        # ComboBox arrow flips up/down + smooth arrow transition
        self.setup_combobox_arrow_behavior()
        # Smooth dropdown open/close (slide + fade) for comboboxes
        self.setup_smooth_combobox_dropdowns()
        
        # Display welcome message
        self.display_welcome_message()

        # Keep Units tab snappy/clean: hide terminal + polling controls when on Units tab
        self._last_settings_tab = None
        self._monitor_settings_tab()

    def normalize_stopbits_value(self, value) -> str:
        """Stop bits must be '1' or '2' for modpoll."""
        try:
            v = str(value).strip()
        except Exception:
            v = ""
        if v in ("1", "2"):
            return v
        # Accept numeric-ish inputs like 1, 2, 1.0
        try:
            iv = int(float(v))
            if iv in (1, 2):
                return str(iv)
        except Exception:
            pass
        return "1"

    def setup_combobox_arrow_behavior(self):
        """Flip ComboBox arrow up when dropdown is open and animate the flip (lightweight)."""
        try:
            import types
        except Exception:
            types = None

        comboboxes = []
        for name in ("cmb_comport", "cmb_baudrate", "cmb_parity", "cmb_databits", "cmb_stopbits"):
            cb = getattr(self, name, None)
            if cb is not None:
                comboboxes.append(cb)

        for cb in comboboxes:
            try:
                if getattr(cb, "_arrow_behavior_patched", False):
                    continue
                cb._arrow_behavior_patched = True

                original_clicked = cb._clicked
                original_dropdown_callback = cb._dropdown_callback

                cb._arrow_anim_job = None
                cb._arrow_monitor_job = None

                def _arrow_geometry(the_cb=cb):
                    try:
                        x = the_cb._apply_widget_scaling(the_cb._current_width - (the_cb._current_height / 2))
                        y = the_cb._apply_widget_scaling(the_cb._current_height / 2)
                        size = the_cb._apply_widget_scaling(the_cb._current_height / 3)
                        return round(x), round(y), round(size)
                    except Exception:
                        return None

                def _set_arrow(opened: bool, t: float = 1.0, the_cb=cb):
                    """t in [0..1] interpolates down->up when opened=True, or up->down when opened=False."""
                    try:
                        geo = _arrow_geometry()
                        if not geo:
                            return
                        x, y, size = geo
                        half = size / 2
                        off = size / 5

                        # Down (▼)
                        d_left_y = y - off
                        d_mid_y = y + off
                        d_right_y = y - off

                        # Up (▲)
                        u_left_y = y + off
                        u_mid_y = y - off
                        u_right_y = y + off

                        if opened:
                            left_y = d_left_y + (u_left_y - d_left_y) * t
                            mid_y = d_mid_y + (u_mid_y - d_mid_y) * t
                            right_y = d_right_y + (u_right_y - d_right_y) * t
                        else:
                            # interpolate up->down
                            left_y = u_left_y + (d_left_y - u_left_y) * t
                            mid_y = u_mid_y + (d_mid_y - u_mid_y) * t
                            right_y = u_right_y + (d_right_y - u_right_y) * t

                        the_cb._canvas.coords("dropdown_arrow",
                                          x - half, left_y,
                                          x, mid_y,
                                          x + half, right_y)
                    except Exception:
                        pass

                def _animate_arrow(opening: bool, the_cb=cb):
                    try:
                        if the_cb._arrow_anim_job is not None:
                            self.root.after_cancel(the_cb._arrow_anim_job)
                            the_cb._arrow_anim_job = None
                    except Exception:
                        the_cb._arrow_anim_job = None

                    steps = 5
                    delay_ms = 14

                    def step(i: int):
                        try:
                            t = i / (steps - 1) if steps > 1 else 1.0
                            _set_arrow(opening, t, the_cb=the_cb)
                            if i < steps - 1:
                                the_cb._arrow_anim_job = self.root.after(delay_ms, lambda: step(i + 1))
                            else:
                                the_cb._arrow_anim_job = None
                        except Exception:
                            the_cb._arrow_anim_job = None

                    step(0)

                # Expose helpers so other code (animated dropdown) can reuse them
                cb._animate_arrow_open = lambda the_cb=cb: _animate_arrow(True, the_cb=the_cb)
                cb._animate_arrow_close = lambda the_cb=cb: _animate_arrow(False, the_cb=the_cb)

                def _monitor_dropdown_until_closed(the_cb=cb):
                    """Keep arrow up while menu is open; reset down when closed."""
                    try:
                        # Works for both native dropdown_menu and our custom popup
                        is_open = False
                        try:
                            is_open = bool(getattr(the_cb, "_animated_popup", None)) and the_cb._animated_popup.winfo_exists()
                        except Exception:
                            is_open = False
                        if not is_open:
                            try:
                                is_open = the_cb._dropdown_menu.winfo_ismapped()
                            except Exception:
                                is_open = False

                        if is_open:
                            the_cb._arrow_monitor_job = self.root.after(60, lambda: _monitor_dropdown_until_closed(the_cb))
                        else:
                            the_cb._arrow_monitor_job = None
                            _animate_arrow(False, the_cb=the_cb)
                    except Exception:
                        the_cb._arrow_monitor_job = None
                        _animate_arrow(False, the_cb=the_cb)

                # Ensure arrow starts down
                _set_arrow(False, 1.0, the_cb=cb)

                def clicked_proxy(self_cb, event=None):
                    # Flip arrow up immediately + animate
                    _animate_arrow(True, the_cb=self_cb)
                    result = original_clicked(event)
                    # Start monitor loop to revert when menu closes
                    try:
                        if self_cb._arrow_monitor_job is not None:
                            self.root.after_cancel(self_cb._arrow_monitor_job)
                        self_cb._arrow_monitor_job = self.root.after(60, lambda: _monitor_dropdown_until_closed(self_cb))
                    except Exception:
                        self_cb._arrow_monitor_job = None
                    return result

                def dropdown_callback_proxy(self_cb, value: str):
                    # On selection, ensure we revert quickly
                    try:
                        if self_cb._arrow_monitor_job is not None:
                            self.root.after_cancel(self_cb._arrow_monitor_job)
                            self_cb._arrow_monitor_job = None
                    except Exception:
                        self_cb._arrow_monitor_job = None
                    _animate_arrow(False, the_cb=self_cb)
                    return original_dropdown_callback(value)

                if types is not None:
                    cb._clicked = types.MethodType(clicked_proxy, cb)
                    cb._dropdown_callback = types.MethodType(dropdown_callback_proxy, cb)
                else:
                    # Fallback (should still work in CPython): assign closures
                    cb._clicked = lambda event=None: clicked_proxy(cb, event)
                    cb._dropdown_callback = lambda value: dropdown_callback_proxy(cb, value)

            except Exception:
                # Best-effort; never break UI
                pass

    def setup_smooth_combobox_dropdowns(self):
        """
        Provide a smoother dropdown transition by replacing the native tk.Menu popup
        with a lightweight CTkToplevel popup (fade + slide). Only for our comboboxes.
        """
        import types

        comboboxes = []
        for name in ("cmb_comport", "cmb_baudrate", "cmb_parity", "cmb_databits", "cmb_stopbits"):
            cb = getattr(self, name, None)
            if cb is not None:
                comboboxes.append(cb)

        for cb in comboboxes:
            try:
                if getattr(cb, "_smooth_dropdown_patched", False):
                    continue
                cb._smooth_dropdown_patched = True

                original_open = cb._open_dropdown_menu

                def open_proxy(self_cb):
                    # Toggle: if already open, close
                    try:
                        popup = getattr(self_cb, "_animated_popup", None)
                        if popup is not None and popup.winfo_exists():
                            self._close_animated_combobox_popup(self_cb)
                            return
                    except Exception:
                        pass

                    # Use animated popup; fallback to native if anything fails
                    try:
                        self._open_animated_combobox_popup(self_cb)
                    except Exception:
                        try:
                            original_open()
                        except Exception:
                            pass

                cb._open_dropdown_menu = types.MethodType(open_proxy, cb)

            except Exception:
                pass

    def _open_animated_combobox_popup(self, cb):
        """Create and animate a dropdown popup for a CTkComboBox."""
        import customtkinter as ctk

        # Close any existing popup on other combobox
        try:
            active = getattr(self, "_active_combo_popup", None)
            if active is not None and active is not cb:
                self._close_animated_combobox_popup(active)
        except Exception:
            pass
        self._active_combo_popup = cb

        values = []
        try:
            values = list(cb.cget("values") or [])
        except Exception:
            try:
                values = list(getattr(cb, "_values", []) or [])
            except Exception:
                values = []

        if not values:
            return

        x = cb.winfo_rootx()
        y = cb.winfo_rooty() + cb.winfo_height()
        w = cb.winfo_width()

        # Base sizing (keep same width as the combobox)
        item_h = 34
        popup_w = w

        # Goal: show ALL options if they fit on screen; otherwise show as many as possible (scroll for the rest).
        # Compute desired height to actually fit all items without showing a scrollbar.
        # Each item has its own height plus our packing paddings.
        per_item = item_h + 4  # pady=2 top + bottom
        desired_h = (len(values) * per_item) + 18  # container padding/borders

        # Decide whether to open up or down based on available space
        open_up = False
        max_space = 420  # fallback
        try:
            screen_h = self.root.winfo_screenheight()
            screen_y = cb.winfo_rooty()
            below_space = max(0, screen_h - (y + 12))
            above_space = max(0, screen_y - 12)

            # Choose the direction with more space
            if above_space > below_space:
                open_up = True
                max_space = above_space
            else:
                open_up = False
                max_space = below_space
        except Exception:
            pass

        # Final height: show everything if possible, else as much as fits (never less than 80)
        target_h = max(80, min(desired_h, max_space))

        popup = ctk.CTkToplevel(self.root)
        cb._animated_popup = popup
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        try:
            popup.attributes("-alpha", 0.0)
        except Exception:
            pass

        # Position with 1px border look
        popup_y = (cb.winfo_rooty() - 1) if open_up else y
        popup.geometry(f"{popup_w}x1+{x}+{popup_y}")

        outer = ctk.CTkFrame(
            popup,
            fg_color=self.bg_secondary,
            border_width=1,
            border_color=self.bg_tertiary,
            corner_radius=10
        )
        outer.pack(fill="both", expand=True)

        scroll = ctk.CTkScrollableFrame(
            outer,
            fg_color=self.bg_secondary,
            corner_radius=10
        )
        scroll.pack(fill="both", expand=True, padx=4, pady=4)

        def pick(v: str):
            try:
                cb.set(v)
            except Exception:
                pass
            try:
                cb._dropdown_callback(v)
            except Exception:
                # If internal callback isn't available, try configured command
                try:
                    cmd = getattr(cb, "_command", None)
                    if cmd:
                        cmd(v)
                except Exception:
                    pass
            self._close_animated_combobox_popup(cb)

        for v in values:
            b = ctk.CTkButton(
                scroll,
                text=v,
                height=item_h,
                fg_color=self.bg_secondary,
                hover_color=self.bg_card,
                text_color=self.text_primary,
                corner_radius=8,
                anchor="w",
                command=lambda vv=v: pick(vv)
            )
            b.pack(fill="x", padx=2, pady=1)

        # Close on escape / focus out
        popup.bind("<Escape>", lambda _e: self._close_animated_combobox_popup(cb))
        popup.bind("<FocusOut>", lambda _e: self._close_animated_combobox_popup(cb))

        # Close when clicking outside (snappy + expected)
        def _is_descendant(widget, ancestor) -> bool:
            try:
                w = widget
                while w is not None:
                    if w == ancestor:
                        return True
                    w = getattr(w, "master", None)
            except Exception:
                return False
            return False

        def _on_root_click(event):
            try:
                # Ignore clicks inside popup or on the combobox itself
                if _is_descendant(event.widget, popup) or _is_descendant(event.widget, cb):
                    return
            except Exception:
                pass
            # Close after idle so the click can still activate the underlying widget
            self.root.after_idle(lambda: self._close_animated_combobox_popup(cb))

        try:
            cb._outside_click_bind_id = self.root.bind("<Button-1>", _on_root_click, add="+")
        except Exception:
            cb._outside_click_bind_id = None

        try:
            popup.focus_force()
        except Exception:
            pass

        # Flip arrow up
        try:
            cb._animate_arrow_open()
        except Exception:
            pass

        # Animate open (fade + slide)
        frames = 8
        step_ms = 14

        def anim_open(i=0):
            t = (i / (frames - 1)) if frames > 1 else 1.0
            h = max(1, int(1 + (target_h - 1) * t))
            if open_up:
                popup.geometry(f"{popup_w}x{h}+{x}+{(cb.winfo_rooty() - h)}")
            else:
                popup.geometry(f"{popup_w}x{h}+{x}+{y}")
            try:
                popup.attributes("-alpha", min(1.0, 0.15 + 0.85 * t))
            except Exception:
                pass
            if i < frames - 1:
                self.root.after(step_ms, lambda: anim_open(i + 1))

        anim_open(0)

    def _close_animated_combobox_popup(self, cb):
        """Animate close and destroy dropdown popup."""
        popup = getattr(cb, "_animated_popup", None)
        if popup is None:
            return
        if not popup.winfo_exists():
            cb._animated_popup = None
            return

        x = popup.winfo_x()
        y = popup.winfo_y()
        w = popup.winfo_width()
        start_h = popup.winfo_height()

        # Animate close
        frames = 6
        step_ms = 14

        def anim_close(i=0):
            t = (i / (frames - 1)) if frames > 1 else 1.0
            h = max(1, int(start_h * (1.0 - t)))
            popup.geometry(f"{w}x{h}+{x}+{y}")
            try:
                popup.attributes("-alpha", max(0.0, 1.0 - t))
            except Exception:
                pass
            if i < frames - 1:
                self.root.after(step_ms, lambda: anim_close(i + 1))
            else:
                # Remove outside-click binding
                try:
                    bind_id = getattr(cb, "_outside_click_bind_id", None)
                    if bind_id:
                        self.root.unbind("<Button-1>", bind_id)
                except Exception:
                    pass
                cb._outside_click_bind_id = None

                try:
                    popup.destroy()
                except Exception:
                    pass
                cb._animated_popup = None
                try:
                    cb._animate_arrow_close()
                except Exception:
                    pass

        anim_close(0)

    def _apply_settings_tab_layout(self, tab_name: str):
        """Make Units tab full-width and hide other UI; restore elsewhere."""
        try:
            if tab_name == "Units":
                # Hide left equipment panel
                try:
                    self.frame_equipment.grid_remove()
                except Exception:
                    pass

                # Hide Terminal Output panel
                try:
                    self.frame_log.grid_remove()
                except Exception:
                    pass

                # Hide Start/Stop + indicator row
                try:
                    self.poll_controls_frame.grid_remove()
                except Exception:
                    pass

                # Expand settings frame to full width (all columns)
                try:
                    self.frame_settings.grid_configure(column=0, columnspan=3, padx=(0, 0))
                except Exception:
                    pass
            else:
                # Restore settings frame size
                try:
                    self.frame_settings.grid_configure(column=1, columnspan=1, padx=(0, 10))
                except Exception:
                    pass

                # Restore polling controls
                try:
                    self.poll_controls_frame.grid()
                except Exception:
                    pass

                # Restore terminal panel
                try:
                    self.frame_log.grid()
                except Exception:
                    pass

                # Restore equipment panel
                try:
                    self.frame_equipment.grid()
                except Exception:
                    pass
        except Exception:
            # Best-effort only; never break UI
            pass

    def _monitor_settings_tab(self):
        """Poll tab selection and apply layout changes (reliable across CustomTkinter versions)."""
        try:
            current = None
            try:
                current = self.settings_notebook.get()
            except Exception:
                current = None

            if current and current != self._last_settings_tab:
                self._last_settings_tab = current
                self._apply_settings_tab_layout(current)
        finally:
            # Light polling for responsiveness without lag
            self.root.after(50, self._monitor_settings_tab)

    def get_best_monospace_font(self):
        """Get the best available monospace font for the terminal"""
        preferred_fonts = [
            "Cascadia Code",
            "Consolas",
            "Courier New",
            "Monaco",
            "Monospace"
        ]
        
        available_fonts = list(tkfont.families())
        
        for font in preferred_fonts:
            if font in available_fonts:
                return font
        
        # Fallback to default monospace
        return "Courier New"
    
    def display_welcome_message(self):
        """Display a welcome message in the log"""
        # Build a clean centered ASCII banner (fixed width so the "square" isn't buggy)
        inner_width = 70  # characters between the borders
        title = "Modpolling Tool"
        subtitle = "Professional Modbus Testing Suite"
        banner_box = "\n".join([
            "╔" + ("═" * inner_width) + "╗",
            "║" + title.center(inner_width) + "║",
            "║" + subtitle.center(inner_width) + "║",
            "╚" + ("═" * inner_width) + "╝",
        ])

        welcome_body = """Welcome to the ModPolling Tool!

Quick Start:
  1. Select equipment from the left panel or configure manually
  2. Choose your COM port or enter TCP/IP address
  3. Click START POLLING to begin communication
  4. Monitor real-time responses in this terminal

Status Indicator: 
  Green  = Successful communication
  Yellow = Warning (checksum errors)
  Red    = Error (timeout/connection issues)

Ready to poll..."""

        self.txt_log.configure(state="normal")
        t = self.txt_log._textbox

        # Insert banner box and center it (white)
        start = t.index("end-1c")
        t.insert("end", banner_box + "\n\n")
        end = t.index("end-1c")
        t.tag_add("welcome_banner", start, end)
        t.tag_configure("welcome_banner", justify="center", foreground=self.text_primary)

        # Insert the rest (white)
        body_start = t.index("end-1c")
        t.insert("end", welcome_body + "\n")
        body_end = t.index("end-1c")
        t.tag_add("welcome_body", body_start, body_end)
        t.tag_configure("welcome_body", foreground=self.text_primary)

        # Colorize status legend words only (Green/Yellow/Red)
        try:
            color_map = {
                "Green": ("welcome_green", self.accent_success),
                "Yellow": ("welcome_yellow", self.accent_warning),
                "Red": ("welcome_red", self.accent_error),
            }

            for word, (tag_name, color) in color_map.items():
                # Find the first occurrence inside the welcome body
                idx = t.search(word, body_start, stopindex=body_end)
                if idx:
                    end_idx = f"{idx}+{len(word)}c"
                    t.tag_add(tag_name, idx, end_idx)
                    t.tag_configure(tag_name, foreground=color)
                    # Ensure colored tags win over the white body tag
                    t.tag_raise(tag_name, "welcome_body")
        except Exception:
            pass
        self.txt_log.configure(state="disabled")
    
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

        # Update the combobox values - CustomTkinter uses configure
        self.cmb_comport.configure(values=combined_com_ports)

        # If there are available COM ports, select the first one
        if combined_com_ports:
            self.cmb_comport.set(combined_com_ports[0])
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

        # Extra safety: if a process is still alive, don't start another
        try:
            if self.modpoll_process and self.modpoll_process.poll() is None:
                self.is_polling = True
                self.update_buttons()
                self.log_queue.put(('info', "Polling is already running."))
                return
        except Exception:
            pass

        # Check if modpoll.exe exists before starting
        if not os.path.exists(self.modpoll_path):
            self.log_queue.put(('error', f"modpoll.exe not found at {self.modpoll_path}."))
            open_dl = messagebox.askyesno(
                "modpoll.exe missing",
                (
                    "modpoll.exe was not found.\n\n"
                    "Open the official download now?\n\n"
                    f"After downloading, save it to:\n{self.modpoll_path}"
                ),
            )
            if open_dl:
                webbrowser.open("https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/modpoll.exe")
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
        stopbits_raw = self.cmb_stopbits.get()
        stopbits = self.normalize_stopbits_value(stopbits_raw)
        if stopbits != str(stopbits_raw).strip():
            try:
                self.cmb_stopbits.set(stopbits)
            except Exception:
                pass
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

        # Determine arguments priority: user-edited command > saved custom > generated
        arguments = None
        if self.custom_arguments:
            arguments = self.custom_arguments
            self.log_queue.put(('info', f"Using custom command arguments: {' '.join(arguments)}"))
        else:
            cmd_text_current = self.cmd_var.get().strip()
            if self.cmd_dirty and cmd_text_current:
                try:
                    # Windows-safe parsing (preserve backslashes like \\\\.\\COM10)
                    arguments = shlex.split(cmd_text_current, posix=False)
                    # Remove "modpoll" if it's the first argument (for display purposes)
                    if arguments and arguments[0].lower() == "modpoll":
                        arguments = arguments[1:]
                    self.log_queue.put(('info', f"Using edited command: {cmd_text_current}"))
                except Exception as e:
                    self.log_queue.put(('error', f"Invalid command format: {e}"))
                    arguments = None
            if arguments is None:
                # Build arguments directly (avoid shlex re-parsing which can break \\\\.\\COM10)
                self.build_and_display_command()
                if self._last_built_arguments:
                    arguments = list(self._last_built_arguments)
                if not arguments:
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
        self.last_seen_reference = None  # Reset cycle detection

        # Prevent double-start race: mark polling active BEFORE starting thread
        self.is_polling = True
        self.update_buttons()

        # Start polling thread
        threading.Thread(target=self.run_modpoll, args=(arguments, com_port, baudrate, parity, databits, stopbits, adresse, start_reference, num_registers, register_data_type), daemon=True).start()

    def _get_register_type_description(self, register_data_type):
        """Map register data type number to description."""
        type_map = {
            "0": "coil",
            "1": "discrete input",
            "3": "16-bit register, input register table",
            "4": "16-bit register, holding register table",
        }
        return type_map.get(str(register_data_type), f"type {register_data_type}")

    def run_modpoll(self, arguments, com_port, baudrate, parity, databits, stopbits, adresse, start_reference, num_registers, register_data_type):
        # start_polling() already set is_polling/update_buttons to avoid double-start races
        self._write_to_terminal("Polling started...", 'info')

        try:
            # Use the hardcoded modpoll path
            modpoll_path = self.modpoll_path
            if not os.path.exists(modpoll_path):
                self._write_to_terminal(f"Error: modpoll.exe not found at {modpoll_path}", 'error')
                return

            cmd = [modpoll_path] + arguments

            # Determine protocol type (RTU or TCP/IP)
            use_tcp = "-m" in arguments and "tcp" in arguments
            protocol_type = "Modbus TCP/IP" if use_tcp else "Modbus RTU"
            
            # Format serial port configuration
            if use_tcp:
                # For TCP/IP, find the IP address in arguments
                tcp_addr = ""
                if "-m" in arguments:
                    tcp_idx = arguments.index("-m")
                    if tcp_idx + 1 < len(arguments) and arguments[tcp_idx + 1] == "tcp":
                        if tcp_idx + 2 < len(arguments):
                            tcp_addr = arguments[tcp_idx + 2]
                serial_config = f"TCP/IP address: {tcp_addr}" if tcp_addr else "TCP/IP"
            else:
                serial_config = f"{com_port}, {baudrate}, {databits}, {stopbits}, {parity}"
            
            # Get register type description
            register_type_desc = self._get_register_type_description(register_data_type)

            # Display configuration immediately (white text like modpoll output)
            self._write_to_terminal(f"Protocol configuration: {protocol_type}", 'normal')
            self._write_to_terminal(f"Slave configuration: Address = {adresse}, start reference = {start_reference}, count = {num_registers}", 'normal')
            if use_tcp:
                self._write_to_terminal(f"TCP/IP configuration: {tcp_addr if tcp_addr else 'TCP/IP'}", 'normal')
            else:
                self._write_to_terminal(f"Serial port configuration: {serial_config}", 'normal')
            self._write_to_terminal(f"Data type: {register_type_desc}", 'normal')
            self._write_to_terminal("Protocol opened successfully.", 'normal')

            # Log the command (mask the full path to show just 'modpoll')
            masked_cmd = ['modpoll'] + arguments
            # Show these in white (normal) for readability
            self._write_to_terminal(f"Running command: {' '.join(masked_cmd)}", 'normal')
            self._write_to_terminal(f"Parameters: COM={com_port}, Baud={baudrate}, Parity={parity}, DataBits={databits}, StopBits={stopbits}, Addr={adresse}, Ref={start_reference}, Count={num_registers}, Type={register_data_type}", 'normal')

            # Prevent a new window from opening
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

            # Ensure unbuffered output from modpoll.exe for real-time display
            env = dict(os.environ)
            env['PYTHONUNBUFFERED'] = '1'
            
            self.modpoll_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,  # merge for ordered output
                text=True,
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags,
                bufsize=1,  # line-buffered (unbuffered mode)
                env=env,
            )

            # Start thread to read stdout immediately (matches example pattern for real-time output)
            stdout_thread = threading.Thread(
                target=self.read_stream, args=(self.modpoll_process.stdout,), daemon=True
            )
            stdout_thread.start()

            # Wait for the process to finish
            self.modpoll_process.wait()

            # Wait for the reader thread to finish (ensures all output is processed)
            stdout_thread.join()

        except Exception as e:
            self._write_to_terminal(f"Error running Modpoll: {str(e)}", 'error')
        finally:
            self.is_polling = False
            try:
                self.root.after(0, self.update_buttons)
            except Exception:
                pass
            # (Removed) Polling finished.

    def _increment_attempt(self):
        self.poll_attempt_counter += 1
        # Log every attempt in white (direct to UI like CMD)
        self._write_to_terminal(f"Attempt {self.poll_attempt_counter}", 'normal')

    def _write_to_terminal(self, message, tag=None):
        """Write instantly to terminal (no batch delay)."""
        try:
            self.txt_log.configure(state="normal")
            if tag:
                self.txt_log._textbox.insert("end", f"{message}\n", tag)
            else:
                self.txt_log._textbox.insert("end", f"{message}\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
            self.root.update_idletasks()
        except Exception as e:
            self.log_queue.put((tag or 'normal', message))

    def _flush_terminal_writes(self):
        """Flush all queued terminal writes in one UI update (like CMD)."""
        self.terminal_flush_scheduled = False
        try:
            t = self.txt_log._textbox
            self.txt_log.configure(state="normal")
            
            # Process queued writes in small batches to keep UI responsive
            count = 0
            max_batch = 20  # Smaller batches = more responsive UI
            while count < max_batch:
                try:
                    message, tag = self.terminal_write_queue.get_nowait()
                    if tag:
                        t.insert("end", f"{message}\n", tag)
                    else:
                        t.insert("end", f"{message}\n")
                    count += 1
                except Exception:
                    break
            
            if count > 0:
                self.txt_log.see("end")
                self.txt_log.configure(state="disabled")
                # If more items queued, schedule another flush immediately
                if not self.terminal_write_queue.empty():
                    self.terminal_flush_scheduled = True
                    self.root.after(0, self._flush_terminal_writes)
        except Exception:
            try:
                self.txt_log.configure(state="disabled")
            except Exception:
                pass

    def read_stream(self, stream):
        """Read modpoll output and write directly to terminal like CMD (real-time)."""
        try:
            # Use iter(stream.readline, '') for line-by-line reading (matches example pattern)
            for raw in iter(stream.readline, ""):
                if not raw:  # EOF
                    break
                line = raw.rstrip()  # Use rstrip() instead of strip() to preserve leading whitespace if needed
                if not line:
                    continue

                lower_line = line.lower()

                # Skip the polling slave line - don't show it
                if "polling slave (ctrl-c to stop) ..." in lower_line:
                    continue

                # Skip modpoll header and copyright information
                if any(header_text in line for header_text in [
                    "modpoll - FieldTalk(tm) Modbus(R) Polling Utility",
                    "Copyright (c) 2002-2004 FOCUS Software Engineering Pty Ltd",
                    "Getopt Library Copyright (C) 1987-1997\tFree Software Foundation, Inc."
                ]):
                    continue

                # Skip "Protocol opened successfully" (we already display it immediately after configuration)
                if "protocol opened successfully" in lower_line:
                    self.root.after_idle(lambda: self.trigger_status_indicator('green'))
                    continue

                # Handle time-outs (Reply / Send)
                if "time-out!" in lower_line:
                    self._increment_attempt()
                    if "send time-out!" in lower_line:
                        self._write_to_terminal("Send time-out! - No response from device", 'error')
                    else:
                        self._write_to_terminal("Time-out - No response from device", 'error')
                    self.root.after_idle(lambda: self.trigger_status_indicator('red'))
                    continue

                # Handle "serial port already open"
                if "serial port already open" in lower_line:
                    self._write_to_terminal("Port already open - Stop plant server first", 'error')
                    self.root.after_idle(lambda: self.trigger_status_indicator('red'))
                    continue

                # Handle "port or socket open error!"
                if "port or socket open error!" in lower_line:
                    self._write_to_terminal("Port error - Check connection", 'error')
                    self.root.after_idle(lambda: self.trigger_status_indicator('red'))
                    continue

                # Handle "checksum error"
                if "checksum error" in lower_line:
                    base_message = "Checksum error"
                    count = self.message_counts.get(base_message, 0) + 1
                    self.message_counts[base_message] = count
                    numbered_message = f"{base_message} [{count}] - Data corruption"
                    self._write_to_terminal(numbered_message, 'error')
                    self.root.after_idle(lambda: self.trigger_status_indicator('yellow'))
                    continue

                # Handle "illegal function exception response!"
                if "illegal function exception response!" in lower_line:
                    self._increment_attempt()
                    self._write_to_terminal("Illegal Function Exception Response!", 'info')
                    self.root.after_idle(lambda: self.trigger_status_indicator('green'))
                    continue

                # Handle "illegal data address exception response!"
                if "illegal data address exception response!" in lower_line:
                    self._increment_attempt()
                    self._write_to_terminal("Illegal Data Address Exception Response! - Device is responding", 'info')
                    self.root.after_idle(lambda: self.trigger_status_indicator('green'))
                    continue

                # Handle "illegal data value exception response!"
                if "illegal data value exception response!" in lower_line:
                    self._increment_attempt()
                    self._write_to_terminal("Illegal Data Value Exception Response!", 'info')
                    self.root.after_idle(lambda: self.trigger_status_indicator('green'))
                    continue

                # Skip Modpoll config lines (we already display them immediately when polling starts)
                if lower_line.startswith(("protocol configuration:", "slave configuration:", "serial port configuration:", "data type:", "tcp/ip configuration:")):
                    continue

                # Can't reach slave (TCP/IP issue) - accent purple
                if "can't reach slave" in lower_line or "cant reach slave" in lower_line:
                    self._write_to_terminal("Can't reach slave (check ip address)!", 'accent')
                    continue

                # Handle lines like "[100]: 70" - successful device response
                if re.match(r'^\[\s*\d+\s*\]\s*:', line):
                    # Extract the reference number from the line
                    ref_match = re.match(r'^\[(\d+)\]\s*:', line)
                    if ref_match:
                        ref_num = ref_match.group(1)
                        # Detect start of a new polling cycle: same reference appears again
                        # (works even if user changed -r, since we track what we actually see)
                        if self.last_seen_reference is not None and ref_num == self.last_seen_reference:
                            self._increment_attempt()
                        self.last_seen_reference = ref_num
                    self._write_to_terminal(f"{line} - Device is responding", 'response_ok')
                    # Throttle status updates: only update once per second for fast responses
                    import time
                    now = time.time()
                    if self.last_status_update != 'green' or (now - self.last_status_time) >= 1.0:
                        self.last_status_update = 'green'
                        self.last_status_time = now
                        self.root.after_idle(lambda: self.trigger_status_indicator('green'))
                    continue

                # Default
                self._write_to_terminal(line, 'normal')
        finally:
            # Close stream properly (matches example pattern)
            try:
                stream.close()
            except Exception:
                pass

    def stop_polling(self):
        if self.modpoll_process and self.modpoll_process.poll() is None:
            try:
                self.modpoll_process.terminate()
                self.modpoll_process.wait(timeout=5)
                # Show in white
                self._write_to_terminal("Polling stopped.", 'normal')
            except subprocess.TimeoutExpired:
                self.modpoll_process.kill()
                self._write_to_terminal("Polling forcefully stopped.", 'info')
            except Exception as e:
                self._write_to_terminal(f"Error stopping polling: {str(e)}", 'error')
            finally:
                self.modpoll_process = None
                self.is_polling = False
                self.update_buttons()
        else:
            self.log_queue.put(('info', "No polling process to stop."))

    def update_buttons(self):
        if self.is_polling:
            # Disable start button and change appearance
            self.btn_start.configure(
                state="disabled",
                fg_color=self.bg_tertiary,
                text_color=self.text_secondary,
                border_color=self.bg_tertiary
            )
            # Enable stop button with active state
            self.btn_stop.configure(
                state="normal",
                fg_color=self.accent_error,
                hover_color="#DC2626",
                text_color="white",
                border_color=self.accent_error
            )
        else:
            # Enable start button and restore appearance
            self.btn_start.configure(
                state="normal",
                fg_color=self.accent_primary,
                hover_color=self.accent_secondary,
                text_color="white",
                border_color=self.accent_glow
            )
            # Show stop button as disabled
            self.btn_stop.configure(
                state="normal",
                fg_color=self.bg_tertiary,
                hover_color=self.accent_error,
                text_color=self.text_secondary,
                border_color=self.bg_tertiary
            )

    def append_log(self, message, tag=None):
        # CustomTkinter textbox with proper tag support
        self.txt_log.configure(state="normal")
        if tag:
            self.txt_log._textbox.insert("end", f"{message}\n", tag)
        else:
            self.txt_log._textbox.insert("end", f"{message}\n")
        
        # Always scroll to end for real-time feel
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def append_log_direct(self, message, tag=None):
        """Direct logging like a real terminal with proper color tags"""
        try:
            self.txt_log.configure(state="normal")
            
            # Insert with tag (supports custom tags like 'attempt')
            if tag:
                self.txt_log._textbox.insert("end", f"{message}\n", tag)
            else:
                self.txt_log._textbox.insert("end", f"{message}\n")
            
            # Scroll to end like a real terminal
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
            
            # Force immediate update like native CMD
            self.root.update_idletasks()
        except Exception as e:
            # Fallback to queue if direct logging fails
            self.log_queue.put((tag or 'normal', message))

    def update_log(self):
        """
        Legacy queue processor (kept for compatibility but no longer used).
        Terminal now writes directly like CMD via _write_to_terminal().
        """
        # Still process any remaining queue items (for backwards compatibility)
        try:
            while not self.log_queue.empty():
                try:
                    item = self.log_queue.get_nowait()
                    if isinstance(item, tuple):
                        tag, message = item
                        if tag == "status":
                            self.trigger_status_indicator(message)
                        else:
                            self._write_to_terminal(message, tag)
                    else:
                        self._write_to_terminal(item)
                except Exception:
                    break
        except Exception:
            pass
        
        # Keep running but less frequently (only for any legacy queue items)
        self.root.after(100, self.update_log)



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
            
        # Display equipment selection in terminal (white)
        self.append_log_direct("═" * 70, 'normal')
        self.append_log_direct(f"📦 EQUIPMENT SELECTED: {selected_equipment}", 'normal')
        self.append_log_direct("─" * 70, 'normal')
            
        # Clear custom command when selecting equipment preset
        if self.custom_arguments:
            self.custom_arguments = None
            self.append_log_direct("⚠ Custom command cleared - using equipment preset", 'accent')
            
        settings = self.equipment_settings.get(selected_equipment, {})
        baudrate = settings.get("baudrate", "9600")
        parity = settings.get("parity", "none")
        stop_bits = settings.get("stop_bits", "1")
        data_bits = settings.get("data_bits", "8")
        
        # Get Modbus settings with defaults
        slave_address = settings.get("address", "1")
        start_reference = settings.get("start_reference", "100")
        num_registers = settings.get("num_registers", "1")
        register_type = settings.get("register_type", "3")

        # Display settings being applied
        self.append_log_direct("⚙ Applying Equipment Settings:", 'normal')
        self.append_log_direct(f"  • Baudrate:        {baudrate}", 'normal')
        self.append_log_direct(f"  • Parity:          {parity}", 'normal')
        self.append_log_direct(f"  • Data Bits:       {data_bits}", 'normal')
        self.append_log_direct(f"  • Stop Bits:       {stop_bits}", 'normal')
        self.append_log_direct(f"  • Address:         {slave_address}", 'normal')
        self.append_log_direct(f"  • Start Reference: {start_reference}", 'normal')
        self.append_log_direct(f"  • Num Registers:   {num_registers}", 'normal')
        self.append_log_direct(f"  • Register Type:   {register_type}", 'normal')

        # Update Baudrate (CustomTkinter uses cget for values)
        try:
            baudrate_values = self.cmb_baudrate.cget("values")
            if baudrate in baudrate_values:
                self.cmb_baudrate.set(baudrate)
            else:
                self.cmb_baudrate.set(baudrate)
        except:
            self.cmb_baudrate.set(baudrate)

        # Update Parity
        try:
            parity_values = self.cmb_parity.cget("values")
            if parity in parity_values:
                self.cmb_parity.set(parity)
            else:
                self.cmb_parity.set("none")
        except:
            self.cmb_parity.set("none")

        # Update Stop Bits
        try:
            stopbits_values = self.cmb_stopbits.cget("values")
            if stop_bits in stopbits_values:
                self.cmb_stopbits.set(stop_bits)
            else:
                self.cmb_stopbits.set("1")
        except:
            self.cmb_stopbits.set("1")

        # Update Data Bits
        try:
            databits_values = self.cmb_databits.cget("values")
            if data_bits in databits_values:
                self.cmb_databits.set(data_bits)
            else:
                self.cmb_databits.set("8")
        except:
            self.cmb_databits.set("8")
        
        # Update Slave Address entry field
        self.entry_slave_address.delete(0, tk.END)
        self.entry_slave_address.insert(0, slave_address)
        
        # Update Start Reference entry field
        self.entry_start_reference.delete(0, tk.END)
        self.entry_start_reference.insert(0, start_reference)
        
        # Update Number of Registers entry field
        self.entry_num_values.delete(0, tk.END)
        self.entry_num_values.insert(0, num_registers)
        
        # Update Register Data Type entry field
        self.entry_register_data_type.delete(0, tk.END)
        self.entry_register_data_type.insert(0, register_type)

        # Build and display the full command immediately
        self.build_and_display_command()

        # Display success message
        self.append_log_direct("✓ Settings updated successfully!", 'info')

        # Switch to Basic tab after selection (CTkTabview uses .set() method)
        self.settings_notebook.set("Basic")

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
            self.listbox_equipment.itemconfig(tk.END, fg=self.text_secondary)

    def clear_search(self):
        """Clear search field and show all equipment"""
        self.search_var.set("")
        self.filter_equipment()

    # ---------------- Units Table Helpers ----------------
    def toggle_hide_no_baudrate(self):
        """Toggle filter to hide/show units without baudrate and refresh table."""
        self.hide_no_baudrate = not self.hide_no_baudrate
        if self.hide_no_baudrate:
            self.btn_toggle_baud_filter.configure(
                text="🔍  Show All Units",
                fg_color=self.accent_primary,
                hover_color=self.accent_secondary,
                border_color=self.accent_glow
            )
        else:
            self.btn_toggle_baud_filter.configure(
                text="✅  Modbus-Supported",
                fg_color=self.bg_tertiary,
                hover_color=self.accent_primary,
                border_color=self.bg_card
            )
        self.refresh_units_table()

    def set_units_rows(self, rows):
        """Store rows and refresh table with current filter."""
        try:
            self.units_rows = rows or []
        except Exception:
            self.units_rows = []
        self.refresh_units_table()

    def refresh_units_table(self):
        """Rebuild the Units table from stored rows, applying current filters."""
        try:
            # Determine filtered view
            if self.hide_no_baudrate:
                filtered = [r for r in (self.units_rows or []) if len(r) > 7 and str(r[7]).strip()]
            else:
                filtered = list(self.units_rows or [])

            # Clear existing
            for item in self.units_tree.get_children():
                self.units_tree.delete(item)

            # Insert filtered with row striping
            for idx, r in enumerate(filtered):
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                self.units_tree.insert('', 'end', values=r, tags=(tag,))

            # Auto-size columns after data load
            self.root.after(100, self.auto_size_columns)
        except Exception:
            # Non-fatal UI update error; ignore
            pass

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
        stopbits_raw = self.cmb_stopbits.get()
        stopbits = self.normalize_stopbits_value(stopbits_raw)
        # If something (like custom COM input) accidentally stomped this value, snap it back.
        if stopbits != str(stopbits_raw).strip():
            try:
                self.cmb_stopbits.set(stopbits)
            except Exception:
                pass
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
                
                # Build FULL command with ALL parameters always visible
                arguments = [
                    formatted_com_port,
                    f"-b{baudrate}",
                    f"-p{parity}",
                    f"-d{databits}",
                    f"-s{stopbits}",
                    f"-a{adresse}",
                    f"-r{start_reference}",
                    f"-c{num_registers}",
                    f"-t{register_data_type}"
                ]
            else:
                # Show placeholder with ALL parameters
                arguments = [
                    "[COM_PORT]",
                    f"-b{baudrate}",
                    f"-p{parity}",
                    f"-d{databits}",
                    f"-s{stopbits}",
                    f"-a{adresse}",
                    f"-r{start_reference}",
                    f"-c{num_registers}",
                    f"-t{register_data_type}"
                ]
        
        # Update command line display (only in Terminal Output box, not in log)
        cmd_string = ' '.join(arguments)
        # Include "modpoll" at the beginning for the Terminal Output display
        full_command = f"modpoll {cmd_string}"
        self.cmd_var.set(full_command)

        # Cache list form to avoid shlex issues with Windows backslashes (e.g. \\\\.\\COM10)
        self._last_built_arguments = list(arguments)
        
        # Programmatic update: not dirty
        self.cmd_dirty = False

    def on_cmd_entry_changed(self, event=None):
        """Mark the command entry as edited by the user."""
        self.cmd_dirty = True

    def on_units_selection_from_table(self, event=None):
        """When selecting a unit in the Units table, fill COM/baudrate/parity/address and update command."""
        try:
            selection = self.units_tree.selection()
            if not selection:
                return
            item_id = selection[0]
            values = self.units_tree.item(item_id, 'values')
            if not values or len(values) < 6:
                return
            # Columns: unit_id(0), unit_name(1), driver_type(2), driver_addr(3), regulator_type(4), com_port(5), baudrate(6), parity(7), ip_address(8)
            driver_addr = values[3] if len(values) > 3 else ''
            com_port = values[5] if len(values) > 5 else ''
            baudrate = values[6] if len(values) > 6 else ''
            parity = values[7] if len(values) > 7 else ''

            # Extract numeric address: prefer part after '_' in patterns like '1_93'
            address = ''
            if driver_addr:
                parts = str(driver_addr).split('_')
                if len(parts) == 2 and parts[1].isdigit():
                    address = parts[1]
                elif str(driver_addr).isdigit():
                    address = str(driver_addr)
                else:
                    import re as _re
                    m = _re.search(r"(\d+)$", str(driver_addr))
                    if m:
                        address = m.group(1)

            # Apply into UI fields if present
            if com_port:
                self.cmb_comport.set(com_port)
            if baudrate:
                self.cmb_baudrate.set(baudrate)
            if parity and parity in ("none", "even", "odd"):
                self.cmb_parity.set(parity)
            if address:
                self.entry_slave_address.delete(0, tk.END)
                self.entry_slave_address.insert(0, address)

            # Clear custom arguments and rebuild command preview
            self.custom_arguments = None
            self.build_and_display_command()

            # Stay on Units tab; do not switch tabs
        except Exception as _e:
            # Best-effort; do not interrupt UI for minor errors
            pass

    def apply_selected_unit_preset(self):
        """Apply COM/baud/parity preset from currently selected Units row."""
        try:
            selection = self.units_tree.selection()
            if not selection:
                messagebox.showinfo("Select a Unit", "Select a unit row first, then click Apply Preset.")
                return

            item_id = selection[0]
            values = self.units_tree.item(item_id, 'values') or []
            if len(values) < 8:
                messagebox.showwarning("Missing Data", "Selected unit row does not contain preset data.")
                return

            # Columns: unit_id(0), unit_name(1), driver_type(2), driver_addr(3), regulator_type(4),
            #          com_port(5), baudrate(6), parity(7), ip_address(8)
            com_port = values[5] if len(values) > 5 else ''
            baudrate = values[6] if len(values) > 6 else ''
            parity = values[7] if len(values) > 7 else ''
            ip_address = values[8] if len(values) > 8 else ''

            # Normalize COM port to "COMx" if it's numeric
            try:
                if isinstance(com_port, str) and com_port.strip().isdigit():
                    com_port = f"COM{com_port.strip()}"
            except Exception:
                pass

            # Apply into UI fields
            if com_port:
                self.entry_modbus_tcp.delete(0, tk.END)
                self.cmb_comport.set(com_port)
            elif ip_address:
                # TCP unit: apply IP instead
                self.cmb_comport.set("")
                self.entry_modbus_tcp.delete(0, tk.END)
                self.entry_modbus_tcp.insert(0, ip_address)

            if baudrate:
                self.cmb_baudrate.set(str(baudrate).strip())

            if parity:
                self.cmb_parity.set(self.normalize_parity_value(str(parity).strip()))

            # Rebuild command preview
            self.custom_arguments = None
            self.build_and_display_command()

            self.append_log_direct("Applied unit preset (COM / baud / parity).", "info")

            # Jump back to Basic tab automatically
            try:
                self.settings_notebook.set("Basic")
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def apply_custom_command(self):
        """Apply custom command from the command line entry"""
        custom_cmd = self.cmd_var.get().strip()
        if custom_cmd:
            # Parse the custom command into arguments
            try:
                # Split the command into arguments, handling quotes properly
                import shlex
                # Windows-safe parsing (preserve backslashes like \\\\.\\COM10)
                self.custom_arguments = shlex.split(custom_cmd, posix=False)
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

    def normalize_parity_value(self, value):
        """Normalize various parity encodings to user-friendly names (none/even/odd)."""
        try:
            s = (value or '').strip()
        except Exception:
            s = ''
        lower = s.lower()
        # Numeric mappings commonly used: 0=none, 1=odd, 2=even
        if lower in ('0', 'none'):
            return 'none'
        if lower in ('1', 'odd'):
            return 'odd'
        if lower in ('2', 'even'):
            return 'even'
        # Letter shorthands
        if lower in ('n',):
            return 'none'
        if lower in ('e',):
            return 'even'
        if lower in ('o',):
            return 'odd'
        # Default: return original trimmed value
        return s

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
            # For COM port, we'll look for comm_port setting
            
            query = (
                "SELECT u.unit_id, u.unit_name, u.driver_type, u.driver_addr, u.regulator_type, "
                "COALESCE(com_port_setting.value, '') as com_port, "
                "COALESCE(REPLACE(REPLACE(ip_setting.value, CHAR(13), ''), CHAR(10), ''), '') as ip_address, "
                "COALESCE(mb_mode_setting.value, '0') as mb_mode, "
                "COALESCE(REPLACE(REPLACE(mb_tcp_servers_setting.value, CHAR(13), ''), CHAR(10), ''), '') as mb_tcp_servers, "
                "COALESCE(baudrate_setting.value, '') as baudrate, "
                "COALESCE(parity_setting.value, '') as parity "
                "FROM iw_sys_plant_units u "
                "LEFT JOIN iw_sys_plant_settings com_port_setting ON u.driver_type = com_port_setting.owner AND com_port_setting.setting = 'comm_port' "
                "LEFT JOIN iw_sys_plant_settings ip_setting ON u.driver_type = ip_setting.owner AND "
                "   (ip_setting.value REGEXP '^[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}$' OR "
                "    ip_setting.value REGEXP 'https?://[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}' OR "
                "    ip_setting.value REGEXP '[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}') "
                "LEFT JOIN iw_sys_plant_settings mb_mode_setting ON u.driver_type = mb_mode_setting.owner AND mb_mode_setting.setting = 'mb_mode' "
                "LEFT JOIN iw_sys_plant_settings mb_tcp_servers_setting ON u.driver_type = mb_tcp_servers_setting.owner AND mb_tcp_servers_setting.setting = 'mb_tcp_servers' "
                "LEFT JOIN iw_sys_plant_settings baudrate_setting ON u.driver_type = baudrate_setting.owner AND baudrate_setting.setting = 'comm_baudrate' "
                "LEFT JOIN iw_sys_plant_settings parity_setting ON u.driver_type = parity_setting.owner AND parity_setting.setting = 'comm_parity' "
                "ORDER BY u.unit_id"
            )
            

            
            # Try primary credentials first (iwmac with blank password)
            cmd = [
                self.mysql_exe_path,
                "-h", "127.0.0.1",
                "-u", "iwmac",
                "--password=",
                "-D", "iw_plant_server3",
                "-N", "-B",
                "--protocol=tcp",
                "--default-character-set=utf8mb4",
                "--connect-timeout=5",
                "-e", query,
            ]
            
            def run_query_with_fallback():
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
                        encoding='utf-8',
                        errors='replace',
                        shell=False,
                        startupinfo=startupinfo,
                        creationflags=creationflags,
                    )
                    
                    if result.returncode != 0:
                        err = result.stderr.strip() or "Unknown MySQL error"
                        self.log_queue.put(('info', f"Primary MySQL query failed (iwmac/blank): rc={result.returncode}, stderr='{err}'. Trying fallback credentials..."))
                        try_fallback_credentials()
                        return
                    
                    # Debug: count raw lines
                    raw_output = result.stdout or ''
                    raw_lines = raw_output.splitlines()
                    if self.debug_mysql_logs:
                        self.log_queue.put(('info', f"MySQL returned {len(raw_lines)} raw lines (iwmac/blank). Parsing... (stdout {len(raw_output)} bytes, stderr {len((result.stderr or '').strip())} bytes)"))

                    # Build rows using a map to deduplicate by unit_id and prefer entries with IP
                    rows_map = {}
                    for line in raw_lines:
                        if not line.strip():
                            continue
                        parts = line.split('\t')

                        # Safely unpack with defaults for optional columns
                        unit_id = parts[0] if len(parts) > 0 else ''
                        unit_name = parts[1] if len(parts) > 1 else ''
                        driver_type = parts[2] if len(parts) > 2 else ''
                        driver_addr = parts[3] if len(parts) > 3 else ''
                        regulator_type = parts[4] if len(parts) > 4 else ''
                        com_port_value = parts[5] if len(parts) > 5 else ''
                        full_ip_value = parts[6] if len(parts) > 6 else ''
                        mb_mode = parts[7] if len(parts) > 7 else '0'
                        mb_tcp_servers = parts[8] if len(parts) > 8 else ''
                        baudrate_value = parts[9] if len(parts) > 9 else ''
                        parity_value_raw = parts[10] if len(parts) > 10 else ''
                        parity_value = self.normalize_parity_value(parity_value_raw)

                        clean_ip_address = self.extract_ip_address(full_ip_value)

                        # Handle TCP mode (mb_mode=2): clear COM and, if needed, extract IP from mb_tcp_servers
                        if mb_mode == '2':
                            com_port_value = ''
                            if not clean_ip_address and mb_tcp_servers:
                                tcp_parts = mb_tcp_servers.split(';')
                                if len(tcp_parts) >= 2:
                                    tcp_ip = tcp_parts[1].strip()
                                    if self.extract_ip_address(tcp_ip):
                                        clean_ip_address = tcp_ip

                        # Skip AK2 entries without IP address
                        if driver_type == 'AK2' and not clean_ip_address:
                            continue

                        reordered_row = [
                            unit_id,
                            unit_name,
                            driver_type,
                            driver_addr,
                            regulator_type,
                            com_port_value,
                            baudrate_value,
                            parity_value,
                            clean_ip_address,
                        ]

                        existing = rows_map.get(unit_id)
                        if existing is None:
                            rows_map[unit_id] = reordered_row
                        else:
                            existing_ip = existing[-1]
                            if not existing_ip and clean_ip_address:
                                rows_map[unit_id] = reordered_row

                    rows = list(rows_map.values())
                    if self.debug_mysql_logs:
                        self.log_queue.put(('info', f"Parsed {len(rows)} rows from MySQL (iwmac/blank)."))

                    # If primary returns zero rows, attempt fallback immediately
                    if len(rows) == 0:
                        self.log_queue.put(('info', "No rows from iwmac/blank; trying fallback credentials (root/blank)..."))
                        try_fallback_credentials()
                        return
                    
                    def update_table():
                        # Enable Modbus-supported units filter by default on fetch
                        self.hide_no_baudrate = True
                        # CustomTkinter uses .configure (not .config)
                        self.btn_toggle_baud_filter.configure(text="🔍  Show All Units")
                        self.set_units_rows(rows)
                    self.root.after(0, update_table)
                    self.log_queue.put(('info', f"Successfully loaded {len(rows)} units with COM port and IP address information!"))
                    
                except Exception as e:
                    # If primary credentials fail for any reason, try fallback credentials
                    self.log_queue.put(('info', f"Primary database connection failed (iwmac/blank). Trying fallback credentials..."))
                    try_fallback_credentials()
            
            def try_fallback_credentials():
                """Try connecting with fallback database credentials."""
                try:
                    # Prevent window popup
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    creationflags = subprocess.CREATE_NO_WINDOW

                    # Fallback: root with blank password (validated)
                    fallback_cmd = [
                        self.mysql_exe_path,
                        "-h", "127.0.0.1",
                        "-u", "root",
                        "--password=",
                        "-D", "iw_plant_server3",
                        "-N", "-B",
                        "--protocol=tcp",
                        "--default-character-set=utf8mb4",
                        "--connect-timeout=5",
                        "-e", query,
                    ]
                    
                    self.log_queue.put(('info', "Attempting connection with fallback credentials (root/blank)..."))
                    
                    result = subprocess.run(
                        fallback_cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        encoding='utf-8',
                        errors='replace',
                        shell=False,
                        startupinfo=startupinfo,
                        creationflags=creationflags,
                    )
                    
                    if result.returncode != 0:
                        err = result.stderr.strip() or "Unknown MySQL error"
                        self.log_queue.put(('error', f"Fallback MySQL query failed (root/blank): rc={result.returncode}, stderr='{err}'"))
                        return
                    
                    # Debug: count raw lines
                    raw_output = result.stdout or ''
                    raw_lines = raw_output.splitlines()
                    if self.debug_mysql_logs:
                        self.log_queue.put(('info', f"MySQL returned {len(raw_lines)} raw lines (root/blank). Parsing... (stdout {len(raw_output)} bytes, stderr {len((result.stderr or '').strip())} bytes)"))

                    # Build rows using a map to deduplicate by unit_id and prefer entries with IP
                    rows_map = {}
                    for line in raw_lines:
                        if not line.strip():
                            continue
                        parts = line.split('\t')

                        # Safely unpack with defaults for optional columns
                        unit_id = parts[0] if len(parts) > 0 else ''
                        unit_name = parts[1] if len(parts) > 1 else ''
                        driver_type = parts[2] if len(parts) > 2 else ''
                        driver_addr = parts[3] if len(parts) > 3 else ''
                        regulator_type = parts[4] if len(parts) > 4 else ''
                        com_port_value = parts[5] if len(parts) > 5 else ''
                        full_ip_value = parts[6] if len(parts) > 6 else ''
                        mb_mode = parts[7] if len(parts) > 7 else '0'
                        mb_tcp_servers = parts[8] if len(parts) > 8 else ''
                        baudrate_value = parts[9] if len(parts) > 9 else ''
                        parity_value_raw = parts[10] if len(parts) > 10 else ''
                        parity_value = self.normalize_parity_value(parity_value_raw)

                        clean_ip_address = self.extract_ip_address(full_ip_value)

                        # Handle TCP mode (mb_mode=2)
                        if mb_mode == '2':
                            com_port_value = ''
                            if not clean_ip_address and mb_tcp_servers:
                                tcp_parts = mb_tcp_servers.split(';')
                                if len(tcp_parts) >= 2:
                                    tcp_ip = tcp_parts[1].strip()
                                    if self.extract_ip_address(tcp_ip):
                                        clean_ip_address = tcp_ip

                        # Skip AK2 entries without IP address
                        if driver_type == 'AK2' and not clean_ip_address:
                            continue

                        reordered_row = [
                            unit_id,
                            unit_name,
                            driver_type,
                            driver_addr,
                            regulator_type,
                            com_port_value,
                            baudrate_value,
                            parity_value,
                            clean_ip_address,
                        ]

                        existing = rows_map.get(unit_id)
                        if existing is None:
                            rows_map[unit_id] = reordered_row
                        else:
                            existing_ip = existing[-1]
                            if not existing_ip and clean_ip_address:
                                rows_map[unit_id] = reordered_row

                    rows = list(rows_map.values())
                    if self.debug_mysql_logs:
                        self.log_queue.put(('info', f"Parsed {len(rows)} rows from MySQL (root/blank)."))
                    
                    def update_table():
                        # Enable Modbus-supported units filter by default on fetch
                        self.hide_no_baudrate = True
                        # CustomTkinter uses .configure (not .config)
                        self.btn_toggle_baud_filter.configure(text="🔍  Show All Units")
                        self.set_units_rows(rows)
                    self.root.after(0, update_table)
                    self.log_queue.put(('info', f"Successfully loaded {len(rows)} units with COM port and IP address information!"))
                    
                except Exception as e:
                    self.log_queue.put(('error', f"Fallback database connection also failed: {e}"))
                    messagebox.showerror("Database Error", f"Both primary and fallback database connections failed:\n{str(e)}")
            
            threading.Thread(target=run_query_with_fallback, daemon=True).start()
            
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

        # If we're already blinking, don't cancel/restart (that causes "stuck" indicator under fast events).
        # Instead: show the main color immediately and extend the blink window.
        if getattr(self, "blinking", False) and getattr(self, "blink_job", None):
            try:
                self.status_canvas.itemconfig(self.circle, fill=self.current_color, outline=self.current_color)
                self.status_canvas.itemconfig(self.glow_circle, outline=self.current_color)
            except Exception:
                pass
            # Next tick will dim, and we extend the window by resetting the counter.
            self.blink_state = True
            self.blink_count = 0
            return

        # Start blinking (fresh)
        self.blink_state = False
        self.blinking = True
        self.blink_count = 0
        self.blink_indicator()

    def blink_indicator(self):
        if not self.blinking:
            return

        if self.blink_state:
            # Set to dim color with glow
            self.status_canvas.itemconfig(self.circle, fill=self.dim_color, outline=self.dim_color)
            self.status_canvas.itemconfig(self.glow_circle, outline=self.dim_color)
        else:
            # Set to main color with glow
            self.status_canvas.itemconfig(self.circle, fill=self.current_color, outline=self.current_color)
            self.status_canvas.itemconfig(self.glow_circle, outline=self.current_color)

        self.blink_state = not self.blink_state
        self.blink_count += 1

        # Schedule the next blink
        self.blink_job = self.root.after(400, self.blink_indicator)  # Faster blink for modern feel

        # Stop blinking after 12 toggles (about 5 seconds)
        if self.blink_count >= 12:
            self.blinking = False
            self.blink_count = 0
            self.status_canvas.itemconfig(self.circle, fill="#3A3A4E", outline="#5A5A6E")
            self.status_canvas.itemconfig(self.glow_circle, outline="#3A3A4E")
            if self.blink_job:
                self.root.after_cancel(self.blink_job)
                self.blink_job = None

    # ---------------- End of Status Indicator Methods ----------------





if __name__ == "__main__":
    root = ctk.CTk()
    tool = ModpollingTool(root)
    root.protocol("WM_DELETE_WINDOW", tool.on_closing)
    root.mainloop()
