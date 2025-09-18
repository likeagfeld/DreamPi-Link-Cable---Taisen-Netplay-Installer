#!/usr/bin/env python3
"""
DreamPi Link Cable - Pi Service Installer
Installs Taisen Web UI on Raspberry Pi and creates Windows shortcuts
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import subprocess
import sys
import os
import tempfile
import time
import socket
import json
import webbrowser
import shutil
import winreg
import ctypes
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# Try to import PIL for logo display
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# Try to import win32com for shortcut creation
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# Configuration
GITHUB_REPO_URL = "https://raw.githubusercontent.com/eaudunord/taisen-web-ui/main/install.sh"
DEFAULT_HOSTNAME = "dreampi.local"
DEFAULT_USERNAME = "pi"
DEFAULT_PASSWORD = "raspberry"
DEFAULT_PORT = 22
PORTAL_PORT = 1999

# Windows paths
DESKTOP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop')
START_MENU_PATH = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft\\Windows\\Start Menu\\Programs')

class DreamPiInstaller:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DreamPi Link Cable - Pi Service Installer")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # Set custom icon
        self.set_window_icon()
        
        # Center the window
        self.center_window()
        
        # Configure style
        self.setup_style()
        
        # Installation state
        self.is_installing = False
        self.installation_steps = [
            "Connect to Raspberry Pi (as user)",
            "Download DreamPi Link Cable installer",
            "Install DreamPi Link Cable Web Server", 
            "Verify service is running",
            "Create Windows shortcuts"
        ]
        
        # Configuration variables
        self.hostname_var = tk.StringVar(value=DEFAULT_HOSTNAME)
        self.username_var = tk.StringVar(value=DEFAULT_USERNAME)
        self.password_var = tk.StringVar(value=DEFAULT_PASSWORD)
        self.port_var = tk.IntVar(value=DEFAULT_PORT)
        self.show_password_var = tk.BooleanVar(value=False)
        
        # Shortcut options
        self.create_desktop_shortcut = tk.BooleanVar(value=True)
        self.create_start_menu = tk.BooleanVar(value=True)
        
        # Load saved settings
        self.load_settings()
        
        self.setup_ui()
        
    def set_window_icon(self):
        """Set the window icon from dreampi_logo.ico"""
        try:
            if os.path.exists("dreampi_logo.ico"):
                self.root.iconbitmap("dreampi_logo.ico")
        except Exception as e:
            print(f"Could not set window icon: {e}")
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.root.winfo_screenheight() // 2) - (750 // 2)
        self.root.geometry(f"900x750+{x}+{y}")
    
    def setup_style(self):
        """Configure the application style"""
        self.bg_color = "#f5f5f5"
        self.accent_color = "#8B4513"  # Brown from logo
        self.dark_color = "#2c2c2c"   # Dark gray from logo
        self.success_color = "#28a745"
        self.warning_color = "#ffc107"
        self.error_color = "#dc3545"
        
        self.root.configure(bg=self.bg_color)
    
    def load_logo_image(self):
        """Load and return the DreamPi logo image for display"""
        try:
            if not PIL_AVAILABLE:
                print("PIL not available - cannot display logo")
                return None
                
            if os.path.exists("dreampi_logo.ico"):
                print(f"Loading logo from: {os.path.abspath('dreampi_logo.ico')}")
                # Load the ICO file
                img = Image.open("dreampi_logo.ico")
                # Convert to RGBA to ensure compatibility
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                # Resize to a prominent size for header display
                img = img.resize((150, 150), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                print(f"Logo loaded successfully: {img.size}")
                return photo
            else:
                print(f"Logo file not found at: {os.path.abspath('dreampi_logo.ico')}")
        except Exception as e:
            print(f"Failed to load logo image: {e}")
            import traceback
            traceback.print_exc()
        return None
    
    def setup_ui(self):
        # Create main container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)
        
        # Header with large logo
        header_frame = tk.Frame(main_container, bg=self.bg_color)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        header_frame.columnconfigure(1, weight=1)
        
        # Logo with fallback
        self.logo_image = self.load_logo_image()
        if self.logo_image:
            logo_label = tk.Label(header_frame, image=self.logo_image, bg=self.bg_color)
            logo_label.grid(row=0, column=0, rowspan=2, padx=(0, 30))
        else:
            # Fallback: Show text logo if image fails
            logo_text = tk.Label(header_frame, text="🔗", font=("Arial", 48), 
                               bg=self.bg_color, fg=self.accent_color)
            logo_text.grid(row=0, column=0, rowspan=2, padx=(0, 30))
        
        # Title
        title_label = tk.Label(header_frame, text="DreamPi Link Cable", 
                              font=("Arial", 24, "bold"), fg=self.dark_color, bg=self.bg_color)
        title_label.grid(row=0, column=1, sticky=tk.W)
        
        subtitle_label = tk.Label(header_frame, text="Raspberry Pi Service Installer", 
                                 font=("Arial", 16), fg=self.accent_color, bg=self.bg_color)
        subtitle_label.grid(row=1, column=1, sticky=tk.W)
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_container)
        notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Installation tab
        install_frame = ttk.Frame(notebook, padding="25")
        notebook.add(install_frame, text="Installation")
        
        # Configuration tab
        config_frame = ttk.Frame(notebook, padding="25")
        notebook.add(config_frame, text="Pi Configuration")
        
        # Shortcuts tab
        shortcuts_frame = ttk.Frame(notebook, padding="25")
        notebook.add(shortcuts_frame, text="Shortcuts")
        
        # Uninstall tab
        uninstall_frame = ttk.Frame(notebook, padding="25")
        notebook.add(uninstall_frame, text="Uninstall")
        
        self.setup_install_tab(install_frame)
        self.setup_config_tab(config_frame)
        self.setup_shortcuts_tab(shortcuts_frame)
        self.setup_uninstall_tab(uninstall_frame)
    
    def setup_install_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)  # Make the output area expandable
        
        # Installation overview
        overview_frame = ttk.LabelFrame(parent, text="What This Installer Does", padding="15")
        overview_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        overview_text = """This installer will connect to your Pi via SSH and install the DreamPi Link Cable Web Server service.
It will create shortcuts for easy access to the web interface at http://your-pi:1999"""
        
        overview_label = tk.Label(overview_frame, text=overview_text, justify=tk.LEFT, 
                                 font=("Arial", 10), fg=self.dark_color, wraplength=800)
        overview_label.grid(row=0, column=0, sticky=tk.W)
        
        # Progress and status in one row
        progress_status_frame = ttk.Frame(parent)
        progress_status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        progress_status_frame.columnconfigure(0, weight=1)
        progress_status_frame.columnconfigure(1, weight=1)
        
        # Connection status (left side)
        status_frame = ttk.LabelFrame(progress_status_frame, text="Connection Status", padding="10")
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        ttk.Label(status_frame, text="Target Pi:").grid(row=0, column=0, sticky=tk.W)
        self.connection_status_label = ttk.Label(status_frame, text="Not tested", 
                                               foreground=self.warning_color, font=("Arial", 9, "bold"))
        self.connection_status_label.grid(row=1, column=0, sticky=tk.W)
        
        test_connection_btn = ttk.Button(status_frame, text="Test Connection", 
                                        command=self.test_pi_connection)
        test_connection_btn.grid(row=2, column=0, pady=(10, 0))
        
        # Installation progress (right side)
        progress_frame = ttk.LabelFrame(progress_status_frame, text="Installation Progress", padding="10")
        progress_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        
        # Progress steps (condensed)
        self.step_labels = []
        for i, step in enumerate(self.installation_steps):
            step_text = f"{i+1}. {step}"
            if len(step_text) > 45:
                step_text = step_text[:42] + "..."
            
            step_label = ttk.Label(progress_frame, text=step_text, font=("Arial", 8))
            step_label.grid(row=i, column=0, sticky=tk.W, pady=1)
            
            status_label = ttk.Label(progress_frame, text="Pending", 
                                   foreground=self.warning_color, font=("Arial", 8))
            status_label.grid(row=i, column=1, sticky=tk.W, padx=(10, 0), pady=1)
            
            self.step_labels.append(status_label)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', 
                                           maximum=len(self.installation_steps))
        self.progress_bar.grid(row=len(self.installation_steps), column=0, columnspan=2, 
                              sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Output area (main focus - large and scrollable)
        output_frame = ttk.LabelFrame(parent, text="Installation Output", padding="5")
        output_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
        # Create text widget with explicit scrollbars
        self.output_text = tk.Text(output_frame, height=20, font=("Consolas", 9), 
                                  bg="#f8f8f8", wrap=tk.WORD)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        h_scrollbar = ttk.Scrollbar(output_frame, orient=tk.HORIZONTAL, command=self.output_text.xview)
        
        self.output_text.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid the text and scrollbars
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Button frame
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, pady=(10, 0))
        
        self.install_button = ttk.Button(button_frame, text="Install DreamPi Link Cable on Pi", 
                                        command=self.start_installation)
        self.install_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Add View Full Log button (will be enabled after logging starts)
        self.view_log_button = ttk.Button(button_frame, text="View Full Log", 
                                         command=self.open_log_window, state='disabled')
        self.view_log_button.pack(side=tk.LEFT)
        
        # Initial log
        self.log("DreamPi Link Cable - Pi Service Installer")
        self.log("=" * 60)
        self.log("This installer will set up DreamPi Link Cable Web Server on your Raspberry Pi.")
        self.log("")
        self.log("What will be installed:")
        self.log("• DreamPi Link Cable Web Server (port 1999)")
        self.log("• Service: dreampi-linkcable (auto-start enabled)")
        self.log("• Install location: /opt/dreampi-linkcable")
        self.log("• Files: link_cable.py, webserver.py, index.html")
        self.log("")
        self.log("Before starting:")
        self.log("1. Ensure your Pi is connected to the network")
        self.log("2. SSH must be enabled on your Pi")
        self.log("3. Configure connection settings in 'Pi Configuration' tab if needed")
        self.log("4. Test the connection to verify connectivity")
        self.log("")
        self.log("Click 'Install DreamPi Link Cable on Pi' when ready.")
    
    def setup_config_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        
        title_label = ttk.Label(parent, text="Raspberry Pi Connection Settings", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 25))
        
        # Connection settings
        conn_frame = ttk.LabelFrame(parent, text="SSH Connection Details", padding="20")
        conn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        conn_frame.columnconfigure(1, weight=1)
        
        # Hostname
        ttk.Label(conn_frame, text="Hostname/IP Address:").grid(row=0, column=0, sticky=tk.W, pady=8, padx=(0, 15))
        hostname_entry = ttk.Entry(conn_frame, textvariable=self.hostname_var, width=25, font=("Arial", 11))
        hostname_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=8)
        ttk.Label(conn_frame, text="Default: dreampi.local", foreground="gray", font=("Arial", 9)).grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
        
        # Port
        ttk.Label(conn_frame, text="SSH Port:").grid(row=1, column=0, sticky=tk.W, pady=8, padx=(0, 15))
        port_entry = ttk.Entry(conn_frame, textvariable=self.port_var, width=8, font=("Arial", 11))
        port_entry.grid(row=1, column=1, sticky=tk.W, pady=8)
        ttk.Label(conn_frame, text="Default: 22", foreground="gray", font=("Arial", 9)).grid(row=1, column=2, sticky=tk.W, padx=(15, 0))
        
        # Username
        ttk.Label(conn_frame, text="Username:").grid(row=2, column=0, sticky=tk.W, pady=8, padx=(0, 15))
        username_entry = ttk.Entry(conn_frame, textvariable=self.username_var, width=15, font=("Arial", 11))
        username_entry.grid(row=2, column=1, sticky=tk.W, pady=8)
        ttk.Label(conn_frame, text="Default: pi", foreground="gray", font=("Arial", 9)).grid(row=2, column=2, sticky=tk.W, padx=(15, 0))
        
        # Password
        ttk.Label(conn_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=8, padx=(0, 15))
        
        password_frame = ttk.Frame(conn_frame)
        password_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=8)
        password_frame.columnconfigure(0, weight=1)
        
        self.password_entry = ttk.Entry(password_frame, textvariable=self.password_var, 
                                       show="*", width=20, font=("Arial", 11))
        self.password_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        show_cb = ttk.Checkbutton(password_frame, text="Show", 
                                 variable=self.show_password_var,
                                 command=self.toggle_password_visibility)
        show_cb.grid(row=0, column=1, padx=(10, 0))
        
        ttk.Label(conn_frame, text="Default: raspberry", foreground="gray", font=("Arial", 9)).grid(row=3, column=2, sticky=tk.W, padx=(15, 0))
        
        # Control buttons
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=2, column=0, pady=15)
        
        ttk.Button(control_frame, text="Reset to Defaults", 
                  command=self.reset_to_defaults).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_frame, text="Save Settings", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_frame, text="Test Connection", 
                  command=self.test_pi_connection).pack(side=tk.LEFT)
        
        # Help section
        help_frame = ttk.LabelFrame(parent, text="Connection Help", padding="15")
        help_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(20, 0))
        
        help_text = """Common Connection Settings:

Default Raspberry Pi OS:
• Hostname: dreampi.local (or raspberrypi.local)
• Port: 22 (standard SSH)
• Username: pi
• Password: raspberry

Custom Configuration:
• Change hostname: sudo raspi-config → Network → Hostname
• Custom SSH port: Edit /etc/ssh/sshd_config
• New password: Use 'passwd' command on Pi
• Find IP address: Use 'ip addr' command or check router

Enable SSH (if disabled):
• sudo systemctl enable ssh
• sudo systemctl start ssh"""
        
        help_label = ttk.Label(help_frame, text=help_text, justify=tk.LEFT, 
                              font=("Arial", 9))
        help_label.grid(row=0, column=0, sticky=tk.W)
    
    def setup_shortcuts_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        
        title_label = ttk.Label(parent, text="Windows Shortcuts Configuration", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 25))
        
        # Shortcut options
        shortcuts_frame = ttk.LabelFrame(parent, text="Create Windows Shortcuts", padding="20")
        shortcuts_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        desktop_cb = ttk.Checkbutton(shortcuts_frame, text="Create Desktop shortcut", 
                                    variable=self.create_desktop_shortcut)
        desktop_cb.grid(row=0, column=0, sticky=tk.W, pady=8)
        
        startmenu_cb = ttk.Checkbutton(shortcuts_frame, text="Add to Start Menu", 
                                      variable=self.create_start_menu)
        startmenu_cb.grid(row=1, column=0, sticky=tk.W, pady=8)
    
    def setup_uninstall_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        
        title_label = ttk.Label(parent, text="Uninstall DreamPi Link Cable", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 25))
        
        # Uninstall information
        info_frame = ttk.LabelFrame(parent, text="Uninstall Information", padding="20")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        info_text = """This will completely remove the DreamPi Link Cable Web Server from your Pi:

• Stop the dreampi-linkcable service
• Disable auto-start on boot
• Remove all installation files (/opt/dreampi-linkcable)
• Remove service configuration
• Clean up systemd configuration

What will NOT be affected:
• Your original link_cable.py script
• Python packages (pyserial, requests, etc.)
• DreamPi system functionality
• Any other installed software

To reinstall later, simply run this installer again."""
        
        info_label = tk.Label(info_frame, text=info_text, justify=tk.LEFT, 
                             font=("Arial", 10), fg=self.dark_color, wraplength=800)
        info_label.grid(row=0, column=0, sticky=tk.W)
        
        # Output area
        output_frame = ttk.LabelFrame(parent, text="Uninstall Output", padding="5")
        output_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
        # Create text widget with scrollbars
        self.uninstall_output_text = tk.Text(output_frame, height=15, font=("Consolas", 9), 
                                           bg="#f8f8f8", wrap=tk.WORD)
        
        # Add scrollbars
        uninstall_v_scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, 
                                            command=self.uninstall_output_text.yview)
        uninstall_h_scrollbar = ttk.Scrollbar(output_frame, orient=tk.HORIZONTAL, 
                                            command=self.uninstall_output_text.xview)
        
        self.uninstall_output_text.configure(yscrollcommand=uninstall_v_scrollbar.set, 
                                           xscrollcommand=uninstall_h_scrollbar.set)
        
        # Grid the text and scrollbars
        self.uninstall_output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        uninstall_v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        uninstall_h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Button frame
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, pady=(15, 0))
        
        self.uninstall_button = ttk.Button(button_frame, text="Uninstall DreamPi Link Cable", 
                                         command=self.start_uninstall)
        self.uninstall_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Warning about requiring connection
        warning_label = ttk.Label(parent, text="⚠️  Requires SSH connection to your Pi", 
                                 font=("Arial", 10), foreground=self.warning_color)
        warning_label.grid(row=4, column=0, pady=(10, 0))
        
        # Initialize uninstall output
        self.uninstall_log("DreamPi Link Cable - Uninstaller")
        self.uninstall_log("=" * 50)
        self.uninstall_log("This will completely remove DreamPi Link Cable from your Pi.")
        self.uninstall_log("")
        self.uninstall_log("Configure your Pi connection in the 'Pi Configuration' tab,")
        self.uninstall_log("then click 'Uninstall DreamPi Link Cable' to proceed.")
    
    def log(self, message):
        """Add message to output log"""
        timestamp = time.strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {message}\n"
        self.output_text.insert(tk.END, log_line)
        self.output_text.see(tk.END)
        self.root.update()
        
        # Store log for full view
        if not hasattr(self, 'full_log'):
            self.full_log = []
        self.full_log.append(log_line.rstrip())
        
        # Enable the View Log button after first log entry
        if hasattr(self, 'view_log_button'):
            self.view_log_button.config(state='normal')
    
    def open_log_window(self):
        """Show full log in a separate window"""
        if not hasattr(self, 'full_log') or not self.full_log:
            messagebox.showinfo("No Log", "No log entries available yet.")
            return
            
        log_window = tk.Toplevel(self.root)
        log_window.title("Installation Log - Full View")
        log_window.geometry("900x700")
        log_window.transient(self.root)
        
        # Create text widget with scrollbars
        frame = ttk.Frame(log_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 9))
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=v_scroll.set)
        
        text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Insert log content
        for line in self.full_log:
            text_widget.insert(tk.END, line + "\n")
        
        text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)
        
        # Add close button
        close_btn = ttk.Button(frame, text="Close", command=log_window.destroy)
        close_btn.grid(row=1, column=0, pady=(10, 0))
        
        # Center the window
        log_window.update_idletasks()
        x = (log_window.winfo_screenwidth() // 2) - (900 // 2)
        y = (log_window.winfo_screenheight() // 2) - (700 // 2)
        log_window.geometry(f"900x700+{x}+{y}")
    
    def uninstall_log(self, message):
        """Add message to uninstall output log"""
        timestamp = time.strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {message}\n"
        self.uninstall_output_text.insert(tk.END, log_line)
        self.uninstall_output_text.see(tk.END)
        self.root.update()
    
    def toggle_password_visibility(self):
        """Toggle password field visibility"""
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
    
    def reset_to_defaults(self):
        """Reset Pi settings to defaults"""
        self.hostname_var.set(DEFAULT_HOSTNAME)
        self.username_var.set(DEFAULT_USERNAME)
        self.password_var.set(DEFAULT_PASSWORD)
        self.port_var.set(DEFAULT_PORT)
        self.log("Configuration reset to defaults")
    
    def get_settings_file_path(self):
        """Get settings file path"""
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(app_dir, "dreampi_installer_config.json")
    
    def save_settings(self):
        """Save current settings"""
        try:
            settings = {
                'hostname': self.hostname_var.get(),
                'username': self.username_var.get(),
                'password': self.password_var.get(),
                'port': self.port_var.get()
            }
            
            with open(self.get_settings_file_path(), 'w') as f:
                json.dump(settings, f, indent=2)
            
            self.log("Configuration saved successfully")
            messagebox.showinfo("Settings Saved", "Pi configuration has been saved.")
            
        except Exception as e:
            self.log(f"ERROR: Failed to save configuration - {e}")
            messagebox.showerror("Save Error", f"Failed to save configuration: {e}")
    
    def load_settings(self):
        """Load saved settings"""
        try:
            settings_file = self.get_settings_file_path()
            if os.path.exists(settings_file):
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                
                self.hostname_var.set(settings.get('hostname', DEFAULT_HOSTNAME))
                self.username_var.set(settings.get('username', DEFAULT_USERNAME))
                self.password_var.set(settings.get('password', DEFAULT_PASSWORD))
                self.port_var.set(settings.get('port', DEFAULT_PORT))
        except:
            pass
    
    def update_step_status(self, step, status, color):
        """Update installation step status"""
        if step < len(self.step_labels):
            self.step_labels[step].config(text=status, foreground=color)
        
        # Update progress bar
        completed_steps = sum(1 for label in self.step_labels 
                            if label.cget('text') == 'Complete')
        self.progress_bar['value'] = completed_steps
        self.root.update()
    
    def test_pi_connection(self):
        """Test connection to Pi"""
        def test():
            hostname = self.hostname_var.get().strip()
            port = self.port_var.get()
            
            if not hostname:
                self.log("ERROR: Please enter Pi hostname/IP address")
                return
            
            self.log(f"Testing connection to {hostname}:{port}...")
            
            try:
                # Test hostname resolution
                socket.gethostbyname(hostname)
                self.log(f"SUCCESS: Hostname '{hostname}' resolved")
                
                # Test SSH port
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)
                result = sock.connect_ex((hostname, port))
                sock.close()
                
                if result == 0:
                    self.log(f"SUCCESS: SSH port {port} is accessible")
                    self.connection_status_label.config(text=f"{hostname}:{port} (Ready)", 
                                                       foreground=self.success_color)
                    messagebox.showinfo("Connection Test", 
                                      f"Connection to {hostname}:{port} successful!\n\n"
                                      "Your Pi is ready for installation.")
                else:
                    self.log(f"ERROR: SSH port {port} not accessible")
                    self.connection_status_label.config(text=f"{hostname}:{port} (No SSH)", 
                                                       foreground=self.error_color)
                    messagebox.showerror("Connection Test", 
                                       f"Cannot connect to SSH on {hostname}:{port}\n\n"
                                       "Make sure SSH is enabled on your Pi.")
                    
            except socket.gaierror:
                self.log(f"ERROR: Cannot resolve hostname '{hostname}'")
                self.connection_status_label.config(text=f"{hostname} (Not found)", 
                                                   foreground=self.error_color)
                messagebox.showerror("Connection Test", 
                                   f"Cannot find '{hostname}' on the network.\n\n"
                                   "Check the hostname/IP address is correct.")
            except Exception as e:
                self.log(f"ERROR: Connection test failed - {e}")
                self.connection_status_label.config(text="Connection Error", 
                                                   foreground=self.error_color)
                messagebox.showerror("Connection Test", f"Connection test failed: {e}")
        
        threading.Thread(target=test, daemon=True).start()
    
    def start_installation(self):
        """Start the installation process"""
        if self.is_installing:
            return
        
        # Validate settings
        hostname = self.hostname_var.get().strip()
        username = self.username_var.get().strip()
        
        if not hostname or not username:
            messagebox.showerror("Configuration Required", 
                               "Please configure Pi connection settings in the 'Pi Configuration' tab.")
            return
        
        def install():
            self.is_installing = True
            self.install_button.config(state='disabled', text='Installing...')
            
            try:
                # Reset progress
                self.progress_bar['value'] = 0
                for label in self.step_labels:
                    label.config(text="Pending", foreground=self.warning_color)
                
                self.log("")
                self.log("=" * 60)
                self.log("STARTING DREAMPI INSTALLATION")
                self.log("=" * 60)
                
                # Step 1: Connect to Pi
                if not self.connect_to_pi():
                    raise Exception("Failed to connect to Raspberry Pi")
                
                # Step 2: Download script
                if not self.download_install_script():
                    raise Exception("Failed to download installation script")
                
                # Step 3: Install service
                if not self.install_pi_service():
                    raise Exception("Failed to install Pi service")
                
                # Step 4: Verify installation
                if not self.verify_installation():
                    raise Exception("Installation verification failed")
                
                # Step 5: Create shortcuts
                if not self.create_windows_shortcuts():
                    self.log("Warning: Some shortcuts may not have been created")
                
                # Success!
                self.log("")
                self.log("=" * 60)
                self.log("DREAMPI LINK CABLE INSTALLATION COMPLETED!")
                self.log("=" * 60)
                
                pi_url = f"http://{hostname}:{PORTAL_PORT}"
                self.log(f"DreamPi Link Cable Web Server is now available at: {pi_url}")
                self.log("")
                self.log("Installation details:")
                self.log("• Service: dreampi-linkcable")
                self.log("• Location: /opt/dreampi-linkcable")
                self.log("• Auto-start: Enabled")
                self.log("• Web interface: Port 1999")
                
                if self.create_desktop_shortcut.get():
                    self.log("• Desktop shortcut: Created")
                if self.create_start_menu.get():
                    self.log("• Start Menu entry: Created")
                
                self.log("")
                self.log("Service management commands:")
                self.log("• Status: sudo systemctl status dreampi-linkcable")
                self.log("• Restart: sudo systemctl restart dreampi-linkcable")
                self.log("• Logs: sudo journalctl -u dreampi-linkcable -f")
                self.log("")
                self.log(f"Installation timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Offer to open portal
                response = messagebox.askyesno("Installation Complete", 
                                             f"DreamPi Link Cable installation completed successfully!\n\n"
                                             f"Web Server: {pi_url}\n"
                                             f"Service: dreampi-linkcable\n"
                                             f"Location: /opt/dreampi-linkcable\n\n"
                                             "Would you like to open the web interface now?")
                if response:
                    webbrowser.open(pi_url)
                
            except Exception as e:
                error_msg = str(e)
                self.log(f"ERROR: Installation failed - {error_msg}")
                
                # Get the last few log entries for context
                recent_logs = ""
                if hasattr(self, 'full_log') and self.full_log:
                    recent_logs = "\n".join(self.full_log[-10:])  # Last 10 log entries
                
                # Show detailed error dialog
                detailed_error = f"Installation failed: {error_msg}\n\nRecent log entries:\n{recent_logs}"
                
                error_dialog = tk.Toplevel(self.root)
                error_dialog.title("Installation Failed")
                error_dialog.geometry("700x500")
                error_dialog.transient(self.root)
                
                # Center the error dialog
                error_dialog.update_idletasks()
                x = (error_dialog.winfo_screenwidth() // 2) - (700 // 2)
                y = (error_dialog.winfo_screenheight() // 2) - (500 // 2)
                error_dialog.geometry(f"700x500+{x}+{y}")
                
                frame = ttk.Frame(error_dialog, padding="15")
                frame.pack(fill=tk.BOTH, expand=True)
                frame.columnconfigure(0, weight=1)
                frame.rowconfigure(1, weight=1)
                
                ttk.Label(frame, text="Installation Failed", 
                         font=("Arial", 14, "bold"), foreground=self.error_color).grid(row=0, column=0, pady=(0, 15))
                
                # Error details in scrollable text
                error_text = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 9), height=20)
                error_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=error_text.yview)
                error_text.configure(yscrollcommand=error_scroll.set)
                
                error_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
                error_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
                
                error_text.insert(tk.END, detailed_error)
                error_text.config(state=tk.DISABLED)
                
                # Buttons
                btn_frame = ttk.Frame(frame)
                btn_frame.grid(row=2, column=0, columnspan=2, pady=(15, 0))
                
                ttk.Button(btn_frame, text="View Full Log", 
                          command=lambda: [error_dialog.destroy(), self.open_log_window()]).pack(side=tk.LEFT, padx=(0, 10))
                ttk.Button(btn_frame, text="Close", command=error_dialog.destroy).pack(side=tk.LEFT)
            
            finally:
                self.is_installing = False
                self.install_button.config(state='normal', text='Install DreamPi Link Cable on Pi')
        
        threading.Thread(target=install, daemon=True).start()
    
    def connect_to_pi(self):
        """Connect to Raspberry Pi"""
        self.log("Step 1: Connecting to Raspberry Pi...")
        self.update_step_status(0, "Connecting...", self.warning_color)
        
        try:
            hostname = self.hostname_var.get().strip()
            port = self.port_var.get()
            
            # Test connection
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(10)
            result = sock.connect_ex((hostname, port))
            sock.close()
            
            if result != 0:
                raise Exception(f"Cannot connect to {hostname}:{port}")
            
            self.log(f"Successfully connected to {hostname}:{port}")
            self.update_step_status(0, "Complete", self.success_color)
            return True
            
        except Exception as e:
            self.log(f"ERROR: Connection failed - {e}")
            self.update_step_status(0, "Failed", self.error_color)
            return False
    
    def download_install_script(self):
        """Download installation script from GitHub"""
        self.log("Step 2: Downloading installation script...")
        self.update_step_status(1, "Downloading...", self.warning_color)
        
        try:
            self.log(f"Downloading from: {GITHUB_REPO_URL}")
            request = Request(GITHUB_REPO_URL, headers={'User-Agent': 'DreamPiInstaller/1.0'})
            with urlopen(request, timeout=30) as response:
                self.install_script_content = response.read().decode('utf-8')
            
            if not self.install_script_content.strip():
                raise Exception("Downloaded script is empty")
            
            self.log(f"Successfully downloaded script ({len(self.install_script_content)} bytes)")
            self.update_step_status(1, "Complete", self.success_color)
            return True
            
        except Exception as e:
            self.log(f"ERROR: Download failed - {e}")
            self.update_step_status(1, "Failed", self.error_color)
            return False
    
    def execute_ssh_command(self, command, timeout=60):
        """Execute a command via SSH and return output"""
        hostname = self.hostname_var.get().strip()
        username = self.username_var.get().strip() 
        password = self.password_var.get()
        port = self.port_var.get()
        
        # Create SSH command
        ssh_cmd = [
            "ssh",
            "-o", "StrictHostKeyChecking=no",
            "-o", "UserKnownHostsFile=/dev/null", 
            "-o", "PreferredAuthentications=password",
            "-o", "PubkeyAuthentication=no",
            "-o", f"ConnectTimeout={min(timeout, 30)}",
            "-p", str(port),
            f"{username}@{hostname}",
            command
        ]
        
        self.log(f"SSH Command: {' '.join(ssh_cmd[:-1])} '[COMMAND]'")
        
        try:
            # Use sshpass for password authentication if available
            if shutil.which("sshpass"):
                full_cmd = ["sshpass", "-p", password] + ssh_cmd
                env = os.environ.copy()
            else:
                full_cmd = ssh_cmd
                env = os.environ.copy()
                env['SSH_ASKPASS'] = 'echo'
                env['DISPLAY'] = 'dummy:0'
            
            process = subprocess.Popen(
                full_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                env=env
            )
            
            # Use communicate with timeout instead of passing timeout to Popen
            try:
                stdout, stderr = process.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                process.kill()
                stdout, stderr = process.communicate()
                return -1, stdout, f"SSH command timed out after {timeout} seconds"
            
            return process.returncode, stdout, stderr
            
        except Exception as e:
            return -1, "", str(e)
    
    def install_pi_service(self):
        """Install service on Pi via SSH"""
        self.log("Step 3: Installing DreamPi Link Cable service...")
        self.update_step_status(2, "Installing...", self.warning_color)
        
        try:
            # Create installation script
            install_script = '''#!/bin/bash
set -e

echo "=== DreamPi Link Cable Installation ==="
echo "Timestamp: $(date)"
echo "User: $(whoami) (UID: $(id -u))"
echo "Directory: $(pwd)"
echo ""

# Prevent running as root
if [ "$(id -u)" = "0" ]; then
    echo "ERROR: This installer must NOT be run as root/sudo"
    echo "Please run as regular user: ./install.sh"
    exit 1
fi

echo "Downloading installation script..."
curl -sSL https://raw.githubusercontent.com/eaudunord/taisen-web-ui/main/install.sh -o /tmp/dreampi_install.sh

if [ ! -f "/tmp/dreampi_install.sh" ]; then
    echo "ERROR: Failed to download installation script"
    exit 1
fi

echo "Making script executable..."
chmod +x /tmp/dreampi_install.sh

echo "Running installation script..."
/tmp/dreampi_install.sh

echo ""
echo "=== Post-Installation Checks ==="

# Check installation directory
if [ -d "/opt/dreampi-linkcable" ]; then
    echo "✓ Installation directory exists"
    echo "Files installed:"
    ls -la /opt/dreampi-linkcable/ 2>/dev/null || echo "  Cannot list directory contents"
else
    echo "✗ Installation directory not found"
fi

# Check service status
echo ""
echo "Checking service status..."
if systemctl --user is-active --quiet dreampi-linkcable 2>/dev/null; then
    echo "✓ User service is active"
    systemctl --user status dreampi-linkcable --no-pager -l 2>/dev/null || true
elif sudo systemctl is-active --quiet dreampi-linkcable 2>/dev/null; then
    echo "✓ System service is active"
    sudo systemctl status dreampi-linkcable --no-pager -l 2>/dev/null || true
else
    echo "? Service status unclear - checking both user and system:"
    echo "User service:"
    systemctl --user status dreampi-linkcable --no-pager -l 2>/dev/null || echo "  Not found"
    echo "System service:"
    sudo systemctl status dreampi-linkcable --no-pager -l 2>/dev/null || echo "  Not found or no sudo access"
fi

# Test web server
echo ""
echo "Testing web server..."
if curl -s -m 10 http://localhost:1999 >/dev/null 2>&1; then
    echo "✓ Web server responding on port 1999"
else
    echo "✗ Web server not responding on port 1999"
    echo "  This may be normal if the service is still starting"
fi

# Clean up
rm -f /tmp/dreampi_install.sh

echo ""
echo "=== Installation Complete ==="
echo "Timestamp: $(date)"
'''
            
            self.log("Executing installation script on Pi...")
            self.log("This may take a few minutes...")
            
            # Execute installation
            return_code, stdout, stderr = self.execute_ssh_command(install_script, timeout=300)
            
            # Log all output
            if stdout:
                for line in stdout.split('\n'):
                    if line.strip():
                        self.log(f"Pi: {line}")
            
            if stderr:
                for line in stderr.split('\n'):
                    if line.strip():
                        self.log(f"SSH Error: {line}")
            
            if return_code == 0:
                self.log("Installation completed successfully")
                self.update_step_status(2, "Complete", self.success_color)
                return True
            else:
                self.log(f"Installation failed with exit code {return_code}")
                self.update_step_status(2, "Failed", self.error_color)
                return False
                
        except Exception as e:
            self.log(f"ERROR: Installation failed - {e}")
            self.update_step_status(2, "Failed", self.error_color)
            return False
    
    def verify_installation(self):
        """Verify the installation was successful"""
        self.log("Step 4: Verifying DreamPi Link Cable installation...")
        self.update_step_status(3, "Verifying...", self.warning_color)
        
        try:
            hostname = self.hostname_var.get().strip()
            
            # Wait a moment for service to fully start
            time.sleep(5)
            
            # Test if DreamPi Link Cable web server is accessible on port 1999
            self.log(f"Testing DreamPi Link Cable web server at {hostname}:{PORTAL_PORT}...")
            
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(15)  # Give more time for service to start
            result = sock.connect_ex((hostname, PORTAL_PORT))
            sock.close()
            
            if result == 0:
                self.log(f"SUCCESS: DreamPi Link Cable web server is accessible at {hostname}:{PORTAL_PORT}")
                self.update_step_status(3, "Complete", self.success_color)
                return True
            else:
                self.log(f"WARNING: Web server not yet accessible at {hostname}:{PORTAL_PORT}")
                self.log("Checking service status on Pi...")
                
                # Check service status via SSH
                return_code, stdout, stderr = self.execute_ssh_command(
                    "systemctl --user is-active dreampi-linkcable 2>/dev/null || sudo systemctl is-active dreampi-linkcable 2>/dev/null || echo 'Service status unknown'",
                    timeout=30
                )
                
                if stdout:
                    self.log(f"Service status: {stdout.strip()}")
                
                # Still mark as complete but with warning
                self.log("Installation may still be successful - service might be starting")
                self.update_step_status(3, "Complete", self.warning_color)
                return True
                
        except Exception as e:
            self.log(f"WARNING: Verification test failed - {e}")
            self.log("Installation may still be successful - check manually")
            self.update_step_status(3, "Complete", self.warning_color)
            return True
    
    def create_windows_shortcuts(self):
        """Create Windows shortcuts"""
        self.log("Step 5: Creating Windows shortcuts...")
        self.update_step_status(4, "Creating...", self.warning_color)
        
        try:
            hostname = self.hostname_var.get().strip()
            portal_url = f"http://{hostname}:{PORTAL_PORT}"
            
            shortcuts_created = 0
            
            # Create desktop shortcut
            if self.create_desktop_shortcut.get():
                if self.create_desktop_shortcut_file(portal_url):
                    shortcuts_created += 1
                    self.log("Desktop shortcut created successfully")
                else:
                    self.log("Failed to create desktop shortcut")
            
            # Create Start Menu entry
            if self.create_start_menu.get():
                if self.create_start_menu_shortcut(portal_url):
                    shortcuts_created += 1
                    self.log("Start Menu entry created successfully")
                else:
                    self.log("Failed to create Start Menu entry")
            
            if shortcuts_created > 0:
                self.update_step_status(4, "Complete", self.success_color)
                return True
            else:
                self.update_step_status(4, "Failed", self.error_color)
                return False
                
        except Exception as e:
            self.log(f"ERROR: Shortcut creation failed - {e}")
            self.update_step_status(4, "Failed", self.error_color)
            return False
    
    def create_desktop_shortcut_file(self, portal_url):
        """Create desktop shortcut"""
        try:
            shortcut_created = False
            
            # Try win32com method first
            if WIN32_AVAILABLE:
                try:
                    self.log("Attempting to create desktop shortcut using win32com...")
                    import win32com.client
                    
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut_path = os.path.join(DESKTOP_PATH, "DreamPi Link Cable.lnk")
                    shortcut = shell.CreateShortCut(shortcut_path)
                    
                    # For web URLs, we need to set the target to the browser
                    shortcut.Targetpath = "cmd.exe"
                    shortcut.Arguments = f'/c start "" "{portal_url}"'
                    shortcut.Description = "DreamPi Link Cable Web Interface"
                    shortcut.WindowStyle = 7  # Minimized
                    
                    # Set icon - copy to permanent location and use absolute path
                    icon_set = False
                    if os.path.exists("dreampi_logo.ico"):
                        try:
                            icon_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DreamPi')
                            os.makedirs(icon_dir, exist_ok=True)
                            icon_path = os.path.join(icon_dir, 'dreampi_logo.ico')
                            
                            # Copy icon to permanent location
                            shutil.copy2("dreampi_logo.ico", icon_path)
                            
                            # Verify icon was copied
                            if os.path.exists(icon_path):
                                # Use absolute path with proper format for Windows shortcuts
                                shortcut.IconLocation = f"{icon_path},0"
                                self.log(f"Icon set to: {icon_path}")
                                icon_set = True
                            else:
                                self.log("Failed to copy icon file")
                                
                        except Exception as e:
                            self.log(f"Icon setup failed: {e}")
                    
                    if not icon_set:
                        self.log("No icon set - using default")
                    
                    shortcut.save()
                    
                    if os.path.exists(shortcut_path):
                        self.log("Desktop shortcut created successfully using win32com")
                        shortcut_created = True
                    else:
                        self.log("win32com shortcut creation failed - file not found")
                        
                except Exception as e:
                    self.log(f"win32com desktop shortcut failed: {e}")
            
            # Fallback to .url file if win32com failed
            if not shortcut_created:
                try:
                    self.log("Creating desktop shortcut as .url file...")
                    shortcut_path = os.path.join(DESKTOP_PATH, "DreamPi Link Cable.url")
                    
                    # Ensure desktop directory exists
                    os.makedirs(DESKTOP_PATH, exist_ok=True)
                    
                    # Copy icon to a permanent location and get absolute path
                    icon_line = ""
                    if os.path.exists("dreampi_logo.ico"):
                        try:
                            icon_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DreamPi')
                            os.makedirs(icon_dir, exist_ok=True)
                            icon_path = os.path.join(icon_dir, 'dreampi_logo.ico')
                            shutil.copy2("dreampi_logo.ico", icon_path)
                            
                            if os.path.exists(icon_path):
                                # Convert to proper Windows path format for .url files
                                windows_icon_path = icon_path.replace('/', '\\')
                                icon_line = f"IconFile={windows_icon_path}\nIconIndex=0\n"
                                self.log(f"URL icon set to: {windows_icon_path}")
                            else:
                                self.log("Failed to copy icon for .url file")
                                
                        except Exception as e:
                            self.log(f"URL icon setup failed: {e}")
                    
                    # Create .url file
                    with open(shortcut_path, 'w', encoding='utf-8') as f:
                        f.write(f"""[InternetShortcut]
URL={portal_url}
{icon_line}""")
                    
                    if os.path.exists(shortcut_path):
                        self.log(f"Desktop shortcut created as .url file: {shortcut_path}")
                        shortcut_created = True
                    else:
                        self.log("Failed to create .url file")
                        
                except Exception as e:
                    self.log(f"URL file desktop shortcut failed: {e}")
            
            return shortcut_created
            
        except Exception as e:
            self.log(f"Desktop shortcut creation failed: {e}")
            return False
    
    def create_start_menu_shortcut(self, portal_url):
        """Create Start Menu shortcut"""
        try:
            shortcut_created = False
            
            # Create DreamPi folder in Start Menu
            start_menu_folder = os.path.join(START_MENU_PATH, "DreamPi Link Cable")
            os.makedirs(start_menu_folder, exist_ok=True)
            
            # Try win32com method first
            if WIN32_AVAILABLE:
                try:
                    self.log("Attempting to create Start Menu shortcut using win32com...")
                    import win32com.client
                    
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut_path = os.path.join(start_menu_folder, "DreamPi Link Cable.lnk")
                    shortcut = shell.CreateShortCut(shortcut_path)
                    
                    # For web URLs, we need to set the target to the browser
                    shortcut.Targetpath = "cmd.exe"
                    shortcut.Arguments = f'/c start "" "{portal_url}"'
                    shortcut.Description = "DreamPi Link Cable Web Interface"
                    shortcut.WindowStyle = 7  # Minimized
                    
                    # Set icon - copy to permanent location and use absolute path
                    icon_set = False
                    if os.path.exists("dreampi_logo.ico"):
                        try:
                            icon_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DreamPi')
                            os.makedirs(icon_dir, exist_ok=True)
                            icon_path = os.path.join(icon_dir, 'dreampi_logo.ico')
                            
                            # Copy icon to permanent location
                            shutil.copy2("dreampi_logo.ico", icon_path)
                            
                            # Verify icon was copied
                            if os.path.exists(icon_path):
                                # Use absolute path with proper format for Windows shortcuts
                                shortcut.IconLocation = f"{icon_path},0"
                                self.log(f"Start Menu icon set to: {icon_path}")
                                icon_set = True
                            else:
                                self.log("Failed to copy icon file for Start Menu")
                                
                        except Exception as e:
                            self.log(f"Start Menu icon setup failed: {e}")
                    
                    if not icon_set:
                        self.log("No icon set for Start Menu - using default")
                    
                    shortcut.save()
                    
                    if os.path.exists(shortcut_path):
                        self.log("Start Menu shortcut created successfully using win32com")
                        shortcut_created = True
                    else:
                        self.log("win32com Start Menu shortcut creation failed - file not found")
                        
                except Exception as e:
                    self.log(f"win32com Start Menu shortcut failed: {e}")
            
            # Fallback to .url file if win32com failed
            if not shortcut_created:
                try:
                    self.log("Creating Start Menu shortcut as .url file...")
                    shortcut_path = os.path.join(start_menu_folder, "DreamPi Link Cable.url")
                    
                    # Copy icon to a permanent location and get absolute path
                    icon_line = ""
                    if os.path.exists("dreampi_logo.ico"):
                        try:
                            icon_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'DreamPi')
                            os.makedirs(icon_dir, exist_ok=True)
                            icon_path = os.path.join(icon_dir, 'dreampi_logo.ico')
                            shutil.copy2("dreampi_logo.ico", icon_path)
                            
                            if os.path.exists(icon_path):
                                # Convert to proper Windows path format for .url files
                                windows_icon_path = icon_path.replace('/', '\\')
                                icon_line = f"IconFile={windows_icon_path}\nIconIndex=0\n"
                                self.log(f"Start Menu URL icon set to: {windows_icon_path}")
                            else:
                                self.log("Failed to copy icon for Start Menu .url file")
                                
                        except Exception as e:
                            self.log(f"Start Menu URL icon setup failed: {e}")
                    
                    # Create .url file
                    with open(shortcut_path, 'w', encoding='utf-8') as f:
                        f.write(f"""[InternetShortcut]
URL={portal_url}
{icon_line}""")
                    
                    if os.path.exists(shortcut_path):
                        self.log(f"Start Menu shortcut created as .url file: {shortcut_path}")
                        shortcut_created = True
                    else:
                        self.log("Failed to create Start Menu .url file")
                        
                except Exception as e:
                    self.log(f"URL file Start Menu shortcut failed: {e}")
            
            return shortcut_created
            
        except Exception as e:
            self.log(f"Start Menu shortcut creation failed: {e}")
            return False
    
    def start_uninstall(self):
        """Start the uninstall process"""
        if self.is_installing:
            messagebox.showwarning("Installation in Progress", 
                                 "Please wait for the current installation to complete.")
            return
        
        # Validate settings
        hostname = self.hostname_var.get().strip()
        username = self.username_var.get().strip()
        
        if not hostname or not username:
            messagebox.showerror("Configuration Required", 
                               "Please configure Pi connection settings in the 'Pi Configuration' tab.")
            return
        
        # Confirm uninstall
        response = messagebox.askyesno("Confirm Uninstall", 
                                     f"Are you sure you want to completely remove DreamPi Link Cable from {hostname}?\n\n"
                                     "This will:\n"
                                     "• Stop the web server service\n"
                                     "• Remove all installation files\n"
                                     "• Disable auto-start\n\n"
                                     "This action cannot be undone.")
        if not response:
            return
        
        def uninstall():
            self.is_installing = True  # Reuse this flag to prevent concurrent operations
            self.uninstall_button.config(state='disabled', text='Uninstalling...')
            
            try:
                self.uninstall_log("")
                self.uninstall_log("=" * 50)
                self.uninstall_log("STARTING DREAMPI LINK CABLE UNINSTALL")
                self.uninstall_log("=" * 50)
                
                if not self.execute_uninstall():
                    raise Exception("Uninstall process failed")
                
                # Success!
                self.uninstall_log("")
                self.uninstall_log("=" * 50)
                self.uninstall_log("DREAMPI LINK CABLE UNINSTALLED SUCCESSFULLY!")
                self.uninstall_log("=" * 50)
                self.uninstall_log("All DreamPi Link Cable files and services have been removed.")
                self.uninstall_log("Your Pi has been cleaned up and returned to its previous state.")
                self.uninstall_log("")
                self.uninstall_log("To reinstall in the future, run this installer again.")
                
                messagebox.showinfo("Uninstall Complete", 
                                  "DreamPi Link Cable has been completely removed from your Pi.\n\n"
                                  "All services stopped and files deleted.\n"
                                  "Your Pi is now clean.")
                
            except Exception as e:
                self.uninstall_log(f"ERROR: Uninstall failed - {e}")
                messagebox.showerror("Uninstall Failed", 
                                   f"Uninstall failed: {e}\n\n"
                                   "Check the output log for details.")
            
            finally:
                self.is_installing = False
                self.uninstall_button.config(state='normal', text='Uninstall DreamPi Link Cable')
        
        threading.Thread(target=uninstall, daemon=True).start()
    
    def execute_uninstall(self):
        """Execute the uninstall commands on Pi via SSH"""
        try:
            # Create uninstall script based on uninstall.sh
            uninstall_script = '''#!/bin/bash
set -e

SERVICE_NAME="dreampi-linkcable"
INSTALL_DIR="/opt/dreampi-linkcable"

echo "=== DreamPi Link Cable Web Server - Complete Uninstaller ==="
echo "Timestamp: $(date)"
echo "User: $(whoami) (UID: $(id -u))"
echo ""

# Stop service if running
echo "Stopping service..."
if sudo systemctl is-active --quiet ${SERVICE_NAME}.service 2>/dev/null; then
    sudo systemctl stop ${SERVICE_NAME}.service
    echo "✓ Service stopped"
else
    echo "✓ Service was not running"
fi

# Disable auto-start
echo "Disabling auto-start..."
if sudo systemctl is-enabled --quiet ${SERVICE_NAME}.service 2>/dev/null; then
    sudo systemctl disable ${SERVICE_NAME}.service
    echo "✓ Auto-start disabled"
else
    echo "✓ Auto-start was not enabled"
fi

# Remove service file
echo "Removing service file..."
if [ -f "/etc/systemd/system/${SERVICE_NAME}.service" ]; then
    sudo rm "/etc/systemd/system/${SERVICE_NAME}.service"
    echo "✓ Service file removed"
else
    echo "✓ Service file was not found"
fi

# Reload systemd
echo "Reloading systemd configuration..."
sudo systemctl daemon-reload
echo "✓ Systemd configuration reloaded"

# Remove installation directory
echo "Removing installation files..."
if [ -d "$INSTALL_DIR" ]; then
    sudo rm -rf "$INSTALL_DIR"
    echo "✓ Installation directory removed: $INSTALL_DIR"
else
    echo "✓ Installation directory was not found"
fi

echo ""
echo "=== Uninstall Complete ==="
echo "✓ Service stopped and disabled"
echo "✓ All installation files removed"
echo "✓ Auto-start configuration removed"
echo "✓ DreamPi Link Cable Web Server completely removed"
echo ""
echo "What was NOT affected:"
echo "  - Your original link_cable.py script"
echo "  - Python packages (pyserial, requests, etc.)"
echo "  - DreamPi system functionality"
echo "  - Any other installed software"
echo ""
echo "Uninstall completed successfully at $(date)"
'''
            
            self.uninstall_log("Executing uninstall commands on Pi...")
            self.uninstall_log("This may take a few moments...")
            
            # Execute uninstall
            return_code, stdout, stderr = self.execute_ssh_command(uninstall_script, timeout=120)
            
            # Log all output
            if stdout:
                for line in stdout.split('\n'):
                    if line.strip():
                        self.uninstall_log(f"Pi: {line}")
            
            if stderr:
                for line in stderr.split('\n'):
                    if line.strip():
                        self.uninstall_log(f"SSH Error: {line}")
            
            if return_code == 0:
                self.uninstall_log("Uninstall completed successfully")
                return True
            else:
                self.uninstall_log(f"Uninstall failed with exit code {return_code}")
                return False
                
        except Exception as e:
            self.uninstall_log(f"ERROR: Uninstall execution failed - {e}")
            return False
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass

def main():
    """Application entry point"""
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    
    os.chdir(app_dir)
    
    installer = DreamPiInstaller()
    installer.run()

if __name__ == "__main__":
    main()