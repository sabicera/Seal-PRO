import sys
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import threading
import json
import tkinter as tk
import datetime

# Import core logic and updater
from converter_core import SealCheckConverterCore
from updater import AutoUpdater

# Version
__version__ = "1.0.2"

# GitHub repo info
GITHUB_OWNER = "sabicera"
GITHUB_REPO = "Seal_PRO"

# Set appearance mode and color theme BEFORE creating any widgets
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Modern DARK color scheme
BG_MAIN = "#3E5761"        # Main window background
BG_FRAME = "#51707C"       # Frame backgrounds
BG_INPUT = "#2d2d2d"       # Input fields
BG_CONSOLE = "#1e1e1e"     # Console background
BG_HOVER = "#3d3d3d"       # Hover state
FG_TEXT = "#ffffff"        # Primary text
FG_SECONDARY = "#f6fa00"   # Secondary text
ACCENT_GREEN = "#00d26a"   # Success/Convert button
ACCENT_RED = "#ff4757"     # Clear/Error button
ACCENT_BLUE = "#1e90ff"    # Info/Change button

class CTkScrollableFrame(ctk.CTkScrollableFrame):
    """Custom scrollable frame for filters"""
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)


class SealCheckConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Seal Check Excel Converter")
        self.root.geometry("900x800")
        
        # CRITICAL: Set root background to dark
        self.root.configure(bg=BG_MAIN)
        
        # Set window icon
        try:
            if getattr(sys, 'frozen', False):
                application_path = sys._MEIPASS
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(application_path, "seal.ico")
            self.root.iconbitmap(icon_path)
            
            if sys.platform == 'win32':
                self.root.iconbitmap(default=icon_path)
        except Exception as e:
            print(f"Icon not loaded: {e}")
        
        # Template path
        if getattr(sys, 'frozen', False):
            script_dir = Path(sys._MEIPASS)
        else:
            script_dir = Path(__file__).parent
        self.template_path = script_dir / 'Seal_Check_Template.xlsx'
        
        # Initialize core converter
        self.converter = SealCheckConverterCore(self.template_path)
        
        # Initialize updater
        self.updater = AutoUpdater(__version__, GITHUB_OWNER, GITHUB_REPO)
        
        # Config file path
        self.config_file = script_dir / 'converter_config.json'
        
        # Load saved settings
        self.load_config()
        
        # Store selected source file
        self.selected_file = None
        
        self.setup_ui()
        
        # Register cleanup on close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Check for updates on startup
        threading.Thread(target=self.check_updates_background, daemon=True).start()
        
    def check_updates_background(self):
        """Check for updates in background"""
        try:
            has_update, latest, url, notes = self.updater.check_for_updates()
            if has_update:
                self.root.after(1000, lambda: self.prompt_update(latest, url, notes))
            else:
                self.log("You're up to date (v{})".format(__version__))
        except:
            self.log("You're up to date (v{})".format(__version__))
    
    def prompt_update(self, latest_version, download_url, release_notes):
        """Show update prompt to user"""
        if messagebox.askyesno("Update Available", 
            f"New version {latest_version} available!\n\nDownload and install?"):
            self.download_and_install_update(download_url)
    
    def download_and_install_update(self, download_url):
        """Download and install update"""
        progress_window = ctk.CTkToplevel(self.root)
        progress_window.title("Downloading Update")
        progress_window.geometry("450x180")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        ctk.CTkLabel(
            progress_window, 
            text="Downloading update...",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=20)
        
        progress_bar = ctk.CTkProgressBar(progress_window, width=350)
        progress_bar.pack(pady=15)
        progress_bar.set(0)
        
        status_label = ctk.CTkLabel(
            progress_window,
            text="0%",
            font=ctk.CTkFont(size=12)
        )
        status_label.pack()
        
        def update_progress(percent):
            progress_bar.set(percent / 100)
            status_label.configure(text=f"{percent}%")
        
        def download_thread():
            new_exe = self.updater.download_update(download_url, update_progress)
            if new_exe:
                progress_window.destroy()
                if messagebox.askyesno("Update Ready", "Restart now?"):
                    self.updater.apply_update(new_exe)
            else:
                progress_window.destroy()
                messagebox.showerror("Update Failed", "Failed to download update.")
        
        threading.Thread(target=download_thread, daemon=True).start()
        
    def load_config(self):
        """Load saved configuration"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.output_dir = Path(config.get('output_dir', str(Path.home() / 'Desktop')))
                    self.vessel_voyage_pairs = config.get('vessel_voyage_pairs', [])
            else:
                self.output_dir = Path.home() / 'Desktop'
                self.vessel_voyage_pairs = []
        except Exception as e:
            print(f"Error loading config: {e}")
            self.output_dir = Path.home() / 'Desktop'
            self.vessel_voyage_pairs = []
    
    def save_config(self):
        """Save current configuration"""
        try:
            # Collect all vessel/voyage pairs
            pairs = []
            for vessel_var, voyage_var in self.vessel_voyage_vars:
                vessel = vessel_var.get().strip()
                voyage = voyage_var.get().strip()
                if vessel or voyage:
                    pairs.append({'vessel': vessel, 'voyage': voyage})
            
            config = {
                'output_dir': str(self.output_dir),
                'vessel_voyage_pairs': pairs
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def on_closing(self):
        """Handle window closing"""
        self.save_config()
        self.root.destroy()
    
    def uppercase_vessel_trace(self, var):
        """Convert vessel name to uppercase"""
        def trace_func(*args):
            value = var.get()
            if value != value.upper():
                var.set(value.upper())
        return trace_func
    
    def uppercase_voyage_trace(self, var):
        """Convert voyage number to uppercase"""
        def trace_func(*args):
            value = var.get()
            if value != value.upper():
                var.set(value.upper())
        return trace_func
    
    def generate_output_filename(self, vessel, pol):
        """Generate output filename based on POL value"""
        now = datetime.datetime.now()
        time_str = now.strftime("%H%M")
        number = now.strftime("%S%f")[:4]
        
        if pol == 'ECGYE':
            return f"{vessel}_NAPORTEC_{time_str}_{number}.xlsx"
        elif pol == 'PEPIO':
            return f"{vessel}_PISCO_{time_str}_{number}.xlsx"
        elif pol == 'PSA-RODMAN':
            return f"{vessel}_RODMAN_{number}.xlsx"
        else:
            return f"{vessel}_{number}.xlsx"
        
    def setup_ui(self):
        # Main container with dark background
        main_container = ctk.CTkFrame(self.root, fg_color=BG_MAIN)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # LEFT PANEL
        left_panel = ctk.CTkFrame(main_container, fg_color=BG_MAIN)
        left_panel.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # RIGHT PANEL
        right_panel = ctk.CTkFrame(main_container, fg_color=BG_MAIN)
        right_panel.pack(side='right', fill='both', expand=False)
        
        # ========== LEFT PANEL ==========
        
        # Drop zone frame
        drop_frame = ctk.CTkFrame(
            left_panel, 
            height=100, 
            corner_radius=10, 
            border_width=2, 
            border_color=ACCENT_BLUE,
            fg_color=BG_FRAME
        )
        drop_frame.pack(fill='x', pady=(0, 20))
        
        # Register drop target
        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        
        self.drop_label = ctk.CTkLabel(
            drop_frame,
            text="Drag & Drop or Click to Browse",
            font=ctk.CTkFont(size=15, weight="bold"),
            cursor="hand2"
        )
        self.drop_label.pack(pady=30)
        self.drop_label.bind('<Button-1>', lambda e: self.browse_source())
        drop_frame.bind('<Button-1>', lambda e: self.browse_source())
        
        # Selected File Section
        ctk.CTkLabel(
            left_panel,
            text="Selected File",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor='w'
        ).pack(anchor='w', pady=(0, 8))
        
        file_frame = ctk.CTkFrame(left_panel, corner_radius=8, fg_color=BG_FRAME)
        file_frame.pack(fill='x', pady=(0, 20))
        
        self.file_label = ctk.CTkLabel(
            file_frame,
            text="No file selected",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=FG_SECONDARY,
            anchor='w'
        )
        self.file_label.pack(fill='x', padx=15, pady=12)
        
        # Vessel & Voyage Filters Section
        filters_header = ctk.CTkFrame(left_panel, fg_color="transparent")
        filters_header.pack(fill='x', pady=(0, 8))
        
        ctk.CTkLabel(
            filters_header,
            text="Vessel & Voyage Filters",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor='w'
        ).pack(side='left')
        
        clear_filters_btn = ctk.CTkButton(
            filters_header,
            text="Clear All",
            command=self.clear_all_filters,
            fg_color=ACCENT_RED,
            hover_color="#e63946",
            width=100,
            height=32,
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        clear_filters_btn.pack(side='right')
        
        # Scrollable frame for filters
        filters_scroll = CTkScrollableFrame(
            left_panel,
            height=250,
            corner_radius=8,
            fg_color=BG_FRAME
        )
        filters_scroll.pack(fill='both', expand=True, pady=(0, 20))
        
        # Store vessel/voyage variable pairs
        self.vessel_voyage_vars = []
        
        # Create 10 vessel/voyage pairs
        for i in range(10):
            pair_frame = ctk.CTkFrame(filters_scroll, fg_color="transparent")
            pair_frame.pack(fill='x', pady=6, padx=5)
            
            # Pair number
            ctk.CTkLabel(
                pair_frame,
                text=f"{i+1}.",
                font=ctk.CTkFont(size=11),
                width=30,
                text_color=FG_SECONDARY
            ).pack(side='left', padx=(0, 8))
            
            # Vessel Name (uppercase)
            vessel_var = ctk.StringVar()
            if i < len(self.vessel_voyage_pairs):
                vessel_var.set(self.vessel_voyage_pairs[i].get('vessel', '').upper())
            
            # Add uppercase trace
            vessel_var.trace('w', self.uppercase_vessel_trace(vessel_var))
            
            vessel_entry = ctk.CTkEntry(
                pair_frame,
                textvariable=vessel_var,
                placeholder_text="Vessel Name",
                font=ctk.CTkFont(size=11),
                height=36,
                corner_radius=6,
                fg_color=BG_INPUT,
                border_width=1,
                border_color="#3d3d3d"
            )
            vessel_entry.pack(side='left', fill='x', expand=True, padx=(0, 8))
            
            # Voyage Number (uppercase)
            voyage_var = ctk.StringVar()
            if i < len(self.vessel_voyage_pairs):
                voyage_var.set(self.vessel_voyage_pairs[i].get('voyage', '').upper())
            
            # Add uppercase trace to voyage
            voyage_var.trace('w', self.uppercase_voyage_trace(voyage_var))
            
            voyage_entry = ctk.CTkEntry(
                pair_frame,
                textvariable=voyage_var,
                placeholder_text="Voyage #",
                font=ctk.CTkFont(size=11),
                height=36,
                corner_radius=6,
                fg_color=BG_INPUT,
                border_width=1,
                border_color="#3d3d3d"
            )
            voyage_entry.pack(side='left', fill='x', expand=True)
            
            # Store variables
            self.vessel_voyage_vars.append((vessel_var, voyage_var))
        
        # Output Location Section
        ctk.CTkLabel(
            left_panel,
            text="Save to:",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor='w'
        ).pack(anchor='w', pady=(0, 8))
        
        location_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        location_frame.pack(fill='x', pady=(0, 20))
        
        self.output_path_var = ctk.StringVar(value=str(self.output_dir))
        
        location_entry = ctk.CTkEntry(
            location_frame,
            textvariable=self.output_path_var,
            font=ctk.CTkFont(size=11),
            state='readonly',
            height=40,
            corner_radius=8,
            fg_color=BG_INPUT,
            border_width=1,
            border_color="#3d3d3d"
        )
        location_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        change_btn = ctk.CTkButton(
            location_frame,
            text="Change",
            command=self.change_output_location,
            fg_color=ACCENT_BLUE,
            hover_color="#1873cc",
            width=120,
            height=40,
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        change_btn.pack(side='right')
        
        # Action Buttons
        self.convert_btn = ctk.CTkButton(
            left_panel,
            text="✓ Convert to Template",
            command=self.process_file,
            fg_color=ACCENT_GREEN,
            hover_color="#00b359",
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(size=15, weight="bold"),
            state='disabled'
        )
        self.convert_btn.pack(fill='x', pady=(0, 12))
        
        clear_btn = ctk.CTkButton(
            left_panel,
            text="Clear",
            command=self.clear_file,
            fg_color=ACCENT_RED,
            hover_color="#e63946",
            height=45,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        clear_btn.pack(fill='x')
        
        # ========== RIGHT PANEL - CONSOLE ==========
        
        console_header = ctk.CTkFrame(right_panel, fg_color="transparent")
        console_header.pack(fill='x', pady=(0, 10))
        
        ctk.CTkLabel(
            console_header,
            text="Console Log",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor='w'
        ).pack(side='left')
        
        clear_console_btn = ctk.CTkButton(
            console_header,
            text="Clear",
            command=lambda: self.console.delete('1.0', 'end'),
            fg_color=ACCENT_RED,
            hover_color="#e63946",
            width=90,
            height=32,
            corner_radius=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        clear_console_btn.pack(side='right')
        
        # Console frame
        console_frame = ctk.CTkFrame(right_panel, corner_radius=10, fg_color=BG_FRAME)
        console_frame.pack(fill='both', expand=True)
        
        # Using standard Text widget with dark styling
        self.console = tk.Text(
            console_frame,
            width=55,
            height=42,
            font=('Roboto', 10),
            bg=BG_CONSOLE,
            fg=FG_TEXT,
            insertbackground=FG_TEXT,
            bd=0,
            padx=15,
            pady=15,
            wrap='word',
            relief='flat',
            selectbackground="#404040",
            selectforeground=FG_TEXT
        )
        
        # Add scrollbar
        scrollbar = ctk.CTkScrollbar(console_frame, command=self.console.yview)
        scrollbar.pack(side='right', fill='y', padx=(0, 5), pady=5)
        self.console.configure(yscrollcommand=scrollbar.set)
        self.console.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Configure color tags for console
        self.console.tag_config('success', foreground='#00d26a')
        self.console.tag_config('error', foreground='#ff4757')
        self.console.tag_config('warning', foreground='#ffa502')
        self.console.tag_config('info', foreground='#1e90ff')
        self.console.tag_config('stat', foreground='#00d26a')
        self.console.tag_config('default', foreground=FG_TEXT)
        
        # Initial log
        self.log(f"🎉 Welcome to Seal Check Converter!")
        self.log(f"📦 Version: v{__version__}")
        self.log("📁 Drag & drop files or click to browse.")
        self.log("🚢 For UnitList: Enter Vessel + Voyage pairs")
        self.log("📝 Each voyage creates a separate output file")
        self.log("🔍 Checking for updates...")
    
    def log(self, message):
        """Add color-coded message to console log"""
        # Determine color based on message content
        if any(indicator in message for indicator in ['✅', '✓', 'Success', 'COMPLETED']):
            tag = 'success'
        elif any(indicator in message for indicator in ['❌', '✗', 'ERROR', 'Failed']):
            tag = 'error'
        elif any(indicator in message for indicator in ['⚠️', 'WARNING', 'Warning', 'No containers']):
            tag = 'warning'
        elif any(indicator in message for indicator in ['ℹ️', 'INFO', 'Filtering', 'ignored']):
            tag = 'info'
        elif any(indicator in message for indicator in ['📊', 'STATISTICS', 'Containers:', 'Seals:']):
            tag = 'stat'
        else:
            tag = 'default'
        
        self.console.insert('end', f"{message}\n", tag)
        self.console.see('end')
        self.root.update_idletasks()
    
    def change_output_location(self):
        """Select output directory"""
        new_dir = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir
        )
        if new_dir:
            self.output_dir = Path(new_dir)
            self.output_path_var.set(str(self.output_dir))
            self.save_config()
    
    def browse_source(self):
        """Browse for source file"""
        file = filedialog.askopenfilename(
            title="Select Source Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file:
            self.set_source_file(file)
    
    def on_drop(self, event):
        """Handle drag and drop event"""
        files = self.root.tk.splitlist(event.data)
        excel_files = [f for f in files if f.lower().endswith(('.xlsx', '.xls'))]
        
        if excel_files:
            self.set_source_file(excel_files[0])
        else:
            self.log("⚠️ Please drop an Excel file")
    
    def set_source_file(self, file_path):
        """Set the source file"""
        self.selected_file = Path(file_path)
        filename = self.selected_file.name
        
        # Update file display
        self.file_label.configure(text=f"✓ {filename}", text_color=ACCENT_GREEN)
        
        self.log(f"\n✅ Selected: {filename}")
        self.update_convert_button()
    
    def clear_file(self):
        """Clear the selected file"""
        if self.selected_file:
            self.selected_file = None
            self.file_label.configure(text="No file selected", text_color=FG_SECONDARY)
            self.log("\n🗑️ File cleared")
            self.update_convert_button()
    
    def clear_all_filters(self):
        """Clear all vessel and voyage filters"""
        for vessel_var, voyage_var in self.vessel_voyage_vars:
            vessel_var.set('')
            voyage_var.set('')
        self.log("\n🗑️ Cleared all voyage filters")
    
    def update_convert_button(self):
        """Enable/disable convert button"""
        if self.template_path.exists() and self.selected_file:
            self.convert_btn.configure(state='normal')
        else:
            self.convert_btn.configure(state='disabled')
    
    def generate_output_filename(self, vessel, pol):
        """Generate output filename based on POL value"""
        now = datetime.datetime.now()
        time_str = now.strftime("%H%M")
        number = now.strftime("%S%f")[:4]
        
        pol_upper = str(pol).upper() if pol else ""
        
        if pol_upper == 'ECGYE':
            return f"{vessel} - NAPORTEC - {time_str} - {number}.xlsx"
        elif pol_upper == 'PACCT':
            return f"{vessel} - COLON - {time_str} - {number}.xlsx"
        elif pol_upper == 'PEPIO':
            return f"{vessel} - PISCO - {time_str} - {number}.xlsx"
        elif pol_upper == 'PSA-RODMAN':
            return f"{vessel} - RODMAN - {time_str} - {number}.xlsx"
        elif pol_upper == 'PSA-RODMAN':
            return f"{vessel} - RODMAN - {time_str} - {number}.xlsx"
        else:
            return f"{vessel} - {time_str} - {number}.xlsx"
    
    def process_file(self):
        """Process the file using core converter"""
        if not self.template_path.exists():
            messagebox.showerror("Template Missing", 
                f"Template file not found:\n{self.template_path}")
            return
        
        if not self.selected_file:
            messagebox.showwarning("No File", "Please select a source file.")
            return
        
        # Save config
        self.save_config()
        
        # Collect vessel/voyage pairs
        vessel_voyage_pairs = []
        for vessel_var, voyage_var in self.vessel_voyage_vars:
            vessel = vessel_var.get().strip().upper()
            voyage = voyage_var.get().strip().upper()
            if vessel and voyage:
                vessel_voyage_pairs.append((vessel, voyage))
        
        # Disable button
        self.convert_btn.configure(state='disabled')
        
        # Process in thread
        def process_thread():
            total_files = 0
            total_containers = 0
            total_seals = 0
            
            # If vessel/voyage pairs provided, process each separately (UnitList mode or COLON YARD mode)
            if vessel_voyage_pairs:
                for vessel, voyage in vessel_voyage_pairs:
                    # Temporary filename for conversion
                    temp_filename = f"TEMP_{voyage}.xlsx"
                    temp_path = self.output_dir / temp_filename
                    
                    self.log(f"\n▶️  Processing filter: {voyage}")
                    
                    result = self.converter.convert(
                        source_path=self.selected_file,
                        output_path=temp_path,
                        progress_callback=self.log,
                        voyage_filters=[voyage],
                        carrier_filters=[voyage]  # For COLON YARD, voyage field is used as carrier
                    )
                    
                    if result and result['containers'] > 0:
                        # Get POL and vessel from result for smart naming
                        pol = result.get('pol')
                        # Use the vessel name from the filter ONLY for this specific filtered output
                        if pol and vessel:
                            smart_filename = self.generate_output_filename(vessel, pol)
                            smart_path = self.output_dir / smart_filename
                            
                            # Rename the file
                            if temp_path.exists():
                                if smart_path.exists():
                                    smart_path.unlink()  # Remove old file
                                temp_path.rename(smart_path)
                                self.log(f"📝 Output: {smart_filename}")
                        else:
                            self.log(f"⚠️  Could not determine POL or vessel for smart naming")
                        
                        total_files += 1
                        total_containers += result['containers']
                        total_seals += result['seals']
                    elif result:
                        self.log(f"⚠️  No containers found for voyage {voyage}")
                        self.log(f"   File NOT created: {temp_filename}")
            else:
                # No filters - process entire file (standard mode)
                output_filename = f"COMPLETED {self.selected_file.stem}.xlsx"
                output_path = self.output_dir / output_filename
                
                result = self.converter.convert(
                    source_path=self.selected_file,
                    output_path=output_path,
                    progress_callback=self.log,
                    voyage_filters=None  # No filtering
                )
                
                if result:
                    total_files = 1
                    total_containers = result['containers']
                    total_seals = result['seals']
            
            if total_files > 0:
                messagebox.showinfo(
                    "Success!",
                    f"Conversion completed!\n\n"
                    f"Files created: {total_files}\n"
                    f"Total containers: {total_containers}\n"
                    f"Total seals: {total_seals}\n\n"
                    f"Saved to:\n{self.output_dir}"
                )
                
                # Ask to open folder
                if messagebox.askyesno("Open Folder?", "Open output folder?"):
                    import subprocess, platform
                    if platform.system() == 'Windows':
                        subprocess.Popen(['explorer', str(self.output_dir)])
                    elif platform.system() == 'Darwin':
                        subprocess.Popen(['open', str(self.output_dir)])
                    else:
                        subprocess.Popen(['xdg-open', str(self.output_dir)])
            else:
                messagebox.showwarning("No Data", "No containers were processed.")
            
            # Re-enable
            self.convert_btn.configure(state='normal')
        
        threading.Thread(target=process_thread, daemon=True).start()

def get_downloads_folder(self):
    """Get the Downloads folder path for any user"""
    import platform
    
    system = platform.system()
    home = Path.home()  # Automatically gets current user's home
    
    if system == 'Windows':
        downloads = home / 'Downloads'  # Works for ANY Windows user
    elif system == 'Darwin':  # macOS
        downloads = home / 'Downloads'
    else:  # Linux
        downloads = home / 'Downloads'

def main():
    # Create root window with TkinterDnD support
    root = TkinterDnD.Tk()
    
    # CRITICAL: Apply dark theme BEFORE creating GUI
    root.configure(bg=BG_MAIN)
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    
    app = SealCheckConverterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()