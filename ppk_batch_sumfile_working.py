import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
import re
import tempfile
import datetime
import shutil
import json
import os


class RoverObservation:
    def __init__(self, filepath):
        self.filepath = Path(filepath)
        self.name = self.filepath.stem
        self.date, self.time = self.extract_date_time()

    def extract_date_time(self):
        try:
            with open(self.filepath, 'r') as f:
                for line in f:
                    if line.startswith('>'):
                        match = re.search(
                            r'(\d{4})\s+(\d{2})\s+(\d{2})\s+(\d{2})\s+(\d{2})\s+(\d{2}\.\d+)',
                            line
                        )
                        if match:
                            date_str = f"{match.group(1)}{match.group(2)}{match.group(3)}"
                            time_str = f"{match.group(4)}{match.group(5)}{match.group(6).split('.')[0]}"
                            return date_str, time_str
            return None, None
        except Exception as e:
            print(f"Error reading file {self.filepath}: {e}")
            return None, None


class BaseObservation:
    def __init__(self, filepath):
        self.filepath = Path(filepath)
        self.name = self.filepath.stem
        self.date, self.time = self.extract_date_time()

    def extract_date_time(self):
        try:
            with open(self.filepath, 'r') as f:
                lines = f.readlines()
                if len(lines) < 2:
                    print(f"File '{self.filepath.name}' has less than 2 lines.")
                    return None, None
                line_2 = lines[1].strip()
                match = re.search(r'(\d{8})\s+(\d{6})', line_2)
                if match:
                    return match.group(1), match.group(2)
            return None, None
        except Exception as e:
            print(f"Error reading file {self.filepath}: {e}")
            return None, None


class NavigationFile(BaseObservation):
    pass


class PPKProcessorGUI:
    CONFIG_FILE = 'config.json'

    def __init__(self, master):
        self.master = master
        self.master.title("Batch PPK Processor")
        self.master.geometry("1000x400")
        self.master.minsize(400, 800)

        # Initialize lists first
        self.rover_obs_list = []
        self.base_obs_list = []
        self.nav_obs_list = []

        # Initialize configuration settings
        self.config_settings = {
            'pos1-antheight': tk.StringVar(value="0.0"),
            'pos2-antheight': tk.StringVar(value="0.0"),
            'ant1-antdelu': tk.StringVar(value="0.0")
        }

        # Indicateur de modifications non sauvegardées
        self.unsaved_changes = False

        # Lier la fonction de fermeture
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

        # Create main frames
        self.paned_window = ttk.PanedWindow(self.master, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # Create left and right frames
        self.left_frame = ttk.Frame(self.paned_window, width=700)
        self.right_frame = ttk.Frame(self.paned_window)
        
        # Add frames to paned window
        self.paned_window.add(self.left_frame, weight=4)
        self.paned_window.add(self.right_frame, weight=4)
        
        # Prevent left frame from shrinking
        self.left_frame.pack_propagate(False)

        # Initialize antenna variables
        self.antenna_types = {
            "EMLID RS2": -0.135,
            "FOIF A30": -0.088
        }
        
        self.selected_antenna = tk.StringVar(value=list(self.antenna_types.keys())[0])
        self.manual_offset = tk.StringVar(value="-0.00")
        self.total_offset = tk.StringVar(value=str(self.antenna_types[self.selected_antenna.get()]))

        # Create UI elements
        self.create_menu()
        self.create_left_frame()
        self.create_right_frame()

        # Add after other initializations
        self.stats_file = Path("ppk_statistics.json")
        self.load_statistics()
        self.stats_file = Path("ppk_statistics.json")
        self.stats_data = {}
        self.load_statistics()
        self.load_config()

        # Add this line after other initializations
        self.current_project_path = None
        
        # Add keyboard binding for Ctrl+S
        self.master.bind('<Control-s>', lambda e: self.save_project())

        self.base_coordinates = {}  # Dictionnaire pour stocker {date: (lat, lon, height)}

        # Initialise la liste des fichiers .sum importés
        self.sum_files = []

    def create_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save", command=self.save_project, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As...", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.master.quit)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Manual", command=self.show_manual)
        help_menu.add_command(label="Support", command=self.show_support)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

    def show_manual(self):
        """Show user manual in new window"""
        manual = tk.Toplevel(self.master)
        manual.title("User Manual")
        manual.geometry("600x400")
        
        text = scrolledtext.ScrolledText(manual, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        manual_content = """
        **Instruction Manual**

**1. Getting Started**  
- **Select RTKLIB Executable:** Click **Browse** under "RTKLIB Executable" to choose your `rnx2rtkp.exe` (or similar) file.  
- **Load Configuration File:** Under "Config File (ppk.conf)", click **Browse** to select your `ppk.conf` file. Click **Edit Config** to view or modify its contents.

**2. Base Coordinates Setup**  
- **Manual Mode:** Select "Mode Manuel" and enter base Latitude, Longitude, and Height. Click **Appliquer** to update the config with these coordinates.  
- **Automatic Mode (.sum Files):** Select "Mode Automatique (fichiers .sum)" and click **Importer .sum** to add `.sum` files. The software will automatically apply coordinates from these files during processing.

**3. Antenna Configuration**  
- Choose your antenna type from the dropdown (e.g., "EMLID RS2").  
- Enter any manual offset needed.  
- Click **Appliquer Configuration Antenne** to update the configuration file with these antenna parameters.

**4. Adding Files**  
- **Rover Files (e.g., *.24o):** Click **Add Rover Files**, select your rover observation files, and confirm.  
- **Base Files (e.g., *.24O):** Click **Add Base Files**, select your base observation files. The application will attempt to match them to rover files by date.  
- **Navigation Files (e.g., *.24P):** Click **Add Navigation Files** and choose your navigation files.

Use the **Delete Selected** button in each section to remove unwanted files.

**5. Output Directory**  
- Choose an output directory for processed results by clicking **Browse** under "Output Directory."

**6. Running Batch Processing**  
- Ensure all required files are loaded, coordinates are set, and the configuration is correct.  
- Click **Run Batch PPK Processing** to start.  
- A progress bar shows processing status, and logs will appear in the "Status and Logs" tab.

**7. Viewing Results & Statistics**  
- Processed results are saved as `.pos` files in the chosen output directory.  
- The "Quality Statistics" tab displays solution quality percentages.  
- Double-click on a file entry in the statistics to open the corresponding `.pos` file.

**8. Project Management**  
- **New Project:** Clears all current selections and settings.  
- **Open Project:** Load previously saved `.ppk` project files to restore all settings.  
- **Save / Save As:** Save the current project configuration, file paths, logs, and statistics to a `.ppk` file.  
- Press **Ctrl+S** to quickly save your current project if already named.

**9. Logs & Exporting**  
- Check the "Status and Logs" tab for detailed messages and any errors.  
- Click **Clear Logs** to reset.  
- Use the "Export Logs" feature (if available) to save logs to a `.txt` file for record-keeping.

**10. Help & Support**  
- **User Manual:** Access the integrated manual via **Help > User Manual**.  
- **Support:** Check **Help > Support** for contact and GitHub repository information.  
- **About:** View version and author information from **Help > About**.

---

Follow these steps to prepare and run batch PPK data processing efficiently within the application.
        """
        text.insert(tk.END, manual_content)
        text.config(state='disabled')

    def show_support(self):
        """Show support information"""
        messagebox.showinfo("Support", 
            "For technical support:\n\n"
            "Email: tofunori@gmail.com\n"
            "Website: https://github.com/tofunori/GNSS-utilities-app")
           
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo("About PPK Processor",
            "PPK Processor v1.1\n\n"
            "A batch processing tool for GNSS post-processing using RTKLIB\n\n"
            "Created by: Thierry Laurent St-Pierre\n"
            "Copyright © 2024")

    def new_project(self):
        if messagebox.askyesno("New Project", "Create new project? This will clear current settings."):
            self.clear_all_fields()
            self.unsaved_changes = True

    def clear_all_fields(self):
        # Clear all variables and lists
        self.exec_path_var.set("")
        self.config_path_var.set("")
        self.rover_obs_list.clear()
        self.base_obs_list.clear()
        self.nav_obs_list.clear()
        
        # Clear config settings
        for var in self.config_settings.values():
            var.set("")
            
        # Update UI
        self.update_file_lists()
        self.append_log("New project created.\n")
        self.unsaved_changes = True

    def save_project(self):
        if not self.current_project_path:
            # First time saving - need to get file path
            file_path = filedialog.asksaveasfilename(
                defaultextension=".ppk",
                filetypes=[("PPK Project files", "*.ppk"), ("All files", "*.*")]
            )
            if not file_path:  # User cancelled
                return
            self.current_project_path = file_path
        
        # Quick save to existing path
        self._do_save(self.current_project_path)
        self.append_log(f"Project saved to {self.current_project_path}\n")
        self.unsaved_changes = False

    def save_project_as(self):
        # Always prompt for new location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".ppk",
            filetypes=[("PPK Project files", "*.ppk"), ("All files", "*.*")]
        )
        if file_path:
            self.current_project_path = file_path  # Update current path
            self._do_save(file_path)
            self.unsaved_changes = False

    def _do_save(self, file_path):
        try:
            # Debug logging
            self.append_log("Starting project save...\n")
            
            # Get base coordinates
            base_coords = {
                'latitude': self.base_lat_var.get(),
                'longitude': self.base_lon_var.get(),
                'height': self.base_height_var.get()
            }
            
            # Debug log pour vérifier les coordonnées
            self.append_log(f"Saving base coordinates: {base_coords}\n")
            
            # Get log content
            log_content = self.log_text.get(1.0, tk.END).strip()
            
            # Build project data
            project_data = {
                'executable_path': str(self.exec_path_var.get()),
                'config_path': str(self.config_path_var.get()),
                'rover_files': [str(rover.filepath) for rover in self.rover_obs_list],
                'base_files': [str(base.filepath) for base in self.base_obs_list],
                'nav_files': [str(nav.filepath) for nav in self.nav_obs_list],
                'sum_files': [str(sum_file) for sum_file in self.sum_files],  # Ajout des fichiers .sum
                'config_settings': {k: str(v.get()) for k, v in self.config_settings.items()},
                'base_coordinates': base_coords,  # Ajout des coordonnées ici
                'logs': log_content,
                'statistics': [
                    {
                        'file': str(self.stats_tree.item(item)['values'][0]),
                        'status': str(self.stats_tree.item(item)['values'][1]),
                        'processing_time': str(self.stats_tree.item(item)['values'][2])
                    }
                    for item in self.stats_tree.get_children()
                ]
            }

            # Save with explicit encoding
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, indent=4, ensure_ascii=False)
            
            self.append_log(f"Project successfully saved to {file_path}\n")
            messagebox.showinfo("Success", f"Project saved to {file_path}")
            
        except Exception as e:
            error_msg = f"Failed to save project: {str(e)}"
            self.append_log(f"Error: {error_msg}\n")
            messagebox.showerror("Error", error_msg)
            raise  # Pour le débogage

    def open_project(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PPK Project files", "*.ppk"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r') as f:
                project_data = json.load(f)

            # Set the current project path
            self.current_project_path = file_path

            # Load project data
            self.exec_path_var.set(project_data['executable_path'])
            self.config_path_var.set(project_data['config_path'])
            
            # Charger les coordonnées de base silencieusement
            if 'base_coordinates' in project_data:
                coords = project_data['base_coordinates']
                self.base_lat_var.set(coords.get('latitude', ''))
                self.base_lon_var.set(coords.get('longitude', ''))
                self.base_height_var.set(coords.get('height', ''))
                
                # Appliquer les coordonnées au fichier de configuration silencieusement
                if all(coords.values()):
                    try:
                        # Modification de apply_base_coordinates pour éviter les messages
                        self.apply_base_coordinates(show_messages=False)
                    except Exception as e:
                        self.append_log(f"Error applying base coordinates: {str(e)}\n")
            else:
                self.append_log("No base coordinates found in project file\n")

            # Load rover files
            self.rover_obs_list.clear()  # Clear existing files first
            self.rover_listbox.delete(0, tk.END)  # Clear the listbox
            for rover_path in project_data['rover_files']:
                if os.path.exists(rover_path):
                    rover = RoverObservation(Path(rover_path))
                    self.rover_obs_list.append(rover)
                    date_display = rover.date if rover.date else "Unknown"
                    time_display = rover.time if rover.time else "Unknown"
                    display_name = f"{rover.filepath.name} (Date: {date_display}, Time: {time_display})"
                    self.rover_listbox.insert(tk.END, display_name)
                else:
                    self.append_log(f"Warning: Rover file not found: {rover_path}\n")

            # Load base files
            self.base_obs_list.clear()  # Clear existing files first
            self.base_listbox.delete(0, tk.END)  # Clear the listbox
            for base_path in project_data['base_files']:
                if os.path.exists(base_path):
                    base = BaseObservation(Path(base_path))
                    self.base_obs_list.append(base)
                    date_display = base.date if base.date else "Unknown"
                    time_display = base.time if base.time else "Unknown"
                    display_name = f"{base.filepath.name} (Date: {date_display}, Time: {time_display})"
                    self.base_listbox.insert(tk.END, display_name)
                else:
                    self.append_log(f"Warning: Base file not found: {base_path}\n")

            # Load navigation files
            self.nav_obs_list.clear()  # Clear existing files first
            self.nav_listbox.delete(0, tk.END)  # Clear the listbox
            for nav_path in project_data['nav_files']:
                if os.path.exists(nav_path):
                    nav = NavigationFile(Path(nav_path))
                    self.nav_obs_list.append(nav)
                    date_display = nav.date if nav.date else "Unknown"
                    time_display = nav.time if nav.time else "Unknown"
                    display_name = f"{nav.filepath.name} (Date: {date_display}, Time: {time_display})"
                    self.nav_listbox.insert(tk.END, display_name)
                else:
                    self.append_log(f"Warning: Navigation file not found: {nav_path}\n")

            # Load config settings
            for key, value in project_data.get('config_settings', {}).items():
                if key in self.config_settings:
                    self.config_settings[key].set(value)

            # Restore logs
            if 'logs' in project_data:
                self.log_text.config(state='normal')
                self.log_text.delete(1.0, tk.END)
                self.log_text.insert(tk.END, project_data['logs'])
                self.log_text.config(state='disabled')

            # Restore statistics
            self.stats_tree.delete(*self.stats_tree.get_children())  # Clear existing stats
            if 'statistics' in project_data:
                for stat in project_data['statistics']:
                    self.stats_tree.insert('', tk.END, values=(
                        stat['file'],
                        stat['status'],
                        stat['processing_time']
                    ))

            # Charger les fichiers .sum
            self.sum_files.clear()
            self.sum_files_listbox.delete(0, tk.END)
            for sum_path in project_data.get('sum_files', []):
                if os.path.exists(sum_path):
                    self.sum_files.append(sum_path)
                    self.sum_files_listbox.insert(tk.END, Path(sum_path).name)
                else:
                    self.append_log(f"Fichier .sum introuvable: {sum_path}\n")

            self.append_log(f"Project loaded from {file_path}\n")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project: {str(e)}")
            self.append_log(f"Error loading project: {str(e)}\n")

    def add_file_to_list(self, filepath, file_list):
        # Helper method to add files to appropriate lists
        if filepath.suffix.lower() in ['.obs', '.OBS', '.nav', '.NAV']:
            file_list.append(BaseObservation(filepath))

    def create_left_frame(self):
        # Create a canvas with scrollbar
        canvas = tk.Canvas(self.left_frame)
        scrollbar = tk.Scrollbar(self.left_frame, orient="vertical", command=canvas.yview, width=20)
        
        # Create a frame inside canvas for content
        self.left_content = ttk.Frame(canvas)
        
        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Create window in canvas
        canvas_window = canvas.create_window((0, 0), window=self.left_content, anchor="nw")
        
        # Configure canvas scrolling
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.left_content.bind("<Configure>", configure_canvas)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))

        # Add the content to left_content
        self.create_left_content()

    def create_left_content(self):
        # RTKLIB Executable Selection
        exec_frame = ttk.LabelFrame(self.left_content, text="RTKLIB Executable")
        exec_frame.pack(fill=tk.X, padx=5, pady=5)

        self.exec_path_var = tk.StringVar()
        exec_entry = ttk.Entry(exec_frame, textvariable=self.exec_path_var)
        exec_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        exec_browse = ttk.Button(exec_frame, text="Browse", command=self.browse_executable)
        exec_browse.pack(side=tk.RIGHT, padx=5, pady=5)

        # Config File Selection
        config_frame = ttk.LabelFrame(self.left_content, text="Config File (ppk.conf)")
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        self.config_path_var = tk.StringVar()
        config_entry = ttk.Entry(config_frame, textvariable=self.config_path_var)
        config_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        config_browse = ttk.Button(config_frame, text="Browse", command=self.browse_config)
        config_browse.pack(side=tk.RIGHT, padx=5, pady=5)

        edit_config_button = ttk.Button(config_frame, text="Edit Config", command=self.edit_config_file)
        edit_config_button.pack(side=tk.RIGHT, padx=5, pady=5)

        # Base Coordinates Frame
        base_coord_frame = ttk.LabelFrame(self.left_content, text="Coordonnées de la Base")
        base_coord_frame.pack(fill=tk.X, padx=5, pady=5)

        # Mode selection
        mode_frame = ttk.Frame(base_coord_frame)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)

        self.coord_mode = tk.StringVar(value="manual")  # Default to manual mode
        ttk.Radiobutton(mode_frame, text="Mode Manuel", variable=self.coord_mode, 
                       value="manual", command=self.update_coord_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Mode Automatique (fichiers .sum)", 
                       variable=self.coord_mode, value="auto", 
                       command=self.update_coord_mode).pack(side=tk.LEFT, padx=5)

        # Variables pour les coordonnées
        self.base_lat_var = tk.StringVar()
        self.base_lon_var = tk.StringVar()
        self.base_height_var = tk.StringVar()

        # Frame pour les coordonnées
        coords_frame = ttk.Frame(base_coord_frame)
        coords_frame.pack(fill=tk.X, padx=5, pady=5)

        # Latitude
        ttk.Label(coords_frame, text="Latitude (deg):").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.lat_entry = ttk.Entry(coords_frame, textvariable=self.base_lat_var, width=20)
        self.lat_entry.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        # Longitude
        ttk.Label(coords_frame, text="Longitude (deg):").grid(row=1, column=0, padx=5, pady=2, sticky='e')
        self.lon_entry = ttk.Entry(coords_frame, textvariable=self.base_lon_var, width=20)
        self.lon_entry.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # Hauteur
        ttk.Label(coords_frame, text="Hauteur (m):").grid(row=2, column=0, padx=5, pady=2, sticky='e')
        self.height_entry = ttk.Entry(coords_frame, textvariable=self.base_height_var, width=20)
        self.height_entry.grid(row=2, column=1, padx=5, pady=2, sticky='w')

        # Boutons
        button_frame = ttk.Frame(base_coord_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.import_sum_button = ttk.Button(button_frame, text="Importer .sum", 
                                          command=self.import_sum_file)
        self.import_sum_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Appliquer", 
                  command=self.apply_base_coordinates).pack(side=tk.LEFT, padx=5)

        # Liste des fichiers .sum
        self.sum_list_label = ttk.Label(base_coord_frame, text="Fichiers .sum importés:")
        self.sum_list_label.pack(anchor='w', padx=5)
        
        list_frame = ttk.Frame(base_coord_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5)
        
        # Création de la Listbox avec scrollbar
        self.sum_files_listbox = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED)
        self.sum_files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.sum_files_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.sum_files_listbox.yview)
        
        # Liste pour stocker les chemins des fichiers
        self.sum_files = []

        self.delete_sum_button = ttk.Button(base_coord_frame, text="Supprimer Sélection", 
                                        command=self.delete_sum_files)
        self.delete_sum_button.pack(pady=5)

        # Initialize mode
        self.update_coord_mode()

        # Antenna Configuration Frame
        antenna_frame = ttk.LabelFrame(self.left_content, text="Configuration Antenne")
        antenna_frame.pack(fill=tk.X, padx=5, pady=5)

        # Type d'antenne
        ttk.Label(antenna_frame, text="Type d'antenne:").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.antenna_type_var = tk.StringVar(value="EMLID RS2")
        antenna_type_combo = ttk.Combobox(antenna_frame, textvariable=self.antenna_type_var, values=["EMLID RS2"])
        antenna_type_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky='w')

        # Offset total
        ttk.Label(antenna_frame, text="Offset total (m):").grid(row=1, column=0, padx=5, pady=2, sticky='e')
        self.antenna_offset_var = tk.StringVar(value="-0.135")
        ttk.Entry(antenna_frame, textvariable=self.antenna_offset_var, width=20).grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # Offset manuel
        ttk.Label(antenna_frame, text="Offset manuel (m):").grid(row=2, column=0, padx=5, pady=2, sticky='e')
        self.manual_offset_var = tk.StringVar(value="-0.00")
        ttk.Entry(antenna_frame, textvariable=self.manual_offset_var, width=20).grid(row=2, column=1, padx=5, pady=2, sticky='w')

        # Bouton Appliquer Configuration Antenne
        ttk.Button(antenna_frame, text="Appliquer Configuration Antenne", 
                  command=self.apply_antenna_config).grid(row=3, column=0, columnspan=3, pady=5)

        # Rover Files Selection
        rover_frame = ttk.LabelFrame(self.left_content, text="Select Rover Files")
        rover_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        add_rover_delete_frame = ttk.Frame(rover_frame)
        add_rover_delete_frame.pack(anchor='w', padx=5, pady=5)

        ttk.Button(add_rover_delete_frame, text="Add Rover Files", command=self.add_rover_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(add_rover_delete_frame, text="Delete Selected", command=self.delete_rover_files).pack(side=tk.LEFT)

        # Créer un conteneur pour la listbox et scrollbar
        rover_container = ttk.Frame(rover_frame)
        rover_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.rover_listbox = tk.Listbox(rover_container, selectmode=tk.MULTIPLE, height=5)
        self.rover_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        rover_scroll = tk.Scrollbar(rover_container, orient="vertical")
        rover_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Connecter la scrollbar et la listbox
        self.rover_listbox.config(yscrollcommand=rover_scroll.set)
        rover_scroll.config(command=self.rover_listbox.yview)

        # Base Files Selection
        base_frame = ttk.LabelFrame(self.left_content, text="Select Base Files")
        base_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        add_base_delete_frame = ttk.Frame(base_frame)
        add_base_delete_frame.pack(anchor='w', padx=5, pady=5)

        ttk.Button(add_base_delete_frame, text="Add Base Files", command=self.add_base_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(add_base_delete_frame, text="Delete Selected", command=self.delete_base_files).pack(side=tk.LEFT)

        # Créer un conteneur pour la listbox et scrollbar
        base_container = ttk.Frame(base_frame)
        base_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.base_listbox = tk.Listbox(base_container, selectmode=tk.MULTIPLE, height=5)
        self.base_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        base_scroll = tk.Scrollbar(base_container, orient="vertical")
        base_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Connecter la scrollbar et la listbox
        self.base_listbox.config(yscrollcommand=base_scroll.set)
        base_scroll.config(command=self.base_listbox.yview)

        # Navigation Files Selection
        nav_frame = ttk.LabelFrame(self.left_content, text="Select Navigation Files")
        nav_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        add_nav_delete_frame = ttk.Frame(nav_frame)
        add_nav_delete_frame.pack(anchor='w', padx=5, pady=5)

        ttk.Button(add_nav_delete_frame, text="Add Navigation Files", command=self.add_nav_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(add_nav_delete_frame, text="Delete Selected", command=self.delete_nav_files).pack(side=tk.LEFT)

        # Créer un conteneur pour la listbox et scrollbar
        nav_container = ttk.Frame(nav_frame)
        nav_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.nav_listbox = tk.Listbox(nav_container, selectmode=tk.MULTIPLE, height=5)
        self.nav_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        nav_scroll = tk.Scrollbar(nav_container, orient="vertical")
        nav_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Connecter la scrollbar et la listbox
        self.nav_listbox.config(yscrollcommand=nav_scroll.set)
        nav_scroll.config(command=self.nav_listbox.yview)

        # Output Directory Selection
        output_frame = ttk.LabelFrame(self.left_content, text="Output Directory")
        output_frame.pack(fill=tk.X, padx=5, pady=5)

        self.output_path_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT, padx=5, pady=5)

        # Run Button and Progress Bar
        run_frame = ttk.Frame(self.left_content)
        run_frame.pack(fill=tk.X, padx=5, pady=10)

        self.run_button = ttk.Button(run_frame, text="Run Batch PPK Processing", command=self.start_batch_processing)
        self.run_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(run_frame, orient='horizontal', mode='determinate', variable=self.progress_var)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

    def apply_base_coordinates(self, show_messages=True):
        """Apply base coordinates to the config file"""
        try:
            # Validate inputs
            lat = float(self.base_lat_var.get())
            lon = float(self.base_lon_var.get())
            height = float(self.base_height_var.get())

            # Read current config file
            config_path = self.config_path_var.get()
            if not config_path:
                if show_messages:
                    messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier de configuration.")
                return

            with open(config_path, 'r') as f:
                config_lines = f.readlines()

            # Update coordinates in config
            for i, line in enumerate(config_lines):
                if line.startswith('ant2-postype'):
                    config_lines[i] = 'ant2-postype       =llh        # (0:llh,1:xyz,2:single,3:posfile,4:rinexhead,5:rtcm)\n'
                elif line.startswith('ant2-pos1'):
                    config_lines[i] = f'ant2-pos1          ={lat}  # (deg|m)\n'
                elif line.startswith('ant2-pos2'):
                    config_lines[i] = f'ant2-pos2          ={lon}  # (deg|m)\n'
                elif line.startswith('ant2-pos3'):
                    config_lines[i] = f'ant2-pos3          ={height}  # (m|m)\n'

            # Write updated config
            with open(config_path, 'w') as f:
                f.writelines(config_lines)

            self.append_log(f"Coordonnées de base mises à jour: Lat={lat}, Lon={lon}, H={height}\n")
            if show_messages:
                messagebox.showinfo("Succès", "Coordonnées de base mises à jour avec succès.")

        except ValueError:
            if show_messages:
                messagebox.showerror("Erreur", "Veuillez entrer des coordonnées valides (nombres décimaux).")
        except Exception as e:
            if show_messages:
                messagebox.showerror("Erreur", f"Erreur lors de la mise à jour des coordonnées: {str(e)}")

        self.unsaved_changes = True

    def create_right_frame(self):
        """Create right frame with logs and quality statistics"""
        notebook = ttk.Notebook(self.right_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Log Tab
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text='Status and Logs')

        # Create a container frame for logs and button
        log_container = ttk.Frame(log_tab)
        log_container.pack(fill=tk.BOTH, expand=True)

        # Add Clear Logs button at the top
        clear_logs_btn = ttk.Button(log_container, text="Clear Logs", command=self.clear_logs)
        clear_logs_btn.pack(side=tk.TOP, pady=(5,0), padx=5)

        # Add the log text widget
        self.log_text = scrolledtext.ScrolledText(log_container, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Statistics Tab
        stats_tab = ttk.Frame(notebook)
        notebook.add(stats_tab, text='Quality Statistics')

        # Create statistics frame
        stats_frame = ttk.LabelFrame(stats_tab, text="Processing Statistics")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Add buttons frame
        button_frame = ttk.Frame(stats_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        # Add Delete Selected button only
        self.delete_stat_btn = ttk.Button(
            button_frame,
            text="Delete Selected",
            command=self.delete_selected_statistics
        )
        self.delete_stat_btn.pack(side=tk.LEFT, padx=5)

        # Create Treeview with quality columns
        columns = ('File', 'q1 (%)', 'q2 (%)', 'q3 (%)', 'q4 (%)', 'q5 (%)')
        self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show='headings')
        
        # Configure columns
        for col in columns:
            self.stats_tree.heading(col, text=col)
            width = 150 if col == 'File' else 80
            self.stats_tree.column(col, width=width, anchor='center')

        # Add scrollbar
        scrollbar = ttk.Scrollbar(stats_frame, orient="vertical", command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack elements
        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind double-click to open file
        self.stats_tree.bind('<Double-1>', self.on_tree_double_click)

    def delete_selected_statistics(self):
        """Delete selected entries from statistics tree"""
        selected_items = self.stats_tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select entries to delete")
            return
            
        if messagebox.askyesno("Confirm Delete", "Delete selected statistics entries?"):
            for item in selected_items:
                self.stats_tree.delete(item)

    def delete_selected_row(self):
        """Delete selected row from statistics tree"""
        selection = self.stats_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a row to delete")
            return
            
        if messagebox.askyesno("Confirm", "Delete selected statistics row?"):
            for item in selection:
                self.stats_tree.delete(item)
            self.append_log("Deleted selected statistics row(s)\n")

    def view_selected_file(self):
        """Open the selected .pos file from the output directory"""
        # Check if output directory is set
        output_dir = self.output_path_var.get()
        if not output_dir:
            messagebox.showerror("Error", "No output directory selected")
            return

        # Get selected item from tree
        selection = self.stats_tree.selection()
        if not selection:
            # Get list of .pos files in output directory
            try:
                pos_files = [f for f in os.listdir(output_dir) if f.endswith('.pos')]
                if not pos_files:
                    messagebox.showinfo("No Files", "No .pos files found in output directory")
                    return
                    
                # Open the most recently modified .pos file
                latest_pos = max([os.path.join(output_dir, f) for f in pos_files], 
                               key=os.path.getmtime)
                self.view_pos_file(latest_pos)
                
            except Exception as e:
                messagebox.showerror("Error", f"Error accessing output directory: {str(e)}")
                return
        else:
            # If item selected, open that specific file
            item = selection[0]
            file_name = self.stats_tree.item(item)['values'][0]
            pos_file_path = os.path.join(output_dir, file_name)
            if os.path.exists(pos_file_path):
                self.view_pos_file(pos_file_path)
            else:
                messagebox.showerror("Error", f"File not found: {pos_file_path}")

    def browse_executable(self):
        path = filedialog.askopenfilename(title="Select RTKLIB Executable", filetypes=[("Executable Files", "*.exe")])
        if path:
            self.exec_path_var.set(path)

    def browse_config(self):
        path = filedialog.askopenfilename(title="Select Config File", filetypes=[("Config Files", "*.conf")])
        if path:
            self.config_path_var.set(path)
            self.load_config_settings(path)

    def load_config_settings(self, config_path):
        try:
            with open(config_path, 'r') as f:
                for line in f:
                    for key in self.config_settings:
                        if line.startswith(key):
                            value = line.split('=')[1].strip()
                            self.config_settings[key].set(value)
        except Exception as e:
            self.append_log(f"Error loading config settings: {e}\n")

    def save_modified_config(self, original_config_path):
        try:
            backup_path = original_config_path.with_suffix('.conf.bak')
            shutil.copy(original_config_path, backup_path)
            self.append_log(f"Backup of config file created at {backup_path}\n")

            with open(original_config_path, 'r') as f:
                config_lines = f.readlines()

            for i, line in enumerate(config_lines):
                for key, var in self.config_settings.items():
                    if line.startswith(key):
                        config_lines[i] = f"{key}={var.get()}\n"

            temp_config = tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.conf')
            temp_config.writelines(config_lines)
            temp_config.close()

            return Path(temp_config.name)
        except Exception as e:
            self.append_log(f"Error saving modified config: {e}\n")
            return None

    def edit_config_file(self):
        config_path = self.config_path_var.get()
        if not config_path:
            messagebox.showwarning("No File", "Please select a config file first.")
            return

        try:
            with open(config_path, 'r') as f:
                config_content = f.read()

            editor = tk.Toplevel(self.master)
            editor.title("Edit Config File")
            
            editor.resizable(True, True)
            editor.minsize(600, 400)
            
            editor.grid_rowconfigure(0, weight=1)
            editor.grid_columnconfigure(0, weight=1)

            text_widget = scrolledtext.ScrolledText(editor, wrap=tk.WORD)
            text_widget.insert(tk.END, config_content)
            text_widget.grid(row=0, column=0, columnspan=2, sticky='nsew', padx=10, pady=10)

            text_widget.config(undo=True, maxundo=-1)
            text_widget.bind('<Control-z>', lambda e: text_widget.edit_undo())
            text_widget.bind('<Control-y>', lambda e: text_widget.edit_redo())

            # Button frame to hold both buttons
            button_frame = ttk.Frame(editor)
            button_frame.grid(row=1, column=0, columnspan=2, pady=5)

            def save_changes():
                edited_content = text_widget.get(1.0, tk.END).strip()
                with open(config_path, 'w') as f:
                    f.write(edited_content)
                editor.destroy()

            def save_as():
                new_path = filedialog.asksaveasfilename(
                    defaultextension=".conf",
                    filetypes=[("Config Files", "*.conf"), ("All Files", "*.*")],
                    initialfile=Path(config_path).name
                )
                if new_path:
                    edited_content = text_widget.get(1.0, tk.END).strip()
                    with open(new_path, 'w') as f:
                        f.write(edited_content)
                    messagebox.showinfo("Success", f"File saved as: {new_path}")

            save_button = ttk.Button(button_frame, text="Save", command=save_changes)
            save_button.pack(side=tk.LEFT, padx=5)

            save_as_button = ttk.Button(button_frame, text="Save As", command=save_as)
            save_as_button.pack(side=tk.LEFT, padx=5)

        except Exception as e:
            messagebox.showerror("Error", f"Unable to open config file for editing: {e}")

    def add_rover_files(self):
        self.append_log("Opening file dialog...\n")
        paths = filedialog.askopenfilenames(title="Select Rover Observation Files", filetypes=[("Rover Files", "*.24o")])
        self.append_log(f"Selected paths: {paths}\n")
        
        if not paths:
            self.append_log("No files selected\n")
            return
            
        for path in paths:
            try:
                self.append_log(f"Processing file: {path}\n")
                if path not in [rover.filepath for rover in self.rover_obs_list]:
                    rover = RoverObservation(path)
                    self.rover_obs_list.append(rover)
                    date_display = rover.date if rover.date else "Unknown"
                    time_display = rover.time if rover.time else "Unknown"
                    display_name = f"{rover.filepath.name} (Date: {date_display}, Time: {time_display})"
                    
                    self.rover_listbox.insert(tk.END, display_name)
                    self.rover_listbox.see(tk.END)
                    self.master.update_idletasks()
                    
                    self.append_log(f"Successfully added rover file: {display_name}\n")
                else:
                    self.append_log(f"File already exists: {path}\n")
                    
            except Exception as e:
                self.append_log(f"Error processing file {path}: {str(e)}\n")
                messagebox.showerror("Error", f"Failed to add file: {path}\n{str(e)}")

        self.unsaved_changes = True

    def delete_rover_files(self):
        selected_indices = list(self.rover_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select at least one rover file to delete.")
            return

        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete {len(selected_indices)} selected rover file(s)?")
        if not confirm:
            return

        for index in sorted(selected_indices, reverse=True):
            self.rover_listbox.delete(index)
            del self.rover_obs_list[index]

        self.unsaved_changes = True

    def add_base_files(self):
        """Ajoute des fichiers de base à la liste."""
        paths = filedialog.askopenfilenames(
            title="Select Base Observation Files",
            filetypes=[("Base Files", "*.24O")]
        )
        for path in paths:
            if path not in [base.filepath for base in self.base_obs_list]:
                base = BaseObservation(path)
                self.base_obs_list.append(base)
                date_display = base.date if base.date else "Unknown"
                time_display = base.time if base.time else "Unknown"
                display_name = f"{base.filepath.name} (Date: {date_display}, Time: {time_display})"
                self.base_listbox.insert(tk.END, display_name)
                self.append_log(f"Fichier de base ajouté: {base.filepath.name}\n")
        
        # Mettre à jour les correspondances des fichiers .sum
        self.update_sum_files_display()

        self.unsaved_changes = True

    def delete_base_files(self):
        """Supprime les fichiers de base sélectionnés."""
        selected_indices = list(self.base_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select at least one base file to delete.")
            return
            
        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete {len(selected_indices)} selected base file(s)?")
        if not confirm:
            return

        for index in sorted(selected_indices, reverse=True):
            self.base_listbox.delete(index)
            del self.base_obs_list[index]
            
        # Mettre à jour les correspondances des fichiers .sum
        self.update_sum_files_display()

        self.unsaved_changes = True

    def add_nav_files(self):
        paths = filedialog.askopenfilenames(title="Select Navigation Files", filetypes=[("Navigation Files", "*.24P")])
        for path in paths:
            if path not in [nav.filepath for nav in self.nav_obs_list]:
                nav = NavigationFile(path)
                self.nav_obs_list.append(nav)
                date_display = nav.date if nav.date else "Unknown"
                time_display = nav.time if nav.time else "Unknown"
                display_name = f"{nav.filepath.name} (Date: {date_display}, Time: {time_display})"
                self.nav_listbox.insert(tk.END, display_name)

        self.unsaved_changes = True

    def delete_nav_files(self):
        selected_indices = list(self.nav_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select at least one navigation file to delete.")
            return

        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete {len(selected_indices)} selected navigation file(s)?")
        if not confirm:
            return

        for index in sorted(selected_indices, reverse=True):
            self.nav_listbox.delete(index)
            del self.nav_obs_list[index]

        self.unsaved_changes = True

    def browse_output(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.output_path_var.set(path)
            self.save_config()

    def start_batch_processing(self):
        # Validate inputs
        if not self.exec_path_var.get():
            messagebox.showerror("Error", "Please select the RTKLIB executable.")
            return
        if not self.config_path_var.get():
            messagebox.showerror("Error", "Please select the config file (ppk.conf).")
            return
        if not self.rover_obs_list:
            messagebox.showerror("Error", "Please add at least one rover observation file.")
            return
        if not self.base_obs_list:
            messagebox.showerror("Error", "Please add at least one base observation file.")
            return
        if not self.nav_obs_list:
            messagebox.showerror("Error", "Please add at least one navigation file.")
            return
        if not self.output_path_var.get():
            messagebox.showerror("Error", "Please select the output directory.")
            return

        # Disable widgets during processing
        self.disable_widgets()

        # Clear only logs, not statistics
        self.clear_logs()

        # Reset progress bar
        self.progress_var.set(0)

        # Start processing in a separate thread
        thread = threading.Thread(target=self.run_batch_processing)
        thread.start()

    def run_batch_processing(self):
        rtk_exe = Path(self.exec_path_var.get())
        original_config_file = Path(self.config_path_var.get())
        output_dir = Path(self.output_path_var.get())

        total_rovers = len(self.rover_obs_list)
        processed_rovers = 0

        # Si mode manuel, appliquer les coordonnées une seule fois au début
        if self.coord_mode.get() == "manual":
            try:
                self.update_config_with_base_coordinates(
                    self.base_lat_var.get(),
                    self.base_lon_var.get(),
                    self.base_height_var.get(),
                    original_config_file
                )
                self.append_log("Mode manuel : utilisation des coordonnées fixes pour tous les fichiers\n")
            except Exception as e:
                self.append_log(f"Erreur lors de la mise à jour des coordonnées manuelles : {str(e)}\n")
                return

        for rover in self.rover_obs_list:
            if not rover.date:
                self.append_log(f"No date found for rover file '{rover.filepath.name}'. Skipping.\n")
                processed_rovers += 1
                self.update_progress(processed_rovers, total_rovers)
                continue

            # Match based on date
            matching_bases = [
                base for base in self.base_obs_list
                if base.date == rover.date
            ]

            if not matching_bases:
                self.append_log(f"No matching base file found for rover file '{rover.filepath.name}' with date '{rover.date}'. Skipping.\n")
                processed_rovers += 1
                self.update_progress(processed_rovers, total_rovers)
                continue

            base_file = matching_bases[0]

            # En mode auto seulement, mettre à jour les coordonnées depuis le fichier .sum
            if self.coord_mode.get() == "auto":
                matching_sum = None
                for sum_file in self.sum_files:
                    sum_filename = Path(sum_file).stem
                    sum_date = self.extract_date_from_filename(sum_filename)
                    if sum_date == rover.date:
                        matching_sum = sum_file
                        break

                if matching_sum:
                    try:
                        sum_data = self.parse_sum_file(matching_sum)
                        self.append_log(f"Mise à jour des coordonnées depuis {Path(matching_sum).name} pour le traitement de {rover.filepath.name}\n")
                        
                        # Mettre à jour les coordonnées dans l'interface
                        self.base_lat_var.set(sum_data.get("Latitude (DD)", ""))
                        self.base_lon_var.set(sum_data.get("Longitude (DD)", ""))
                        self.base_height_var.set(sum_data.get("Elevation (m)", ""))
                        
                        # Mettre à jour le fichier de configuration
                        self.update_config_with_base_coordinates(
                            sum_data.get("Latitude (DD)", ""),
                            sum_data.get("Longitude (DD)", ""),
                            sum_data.get("Elevation (m)", ""),
                            original_config_file
                        )
                    except Exception as e:
                        self.append_log(f"Erreur lors de la mise à jour des coordonnées depuis {matching_sum}: {str(e)}\n")

            nav_file = next(
                (nav for nav in self.nav_obs_list if nav.date == base_file.date),
                None
            )

            if not nav_file:
                self.append_log(f"No matching navigation file found for base file '{base_file.filepath.name}' with date '{base_file.date}'. Skipping rover file '{rover.filepath.name}'.\n")
                processed_rovers += 1
                self.update_progress(processed_rovers, total_rovers)
                continue

            # Simple file naming - just use rover name
            base_filename = f"{rover.name}.pos"
            output_pos = output_dir / base_filename

            # If file exists, add simple numeric suffix
            if output_pos.exists():
                counter = 1
                while (output_dir / f"{rover.name}_{counter}.pos").exists():
                    counter += 1
                output_pos = output_dir / f"{rover.name}_{counter}.pos"
                self.append_log(f"File '{base_filename}' exists, saving as '{output_pos.name}'\n")

            command = [
                str(rtk_exe),
                "-k", str(original_config_file),
                "-o", str(output_pos),
                str(rover.filepath),
                str(base_file.filepath),
                str(nav_file.filepath),
            ]
            self.append_log(f"Executing: {' '.join(command)}\n")

            try:
                process = subprocess.run(command, capture_output=True, text=True)
                if process.returncode == 0:
                    self.append_log(f"Processing completed for Rover: {rover.filepath.name}\n")
                    # Compute statistics and update treeview immediately
                    stats = self.compute_quality_statistics(output_pos)
                    if stats:
                        stats_values = {f'q{q} (%)': f"{stats.get(q, 0):.2f}" for q in range(1, 6)}
                        row_values = (output_pos.name,) + tuple(stats_values.values())
                        self.master.after(0, lambda: self.stats_tree.insert('', 0, values=row_values))
                        self.latest_pos_file = str(output_pos)
                else:
                    self.append_log(f"Error for Rover: {rover.filepath.name}. Error Message: {process.stderr}\n")
            except Exception as e:
                self.append_log(f"Error during processing: {str(e)}\n")

            processed_rovers += 1
            self.update_progress(processed_rovers, total_rovers)

        # Re-enable widgets
        self.enable_widgets()
        self.append_log("Batch PPK Processing Completed.\n")
        messagebox.showinfo("Batch Processing", "Batch PPK Processing Completed.")

    def compute_quality_statistics(self, pos_file):
        try:
            with open(pos_file, 'r') as f:
                lines = f.readlines()
            data_start = next((i for i, line in enumerate(lines) if not line.startswith('%') and line.strip()), None)
            if data_start is None:
                self.append_log(f"Failed to find data in {pos_file}\n")
                return None
            data = [line.split() for line in lines[data_start:] if line.strip()]
            qualities = [int(row[5]) for row in data if len(row) > 5 and row[5].isdigit()]
            quality_counts = {}
            for q in qualities:
                quality_counts[q] = quality_counts.get(q, 0) + 1
            total = sum(quality_counts.values())
            if total == 0:
                self.append_log(f"No valid quality data found in {pos_file}\n")
                return None
            percentages = {q: (count / total) * 100 for q, count in quality_counts.items()}
            return percentages
        except Exception as e:
            self.append_log(f"Error computing quality statistics for {pos_file}: {e}\n")
            return None

    def populate_statistics_table(self, all_file_stats):
        """Populate statistics table with just filename and quality stats"""
        if not all_file_stats:
            self.append_log("No quality statistics to display.\n")
            return

        # Clear existing entries
        self.stats_tree.delete(*self.stats_tree.get_children())

        for filename, stats in all_file_stats.items():
            # Just use the .pos filename without timestamp
            stats_values = {f'q{q} (%)': f"{stats.get(q, 0):.2f}" for q in range(1, 6)}
            row_values = (filename,) + tuple(stats_values.values())
            self.stats_tree.insert('', tk.END, values=row_values)
            
            # Update latest file reference
            if hasattr(self, 'latest_pos_file'):
                self.latest_pos_file = os.path.join(self.output_path_var.get(), filename)

    def append_log(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def clear_logs(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def clear_statistics(self):
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

    def set_widget_state(self, widget, state):
        try:
            widget.configure(state=state)
        except tk.TclError:
            pass

    def disable_widgets(self):
        for child in self.left_content.winfo_children():
            if isinstance(child, (ttk.LabelFrame, ttk.Frame)):
                for widget in child.winfo_children():
                    self.set_widget_state(widget, 'disabled')
            else:
                self.set_widget_state(child, 'disabled')
        self.run_button.configure(state='disabled')

    def enable_widgets(self):
        for child in self.left_content.winfo_children():
            if isinstance(child, (ttk.LabelFrame, ttk.Frame)):
                for widget in child.winfo_children():
                    self.set_widget_state(widget, 'normal')
            else:
                self.set_widget_state(child, 'normal')
        self.run_button.configure(state='normal')

    def update_progress(self, processed, total):
        progress = (processed / total) * 100
        self.progress_var.set(progress)
        self.master.update_idletasks()

    def validate_antheight(self, P):
        """Validate that antenna height is a positive number."""
        if P == "":
            return True
        try:
            value = float(P)
            if value < 0:
                return False
            return True
        except ValueError:
            return False

    def export_logs(self):
        logs = self.log_text.get(1.0, tk.END).strip()
        if not logs:
            messagebox.showwarning("No Logs", "No logs available to export.")
            return

        export_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if export_path:
            try:
                with open(export_path, 'w', encoding='utf-8') as f:
                    f.write(logs)
                messagebox.showinfo("Export Successful", f"Logs exported to {export_path}.")
            except Exception as e:
                messagebox.showerror("Export Failed", f"An error occurred while exporting logs:\n{str(e)}")

    def update_file_lists(self):
        """Update all file listboxes with current file lists"""
        # Clear and update rover listbox
        self.rover_listbox.delete(0, tk.END)
        for rover in self.rover_obs_list:
            date_display = rover.date if rover.date else "Unknown"
            time_display = rover.time if rover.time else "Unknown"
            display_name = f"{rover.filepath.name} (Date: {date_display}, Time: {time_display})"
            self.rover_listbox.insert(tk.END, display_name)
        
        # Clear and update base listbox
        self.base_listbox.delete(0, tk.END)
        for base in self.base_obs_list:
            date_display = base.date if base.date else "Unknown"
            time_display = base.time if base.time else "Unknown"
            display_name = f"{base.filepath.name} (Date: {date_display}, Time: {time_display})"
            self.base_listbox.insert(tk.END, display_name)
        
        # Clear and update navigation listbox
        self.nav_listbox.delete(0, tk.END)
        for nav in self.nav_obs_list:
            date_display = nav.date if nav.date else "Unknown"
            time_display = nav.time if nav.time else "Unknown"
            display_name = f"{nav.filepath.name} (Date: {date_display}, Time: {time_display})"
            self.nav_listbox.insert(tk.END, display_name)

    def create_antenna_section(self, antheight_frame):
        # Antenna type dropdown
        ttk.Label(antheight_frame, text="Antenna Type:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        antenna_dropdown = ttk.Combobox(antheight_frame, 
                                      textvariable=self.selected_antenna,
                                      values=list(self.antenna_types.keys()),
                                      state="readonly")
        antenna_dropdown.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        # Rover antenna height entry
        ttk.Label(antheight_frame, text="Rover Antenna Height (pos1-antheight):").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.rover_height_entry = ttk.Entry(antheight_frame, textvariable=self.config_settings['pos1-antheight'])
        self.rover_height_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # Base antenna height entry
        ttk.Label(antheight_frame, text="Base Antenna Height (pos2-antheight):").grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.base_height_entry = ttk.Entry(antheight_frame, textvariable=self.config_settings['pos2-antheight'])
        self.base_height_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        # Manual offset entry (row 1)
        ttk.Label(antheight_frame, text="Offset manuel (m):").grid(row=1, column=2, padx=5, pady=5)
        self.manual_offset_entry = ttk.Entry(
            antenna_frame,
            textvariable=self.manual_offset,
            width=10
        )
        self.manual_offset_entry.grid(row=1, column=3, padx=5, pady=5)

        # Bind events
        self.antenna_combo.bind('<<ComboboxSelected>>', self.update_total_offset)
        self.manual_offset_entry.bind('<KeyRelease>', self.update_total_offset)

    def update_antenna_height(self, *args):
        try:
            # Get the current and previous antenna selections
            new_antenna = self.selected_antenna.get()
            current_height = float(self.config_settings['pos1-antheight'].get() or 0)
            
            # If changing from one antenna to another, first remove the old offset
            if hasattr(self, '_previous_antenna') and self._previous_antenna != "Select Antenna...":
                current_height += abs(self.antenna_types[self._previous_antenna])
                
            # Apply new offset
            new_offset = self.antenna_types[new_antenna]
            adjusted_height = current_height + new_offset if new_antenna != "Select Antenna..." else current_height
            
            # Update height value
            self.config_settings['pos1-antheight'].set(f"{adjusted_height:.3f}")
            
            # Store current antenna as previous for next change
            self._previous_antenna = new_antenna
            
            self.append_log(f"Antenna height adjusted for {new_antenna}: {new_offset:+.3f}m\n")
        except ValueError:
            self.append_log("Please enter a valid antenna height first\n")

    def ant1_antdelu(self):
        # Implementation of the antenna height calculation
        return 0.0  # Default value when no antenna is selected

    def update_total_height(self, event=None):
        """Update the total antenna height based on antenna type and additional height"""
        try:
            # Get antenna offset from selected antenna type
            antenna_type = self.selected_antenna.get()
            antenna_offset = self.antenna_types.get(antenna_type, 0.0)
            
            # Get additional height value
            try:
                additional_height = float(self.config_settings['pos1-antheight'].get())
            except ValueError:
                additional_height = 0.0
                
            # Calculate total height
            total_height = antenna_offset + additional_height
            
            # Update the height values
            self.config_settings['pos1-antheight'].set(f"{total_height:.4f}")
            self.config_settings['pos2-antheight'].set(f"{total_height:.4f}")
            
            # If we have a total height display label, update it
            if hasattr(self, 'total_height_value'):
                self.total_height_value.config(text=f"{total_height:.4f}")
                
        except Exception as e:
            print(f"Error updating total height: {e}")
            # Set default values if calculation fails
            self.config_settings['pos1-antheight'].set("0.0000")
            self.config_settings['pos2-antheight'].set("0.0000")
            if hasattr(self, 'total_height_value'):
                self.total_height_value.config(text="0.0000")

    def update_total_offset(self, *args):
        """Calculate and update the antenna offset for ant1-antdelu"""
        try:
            # Get antenna offset
            antenna_offset = self.antenna_types[self.selected_antenna.get()]
            manual_offset = float(self.manual_offset.get() or 0.0)
            total = antenna_offset + manual_offset
            
            # Update displays and config
            self.total_offset.set(f"{total:.3f}")
            self.config_settings['ant1-antdelu'].set(f"{total:.3f}")  # Update ant1-antdelu
            
        except (ValueError, KeyError):
            self.total_offset.set("Error")
            self.config_settings['ant1-antdelu'].set("0.000")

    def view_pos_file(self, pos_file_path):
        """Open .pos file with default system viewer"""
        try:
            os.startfile(pos_file_path)  # Windows specific
            self.append_log(f"Opening result file: {pos_file_path}\n")
        except Exception as e:
            self.append_log(f"Error opening file: {str(e)}\n")
            messagebox.showerror("Error", f"Failed to open file: {pos_file_path}\n{str(e)}")

    def on_tree_double_click(self, event):
        """Handle double click on statistics tree item"""
        item = self.stats_tree.selection()[0]
        file_name = self.stats_tree.item(item)['values'][0]
        pos_file_path = os.path.join(self.output_path_var.get(), file_name)
        if os.path.exists(pos_file_path):
            self.view_pos_file(pos_file_path)

    def update_statistics_tree(self):
        """Update the statistics tree with processing results"""
        self.stats_tree.delete(*self.stats_tree.get_children())
        
        output_dir = self.output_path_var.get()
        
        if not os.path.exists(output_dir):
            return
            
        for file in os.listdir(output_dir):
            if file.endswith('.pos'):
                file_path = os.path.join(output_dir, file)
                timestamp = os.path.getmtime(file_path)
                formatted_time = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
                
                # Insert into tree with just filename and timestamp
                self.stats_tree.insert('', 'end', values=(file, formatted_time))

    def load_statistics(self):
        """Load statistics showing only .pos filenames"""
        try:
            if not self.stats_file.exists():
                return
                
            with open(self.stats_file, 'r') as f:
                self.stats_data = json.load(f)
            
            self.stats_tree.delete(*self.stats_tree.get_children())
            
            # Only show filename in tree
            for pos_filename, data in self.stats_data.items():
                filepath = Path(data['filepath'])
                if filepath.exists():
                    self.stats_tree.insert('', 'end',
                                         values=(pos_filename,),  # Only filename
                                         tags=(data['filepath'],))
                    
        except Exception as e:
            print(f"Error loading statistics: {e}")
            self.append_log(f"Error loading statistics: {e}\n")

    def save_statistics(self):
        """Save minimal statistics data"""
        try:
            with open(self.stats_file, 'w') as f:
                json.dump(self.stats_data, f, indent=4)
        except Exception as e:
            print(f"Error saving statistics: {e}")
            self.append_log(f"Error saving statistics: {e}\n")

    def on_tree_select(self, event):
        """Handle click on statistics tree item"""
        try:
            selection = self.stats_tree.selection()
            if not selection:
                return

            item = selection[0]
            filename = self.stats_tree.item(item)['values'][0]  # Get filename
            output_dir = self.output_path_var.get()
            file_path = os.path.join(output_dir, filename)
            
            if os.path.exists(file_path):
                self.view_pos_file(file_path)
            else:
                messagebox.showerror("Error", f"Could not find file: {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def create_statistics_frame(self):
        """Create simplified statistics frame"""
        stats_frame = ttk.LabelFrame(self.right_frame, text="Processing Statistics")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Create button frame
        button_frame = ttk.Frame(stats_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        # Add Show Latest button
        self.show_latest_btn = ttk.Button(
            button_frame, 
            text="Show Latest Result",
            command=self.show_latest_result
        )
        self.show_latest_btn.pack(side=tk.LEFT, padx=5)

        # Create Treeview
        self.stats_tree = ttk.Treeview(
            stats_frame, 
            columns=('File', 'Processed Time'),
            show='headings'
        )
        
        # Configure columns
        self.stats_tree.heading('File', text='File')
        self.stats_tree.heading('Processed Time', text='Processed Time')
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(
            stats_frame, 
            orient=tk.VERTICAL,
            command=self.stats_tree.yview
        )
        self.stats_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack elements
        self.stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind double-click
        self.stats_tree.bind('<Double-1>', self.on_tree_double_click)
        
        # Track latest file
        self.latest_pos_file = None

    def update_statistics_tree(self):
        """Update statistics tree and track latest file"""
        try:
            output_dir = self.output_path_var.get()
            if not os.path.exists(output_dir):
                return
                
            # Find newest .pos file
            pos_files = []
            for file in os.listdir(output_dir):
                if file.endswith('.pos'):
                    file_path = os.path.join(output_dir, file)
                    pos_files.append((file_path, os.path.getmtime(file_path)))
                    
            if pos_files:
                # Sort by modification time
                pos_files.sort(key=lambda x: x[1], reverse=True)
                self.latest_pos_file = pos_files[0][0]
                
                # Update tree
                self.stats_tree.delete(*self.stats_tree.get_children())
                for file_path, mtime in pos_files:
                    filename = os.path.basename(file_path)
                    formatted_time = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                    self.stats_tree.insert('', 'end', values=(filename, formatted_time))
                    
        except Exception as e:
            self.append_log(f"Error updating statistics: {e}\n")

    def show_latest_result(self):
        """Show most recently processed .pos file"""
        if self.latest_pos_file and os.path.exists(self.latest_pos_file):
            self.view_pos_file(self.latest_pos_file)
        else:
            messagebox.showinfo("No Results", "No processed files found.")

    def on_output_directory_change(self, new_output_dir):
        """Handle output directory change"""
        self.output_path_var.set(new_output_dir)
        self.save_config()

    def load_config(self):
        """Load configuration from file"""
        if os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)
                self.output_path_var.set(config.get('output_directory', ''))

    def save_config(self):
        """Save configuration to file"""
        config = {
            'output_directory': self.output_path_var.get()
        }
        with open(self.CONFIG_FILE, 'w') as f:
            json.dump(config, f)

    def import_sum_file(self):
        """Permet d'importer plusieurs fichiers .sum et de remplir les coordonnées de base."""
        sum_file_paths = filedialog.askopenfilenames(
            title="Sélectionnez des fichiers .sum",
            filetypes=[("Fichiers SUM", "*.sum"), ("Tous les fichiers", "*.*")]
        )
        
        if sum_file_paths:
            for sum_file_path in sum_file_paths:
                # Vérifier si le fichier n'est pas déjà dans la liste
                if sum_file_path not in self.sum_files:
                    try:
                        # Utiliser la méthode pour analyser le fichier .sum
                        parsed_data = self.parse_sum_file(sum_file_path)

                        # Extraire la date du nom du fichier .sum
                        sum_filename = Path(sum_file_path).stem
                        sum_date = self.extract_date_from_filename(sum_filename)
                        formatted_date = self.format_date(sum_date) if sum_date else "Date inconnue"

                        # Vérifier la correspondance avec les fichiers de base
                        matching_base = self.find_matching_base_file(sum_date)
                        if matching_base:
                            self.append_log(f"Fichier .sum {sum_filename} correspond au fichier de base {matching_base.name}\n")
                            status = "✓"
                        else:
                            self.append_log(f"ATTENTION: Aucun fichier de base correspondant trouvé pour {sum_filename}\n")
                            status = "❌"

                        # Ajouter à la liste des fichiers .sum importés
                        self.sum_files.append(sum_file_path)
                        
                        # Afficher dans la Listbox avec plus d'informations
                        display_text = f"{status} {Path(sum_file_path).name} ({formatted_date})"
                        if matching_base:
                            display_text += f" → {matching_base.name}"
                        
                        # S'assurer que la Listbox est en état normal avant d'insérer
                        self.sum_files_listbox.config(state='normal')
                        self.sum_files_listbox.insert(tk.END, display_text)
                        
                        # Mettre à jour les coordonnées avec les données du dernier fichier importé
                        self.base_lat_var.set(parsed_data.get("Latitude (DD)", ""))
                        self.base_lon_var.set(parsed_data.get("Longitude (DD)", ""))
                        self.base_height_var.set(parsed_data.get("Elevation (m)", ""))

                        self.append_log(f"Coordonnées importées depuis: {sum_file_path}\n")
                    except Exception as e:
                        self.append_log(f"Impossible de lire le fichier .sum {sum_file_path}:\n{e}\n")
                        messagebox.showerror("Erreur", f"Impossible de lire le fichier .sum:\n{sum_file_path}\n{e}")
                else:
                    self.append_log(f"Le fichier {Path(sum_file_path).name} est déjà importé\n")

            # Forcer la mise à jour de l'affichage
            self.sum_files_listbox.update()
            self.master.update_idletasks()

        self.unsaved_changes = True

    def find_matching_rover_file(self, sum_date):
        """Trouve le fichier rover correspondant à la date du fichier .sum."""
        if not sum_date:
            return None
            
        for rover in self.rover_obs_list:
            if rover.date == sum_date:
                return rover
        return None

    def format_date(self, date_str):
        """Formate la date YYYYMMDD en format lisible."""
        try:
            if date_str and len(date_str) == 8:
                year = date_str[:4]
                month = date_str[4:6]
                day = date_str[6:8]
                return f"{year}-{month}-{day}"
            return date_str
        except Exception:
            return date_str

    def extract_date_from_filename(self, filename):
        """Extrait la date du nom du fichier."""
        try:
            # Supposons que la date est au format YYYYMMDD dans le nom du fichier
            date_match = re.search(r'(\d{8})', filename)
            if date_match:
                return date_match.group(1)
            return None
        except Exception:
            return None

    def find_matching_base_file(self, sum_date):
        """Trouve le fichier de base correspondant à la date du fichier .sum."""
        if not sum_date:
            return None
            
        for base in self.base_obs_list:
            if base.date == sum_date:
                return base.filepath
        return None

    def update_config_with_base_coordinates(self, lat, lon, height, config_file):
        """Met à jour le fichier de configuration avec les nouvelles coordonnées."""
        try:
            # Lire le contenu du fichier
            with open(config_file, 'r') as f:
                lines = f.readlines()

            # Mettre à jour les coordonnées et la configuration de l'antenne
            updated_lines = []
            for line in lines:
                if line.startswith('ant1-pos1'):
                    updated_lines.append(f'ant1-pos1          ={lat}       # (deg)\n')
                elif line.startswith('ant1-pos2'):
                    updated_lines.append(f'ant1-pos2          ={lon}       # (deg)\n')
                elif line.startswith('ant1-pos3'):
                    updated_lines.append(f'ant1-pos3          ={height}       # (m)\n')
                elif line.startswith('ant2-pos1'):
                    updated_lines.append(f'ant2-pos1          ={lat}  # (deg|m)\n')
                elif line.startswith('ant2-pos2'):
                    updated_lines.append(f'ant2-pos2          ={lon}  # (deg|m)\n')
                elif line.startswith('ant2-pos3'):
                    updated_lines.append(f'ant2-pos3          ={height}  # (m|m)\n')
                else:
                    updated_lines.append(line)

            # Écrire le fichier mis à jour
            with open(config_file, 'w') as f:
                f.writelines(updated_lines)

            self.append_log(f"Fichier de configuration mis à jour avec succès:\n")
            self.append_log(f"Latitude: {lat}\n")
            self.append_log(f"Longitude: {lon}\n")
            self.append_log(f"Hauteur: {height}\n")

        except Exception as e:
            raise ValueError(f"Erreur lors de la mise à jour du fichier de configuration: {str(e)}")

    def apply_antenna_config(self):
        """Applique la configuration de l'antenne."""
        try:
            # Utiliser config_path_var au lieu de config_file_var
            config_file = self.config_path_var.get()
            if not config_file:
                raise ValueError("Aucun fichier de configuration sélectionné")

            # Récupérer l'offset
            offset = self.antenna_offset_var.get()
            if not offset:
                offset = "0"

            # Lire le fichier de configuration
            with open(config_file, 'r') as f:
                lines = f.readlines()

            # Mettre à jour la configuration
            updated_lines = []
            for line in lines:
                if line.startswith('ant1-antdelu'):
                    updated_lines.append(f'ant1-antdelu       ={offset}          # (m)\n')
                elif line.startswith('ant2-antdelu'):
                    updated_lines.append(f'ant2-antdelu       ={offset}          # (m)\n')
                else:
                    updated_lines.append(line)

            # Écrire le fichier mis à jour
            with open(config_file, 'w') as f:
                f.writelines(updated_lines)

            self.append_log(f"\nConfiguration antenne mise à jour:\n")
            self.append_log(f"Offset vertical: {offset} m\n")

            # Afficher un message de confirmation
            messagebox.showinfo("Succès", "Configuration de l'antenne mise à jour avec succès")

        except Exception as e:
            self.append_log(f"Erreur lors de l'application de la configuration antenne: {str(e)}\n")
            messagebox.showerror("Erreur", str(e))

        self.unsaved_changes = True

    def delete_sum_files(self):
        """Supprime les fichiers .sum sélectionnés de la liste."""
        selected_indices = list(self.sum_files_listbox.curselection())
        if not selected_indices:
            return

        # Supprimer les fichiers sélectionnés
        for index in reversed(selected_indices):
            del self.sum_files[index]
            self.sum_files_listbox.delete(index)

        self.append_log("Fichiers .sum supprimés\n")

        self.unsaved_changes = True

    def parse_sum_file(self, file_path):
        """Analyse le fichier .sum pour extraire les coordonnées et autres informations."""
        try:
            parsed_data = {
                "Date (UTC)": None,
                "Latitude (DD)": None,
                "Longitude (DD)": None,
                "Elevation (m)": None
            }

            with open(file_path, "r") as file:
                content = file.readlines()

            for line in content:
                if "POS LAT IGS20" in line:
                    parts = line.split()
                    try:
                        lat_deg = float(parts[7])
                        lat_min = float(parts[8])
                        lat_sec = float(parts[9])
                        lat_dd = lat_deg + lat_min/60 + lat_sec/3600
                        parsed_data["Latitude (DD)"] = f"{lat_dd:.9f}"
                    except (IndexError, ValueError) as e:
                        self.append_log(f"Erreur lors de l'extraction de la latitude: {str(e)}\n")

                elif "POS LON IGS20" in line:
                    parts = line.split()
                    try:
                        lon_deg = float(parts[7])
                        lon_min = float(parts[8])
                        lon_sec = float(parts[9])
                        lon_dd = lon_deg - (lon_min/60 + lon_sec/3600) if lon_deg < 0 else lon_deg + lon_min/60 + lon_sec/3600
                        parsed_data["Longitude (DD)"] = f"{lon_dd:.9f}"
                    except (IndexError, ValueError) as e:
                        self.append_log(f"Erreur lors de l'extraction de la longitude: {str(e)}\n")

                elif "POS HGT IGS20" in line:
                    parts = line.split()
                    try:
                        # Ne garder que la valeur de la hauteur sans le sigma
                        parsed_data["Elevation (m)"] = parts[5]
                    except (IndexError, ValueError) as e:
                        self.append_log(f"Erreur lors de l'extraction de la hauteur: {str(e)}\n")

            # Logs pour le débogage
            self.append_log(f"Données extraites du fichier .sum:\n")
            self.append_log(f"Latitude: {parsed_data['Latitude (DD)']}\n")
            self.append_log(f"Longitude: {parsed_data['Longitude (DD)']}\n")
            self.append_log(f"Hauteur: {parsed_data['Elevation (m)']}\n")

            return parsed_data

        except Exception as e:
            raise ValueError(f"Erreur lors de l'analyse du fichier .sum: {e}")

    def update_coord_mode(self):
        """Met à jour l'interface selon le mode sélectionné"""
        mode = self.coord_mode.get()
        if mode == "manual":
            # Mode manuel : activer les champs de saisie, désactiver les contrôles .sum
            self.lat_entry.config(state='normal')
            self.lon_entry.config(state='normal')
            self.height_entry.config(state='normal')
            self.import_sum_button.config(state='disabled')
            self.sum_files_listbox.config(state='disabled')
            self.delete_sum_button.config(state='disabled')
            self.sum_list_label.config(state='disabled')
        else:
            # Mode auto : désactiver les champs de saisie, activer les contrôles .sum
            self.lat_entry.config(state='readonly')
            self.lon_entry.config(state='readonly')
            self.height_entry.config(state='readonly')
            self.import_sum_button.config(state='normal')
            self.sum_files_listbox.config(state='normal')
            self.delete_sum_button.config(state='normal')
            self.sum_list_label.config(state='normal')
            # Rafraîchir l'affichage des fichiers
            self.refresh_sum_files_display()

    def refresh_sum_files_display(self):
        """Rafraîchit l'affichage des fichiers .sum dans la Listbox."""
        try:
            # Vider la Listbox
            self.sum_files_listbox.delete(0, tk.END)
            
            # Réafficher tous les fichiers
            for sum_file_path in self.sum_files:
                sum_filename = Path(sum_file_path).stem
                sum_date = self.extract_date_from_filename(sum_filename)
                formatted_date = self.format_date(sum_date) if sum_date else "Date inconnue"
                
                # Vérifier la correspondance
                matching_base = self.find_matching_base_file(sum_date)
                status = "✓" if matching_base else "❌"
                
                # Créer le texte d'affichage
                display_text = f"{status} {Path(sum_file_path).name} ({formatted_date})"
                if matching_base:
                    display_text += f" → {matching_base.name}"
                
                self.sum_files_listbox.insert(tk.END, display_text)
            
            # Forcer la mise à jour
            self.sum_files_listbox.update()
            self.master.update_idletasks()
        except Exception as e:
            self.append_log(f"Erreur lors du rafraîchissement de l'affichage: {str(e)}\n")

    def on_close(self):
        if self.unsaved_changes:
            response = messagebox.askyesnocancel("Quitter", "Des modifications non sauvegardées existent.\nVoulez-vous sauvegarder avant de quitter ?")
            if response:  # Oui
                self.save_project()
                self.master.destroy()
            elif response is None:  # Annuler
                return
            else:  # Non
                self.master.destroy()
        else:
            self.master.destroy()

    def on_entry_change(self, *args):
        self.unsaved_changes = True

def main():
    root = tk.Tk()
    app = PPKProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()