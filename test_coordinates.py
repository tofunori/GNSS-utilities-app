import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
from pathlib import Path

class TestCoordinatesGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Test Coordonnées Base")
        
        # Variables pour la gestion des fichiers de configuration
        self.original_config_path = None
        self.config_path_var = tk.StringVar()
        self.temp_counter = 0

        # Variables pour les coordonnées
        self.base_lat_var = tk.StringVar()
        self.base_lon_var = tk.StringVar()
        self.base_height_var = tk.StringVar()
        self.coord_mode_var = tk.StringVar(value="manual")

        # Créer l'interface
        self.create_gui()

        # Configurer la fermeture propre
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_gui(self):
        # Frame principal
        main_frame = ttk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Config file selection
        config_frame = ttk.LabelFrame(main_frame, text="Fichier de configuration")
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Entry(config_frame, textvariable=self.config_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Button(config_frame, text="Parcourir", command=self.browse_config).pack(side=tk.RIGHT, padx=5, pady=5)

        # Base files selection
        base_frame = ttk.LabelFrame(main_frame, text="Fichiers de base")
        base_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Button(base_frame, text="Ajouter fichiers", command=self.add_base_files).pack(anchor='w', padx=5, pady=5)
        
        # Listbox pour les fichiers de base
        self.base_listbox = tk.Listbox(base_frame, selectmode=tk.MULTIPLE, height=5)
        self.base_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Base Coordinates Frame
        base_coord_frame = ttk.LabelFrame(main_frame, text="Coordonnées de la Base")
        base_coord_frame.pack(fill=tk.X, padx=5, pady=5)

        # Mode selection
        mode_frame = ttk.Frame(base_coord_frame)
        mode_frame.grid(row=0, column=0, columnspan=2, pady=5)
        
        ttk.Radiobutton(mode_frame, text="Mode Manuel", variable=self.coord_mode_var, 
                       value="manual", command=self.toggle_coordinate_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Mode Automatique", variable=self.coord_mode_var, 
                       value="auto", command=self.toggle_coordinate_mode).pack(side=tk.LEFT, padx=5)

        # Manual coordinates frame
        self.manual_coord_frame = ttk.Frame(base_coord_frame)
        self.manual_coord_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')

        # Latitude
        ttk.Label(self.manual_coord_frame, text="Latitude (deg):").grid(row=0, column=0, padx=5, pady=2, sticky='e')
        self.lat_entry = ttk.Entry(self.manual_coord_frame, textvariable=self.base_lat_var)
        self.lat_entry.grid(row=0, column=1, padx=5, pady=2, sticky='ew')

        # Longitude
        ttk.Label(self.manual_coord_frame, text="Longitude (deg):").grid(row=1, column=0, padx=5, pady=2, sticky='e')
        self.lon_entry = ttk.Entry(self.manual_coord_frame, textvariable=self.base_lon_var)
        self.lon_entry.grid(row=1, column=1, padx=5, pady=2, sticky='ew')

        # Height
        ttk.Label(self.manual_coord_frame, text="Hauteur (m):").grid(row=2, column=0, padx=5, pady=2, sticky='e')
        self.height_entry = ttk.Entry(self.manual_coord_frame, textvariable=self.base_height_var)
        self.height_entry.grid(row=2, column=1, padx=5, pady=2, sticky='ew')

        # Auto coordinates frame
        self.auto_coord_frame = ttk.Frame(base_coord_frame)
        self.auto_coord_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')
        
        ttk.Label(self.auto_coord_frame, text="Coordonnées extraites automatiquement du fichier .sum").pack(pady=10)
        self.auto_coords_label = ttk.Label(self.auto_coord_frame, text="En attente des fichiers...")
        self.auto_coords_label.pack(pady=5)

        # Add Apply button
        ttk.Button(base_coord_frame, text="Appliquer", command=self.apply_base_coordinates).grid(row=2, column=0, columnspan=2, pady=5)

        # Initially hide auto frame
        self.auto_coord_frame.grid_remove()

        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def append_log(self, message):
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def browse_config(self):
        """Sélectionner le fichier de configuration"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Config files", "*.conf"), ("All files", "*.*")]
        )
        if file_path:
            self.original_config_path = file_path
            self.config_path_var.set(file_path)
            self.append_log(f"Fichier de configuration sélectionné: {file_path}\n")

    def add_base_files(self):
        """Ajouter des fichiers de base"""
        paths = filedialog.askopenfilenames(
            title="Sélectionner les fichiers de base",
            filetypes=[("Base Files", "*.22O *.23O *.24O *.sum"), ("All files", "*.*")]
        )
        for path in paths:
            if path not in [self.base_listbox.get(i) for i in range(self.base_listbox.size())]:
                self.base_listbox.insert(tk.END, path)
                self.append_log(f"Fichier ajouté: {path}\n")
        
        if self.coord_mode_var.get() == "auto":
            self.update_auto_coordinates()

    def toggle_coordinate_mode(self):
        """Bascule entre les modes manuel et automatique"""
        if self.coord_mode_var.get() == "manual":
            self.auto_coord_frame.grid_remove()
            self.manual_coord_frame.grid()
        else:
            self.manual_coord_frame.grid_remove()
            self.auto_coord_frame.grid()
            self.update_auto_coordinates()

    def parse_sum_file(self, file_path):
        """Parse le fichier .sum pour extraire les coordonnées"""
        try:
            with open(file_path, "r") as file:
                content = file.readlines()

            lat = lon = height = None
            for line in content:
                if "POS LAT" in line:
                    parts = line.split()
                    lat = float(parts[7]) + float(parts[8])/60 + float(parts[9].strip('"'))/3600
                elif "POS LON" in line:
                    parts = line.split()
                    lon = float(parts[7]) + float(parts[8])/60 + float(parts[9].strip('"'))/3600
                elif "POS HGT" in line:
                    parts = line.split()
                    height = float(parts[5])

            return lat, lon, height
        except Exception as e:
            self.append_log(f"Erreur lors du parsing du fichier .sum: {str(e)}\n")
            return None, None, None

    def update_auto_coordinates(self):
        """Met à jour les coordonnées automatiquement"""
        try:
            base_files = [self.base_listbox.get(i) for i in range(self.base_listbox.size())]
            if not base_files:
                self.auto_coords_label.config(text="Aucun fichier de base sélectionné")
                return

            sum_files = [f for f in base_files if f.endswith('.sum')]
            if not sum_files:
                self.auto_coords_label.config(text="Aucun fichier .sum trouvé")
                return

            lat, lon, height = self.parse_sum_file(sum_files[0])
            
            if lat is not None and lon is not None and height is not None:
                self.base_lat_var.set(f"{lat:.8f}")
                self.base_lon_var.set(f"{lon:.8f}")
                self.base_height_var.set(f"{height:.3f}")
                
                self.auto_coords_label.config(
                    text=f"Coordonnées extraites:\nLat: {lat:.8f}°\nLon: {lon:.8f}°\nH: {height:.3f}m"
                )
                
                self.apply_base_coordinates(show_messages=False)
            else:
                self.auto_coords_label.config(text="Coordonnées non trouvées dans le fichier .sum")

        except Exception as e:
            self.auto_coords_label.config(text=f"Erreur: {str(e)}")
            self.append_log(f"Erreur lors de l'extraction des coordonnées: {str(e)}\n")

    def create_temp_config(self):
        """Créer une copie temporaire du fichier de configuration"""
        try:
            config_path = self.config_path_var.get()
            if not config_path:
                return None
                
            config_dir = os.path.dirname(config_path)
            config_name = os.path.basename(config_path)
            self.temp_counter += 1
            temp_config_path = os.path.join(config_dir, f"temp_{self.temp_counter}_{config_name}")
            
            shutil.copy2(config_path, temp_config_path)
            return temp_config_path
        except Exception as e:
            self.append_log(f"Erreur lors de la création du fichier temporaire: {str(e)}\n")
            return None

    def cleanup_temp_files(self):
        """Nettoyer les fichiers temporaires"""
        try:
            if self.original_config_path:
                config_dir = os.path.dirname(self.original_config_path)
                config_name = os.path.basename(self.original_config_path)
                for file in os.listdir(config_dir):
                    if file.startswith("temp_") and file.endswith(config_name):
                        try:
                            os.remove(os.path.join(config_dir, file))
                        except:
                            pass
        except:
            pass

    def apply_base_coordinates(self, show_messages=True):
        """Applique les coordonnées de la base"""
        try:
            lat = float(self.base_lat_var.get())
            lon = float(self.base_lon_var.get())
            height = float(self.base_height_var.get())

            temp_config_path = self.create_temp_config()
            if not temp_config_path:
                if show_messages:
                    messagebox.showerror("Erreur", "Impossible de créer le fichier temporaire.")
                return

            with open(temp_config_path, 'r') as f:
                config_lines = f.readlines()

            for i, line in enumerate(config_lines):
                if line.startswith('ant2-postype'):
                    config_lines[i] = 'ant2-postype       =llh        # (0:llh,1:xyz,2:single,3:posfile,4:rinexhead,5:rtcm)\n'
                elif line.startswith('ant2-pos1'):
                    config_lines[i] = f'ant2-pos1          ={lat:.8f}  # (deg|m)\n'
                elif line.startswith('ant2-pos2'):
                    config_lines[i] = f'ant2-pos2          ={lon:.8f}  # (deg|m)\n'
                elif line.startswith('ant2-pos3'):
                    config_lines[i] = f'ant2-pos3          ={height:.3f}  # (m|m)\n'

            with open(temp_config_path, 'w') as f:
                f.writelines(config_lines)

            self.config_path_var.set(temp_config_path)
            self.append_log(f"Coordonnées mises à jour dans {temp_config_path}\n")
            
            if show_messages:
                messagebox.showinfo("Succès", "Coordonnées mises à jour dans le fichier temporaire.")

        except ValueError:
            if show_messages:
                messagebox.showerror("Erreur", "Veuillez entrer des coordonnées valides.")
        except Exception as e:
            if show_messages:
                messagebox.showerror("Erreur", f"Erreur lors de la mise à jour: {str(e)}")
            self.append_log(f"Erreur: {str(e)}\n")

    def on_closing(self):
        """Gérer la fermeture de l'application"""
        self.cleanup_temp_files()
        self.master.destroy()

def main():
    root = tk.Tk()
    app = TestCoordinatesGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 