import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import shutil
import re
import matplotlib.pyplot as plt
from datetime import datetime
import matplotlib.dates as mdates
import json


class GNSSViewer:
    def __init__(self, window):
        self.window = window
        self.setup_ui()

    def setup_ui(self):
        """Configure l'interface utilisateur"""
        # Configuration de la fenêtre principale
        self.window.title("GNSS Data Viewer")
        
        # Création du menu en premier
        self.create_menu()
        
        # Création des widgets
        self.create_data_entry()
        self.create_button_bar()
        self.create_main_container()
        self.create_table()
        self.create_stats_panel()

        # Ajouter les gestionnaires d'événements pour le focus
        self.window.bind("<Button-1>", self.on_window_click)
        self.window.bind("<FocusIn>", self.on_window_focus)

    def on_window_click(self, event):
        """Met la fenêtre au premier plan lors d'un clic"""
        self.window.attributes('-topmost', True)
        self.window.attributes('-topmost', False)

    def on_window_focus(self, event):
        """Met la fenêtre au premier plan lors de la prise de focus"""
        self.window.attributes('-topmost', True)
        self.window.attributes('-topmost', False)

    def create_main_container(self):
        """Crée le conteneur principal pour la table et les stats"""
        self.main_container = ttk.Frame(self.window)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def create_stats_panel(self):
        """Crée le panneau de statistiques"""
        self.stats_frame = ttk.LabelFrame(self.main_container, text="Statistics", padding=10)
        self.stats_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5,0), pady=0)

        # Bouton de calcul
        ttk.Button(
            self.stats_frame, 
            text="Calculate Mean", 
            command=self.calculate_mean
        ).pack(fill=tk.X, padx=5, pady=5)

        # Zone de résultats
        ttk.Label(
            self.stats_frame, 
            text="Results:"
        ).pack(fill=tk.X, padx=5, pady=(10,0))

        self.stats_text = tk.Text(
            self.stats_frame, 
            width=30, 
            height=10
        )
        self.stats_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def calculate_mean(self):
        """Calcule les statistiques des coordonnées"""
        if not self.tree.get_children():
            messagebox.showwarning("Warning", "No data in table.")
            return

        try:
            # Récupération des indices des colonnes
            lat_index = self.fields.index("Latitude (DD)")
            lon_index = self.fields.index("Longitude (DD)")
            alt_index = self.fields.index("Elevation (m)")

            # Listes pour stocker les valeurs
            lats = []
            lons = []
            alts = []

            # Collecter toutes les valeurs
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                lat = float(values[lat_index])
                lon = float(values[lon_index])
                alt = float(values[alt_index])
                
                lats.append(lat)
                lons.append(lon)
                alts.append(alt)

            # Calculer les statistiques
            count = len(lats)
            stats = {
                'Latitude': {
                    'Mean': sum(lats) / count,
                    'Min': min(lats),
                    'Max': max(lats),
                    'Range': max(lats) - min(lats)
                },
                'Longitude': {
                    'Mean': sum(lons) / count,
                    'Min': min(lons),
                    'Max': max(lons),
                    'Range': max(lons) - min(lons)
                },
                'Elevation': {
                    'Mean': sum(alts) / count,
                    'Min': min(alts),
                    'Max': max(alts),
                    'Range': max(alts) - min(alts)
                }
            }

            # Créer une nouvelle fenêtre pour les résultats
            result_window = tk.Toplevel()
            result_window.title("Coordinate Statistics")
            result_window.geometry("500x400")
            
            # Désactiver temporairement l'attribut topmost de la fenêtre principale
            self.window.attributes('-topmost', False)
            # Configurer la fenêtre des résultats
            result_window.attributes('-topmost', True)
            result_window.focus_force()

            # Zone de texte pour les résultats
            text = tk.Text(result_window, wrap=tk.WORD, height=20, width=60)
            text.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

            # Créer les tags pour le texte en gras
            text.tag_configure("bold", font=("TkDefaultFont", 9, "bold"))

            # Formater et afficher les résultats
            result = f"Statistics from {count} points:\n\n"
            text.insert(tk.END, result)
            
            for coord in ['Latitude', 'Longitude', 'Elevation']:
                text.insert(tk.END, f"{coord}:\n")
                text.insert(tk.END, "  Mean: ", "bold")
                text.insert(tk.END, f"{stats[coord]['Mean']:.9f}°\n" if coord != 'Elevation' else f"{stats[coord]['Mean']:.4f} m\n")
                text.insert(tk.END, f"  Min:  {stats[coord]['Min']:.9f}°\n" if coord != 'Elevation' else f"  Min:  {stats[coord]['Min']:.4f} m\n")
                text.insert(tk.END, f"  Max:  {stats[coord]['Max']:.9f}°\n" if coord != 'Elevation' else f"  Max:  {stats[coord]['Max']:.4f} m\n")
                text.insert(tk.END, f"  Range:{stats[coord]['Range']:.9f}°\n" if coord != 'Elevation' else f"  Range:{stats[coord]['Range']:.4f} m\n")
                text.insert(tk.END, "\n")

            text.configure(state='normal')

            # Frame pour les boutons
            button_frame = ttk.Frame(result_window)
            button_frame.pack(pady=5)

            # Bouton pour copier
            copy_button = ttk.Button(
                button_frame,
                text="Copy All",
                command=lambda: self.copy_to_clipboard(text.get("1.0", tk.END))
            )
            copy_button.pack(side=tk.LEFT, padx=5)

            # Bouton pour sauvegarder
            save_button = ttk.Button(
                button_frame,
                text="Save As",
                command=lambda: self.save_statistics(text.get("1.0", tk.END))
            )
            save_button.pack(side=tk.LEFT, padx=5)

            # Gérer la fermeture de la fenêtre
            def on_closing():
                self.window.attributes('-topmost', True)  # Réactiver topmost pour la fenêtre principale
                result_window.destroy()

            result_window.protocol("WM_DELETE_WINDOW", on_closing)

        except Exception as e:
            messagebox.showerror("Error", f"Error calculating statistics: {str(e)}")

    def save_statistics(self, content):
        """Sauvegarde les statistiques dans un fichier texte"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(content)
                messagebox.showinfo("Success", "Statistics saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving statistics: {str(e)}")

    def copy_to_clipboard(self, text):
        """Copie le texte dans le presse-papiers"""
        self.window.clipboard_clear()
        self.window.clipboard_append(text)
        messagebox.showinfo("Success", "Results copied to clipboard!")

    def create_data_entry(self):
        """Crée la zone de saisie des données"""
        self.fields = [
            "Date (UTC)", "File", "GNSS Model", "Description of Occupation",
            "Time Start (UTC)", "Time End (UTC)", "Duration", "Interval",
            "Latitude (DD)", "Longitude (DD)", "UTM N (m)", "UTM E (m)", "Elevation (m)",
            "Reference Point", "Sigma UTM N (m)", "Sigma UTM E (m)", "Sigma Elev. (m)", 
            "Datum", "Solution"
        ]
        
        self.field_entries = {}
        
        # Frame for Data Entry
        form_frame = ttk.LabelFrame(self.window, text="Data Entry", padding=10)
        form_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)

        # Create form fields
        for i, field in enumerate(self.fields):
            ttk.Label(form_frame, text=field).grid(row=i, column=0, padx=5, pady=5)
            if field == "GNSS Model":
                entry = ttk.Combobox(form_frame, values=["N/A", "EMLID INREACH RS2", "FOIF A30"], state="readonly")
                entry.set("N/A")
                entry.bind('<<ComboboxSelected>>', self.on_gnss_model_change)
            elif field == "Description of Occupation":
                entry = ttk.Combobox(form_frame, values=["Base", "Rover"], state="readonly")
                entry.set("Base")
                entry.bind('<<ComboboxSelected>>', self.on_occupation_change)
            elif field == "Reference Point":
                entry = ttk.Combobox(form_frame, values=["APC"], state="readonly")
                entry.set("APC")
                entry.bind('<<ComboboxSelected>>', self.on_reference_point_change)
            elif field == "Solution":
                entry = ttk.Entry(form_frame, width=30)
                entry.insert(0, "PPP")
            elif field == "Datum":
                entry = ttk.Entry(form_frame, width=30)
                entry.insert(0, "ITRF20")
            else:
                entry = ttk.Entry(form_frame, width=30)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.field_entries[field] = entry

            # Add Convert DMS buttons
            if field in ["Latitude (DD)", "Longitude (DD)"]:
                ttk.Button(
                    form_frame, 
                    text="Convert DMS",
                    command=lambda f=field: self.convert_dms_to_dd(f)
                ).grid(row=i, column=2, padx=5, pady=5)

    def create_button_bar(self):
        """Crée la barre de boutons"""
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(button_frame, text="Import .sum File", command=self.load_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Add to Table", command=self.add_to_treeview).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Entry", command=self.delete_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Fields", command=self.clear_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save to Excel", command=self.save_to_excel).pack(side=tk.LEFT, padx=5)

    def create_table(self):
        """Crée la table principale"""
        table_frame = ttk.LabelFrame(self.main_container, text="GNSS Data Table", padding=10)
        table_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5), pady=0)

        # Create Treeview
        self.tree = ttk.Treeview(table_frame, columns=self.fields, show="headings", height=15)
        for col in self.fields:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        # Add scrollbars
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def on_gnss_model_change(self, event):
        """Gère le changement de modèle GNSS"""
        selected_model = self.field_entries["GNSS Model"].get()
        occupation_type = self.field_entries["Description of Occupation"].get()
        current_elevation = self.field_entries["Elevation (m)"].get()
        
        try:
            # Stocker l'élévation originale si elle n'existe pas déjà
            if not hasattr(self, 'original_elevation') and current_elevation:
                self.original_elevation = float(current_elevation)
            elif not current_elevation:
                return
            
            if selected_model == "EMLID INREACH RS2":
                if occupation_type == "Base":
                    # Pour Base: APC avec offset
                    self.field_entries["Reference Point"].set("APC")
                    if hasattr(self, 'original_elevation'):
                        offset = 0.136
                        adjusted_elevation = self.original_elevation + offset
                        self.field_entries["Elevation (m)"].delete(0, tk.END)
                        self.field_entries["Elevation (m)"].insert(0, f"{adjusted_elevation:.6f}")
                else:
                    # Pour Rover: ARP sans offset
                    self.field_entries["Reference Point"].set("ARP")
                    if hasattr(self, 'original_elevation'):
                        self.field_entries["Elevation (m)"].delete(0, tk.END)
                        self.field_entries["Elevation (m)"].insert(0, f"{self.original_elevation:.6f}")
                
                self.field_entries["Reference Point"]["state"] = "disabled"
            
            else:  # Pour N/A ou autres modèles
                # Restaurer l'élévation originale
                if hasattr(self, 'original_elevation'):
                    self.field_entries["Elevation (m)"].delete(0, tk.END)
                    self.field_entries["Elevation (m)"].insert(0, f"{self.original_elevation:.6f}")
                
                self.field_entries["Reference Point"].set("APC")
                self.field_entries["Reference Point"]["state"] = "disabled"

        except ValueError:
            pass

    def on_reference_point_change(self, event):
        """Gère le changement de point de référence"""
        ref_point = self.field_entries["Reference Point"].get()
        gnss_model = self.field_entries["GNSS Model"].get()
        
        try:
            current_elevation = float(self.field_entries["Elevation (m)"].get())
            if gnss_model == "EMLID INREACH RS2":
                offset = 0.136  # Offset total ARP à APC
                adjusted_elevation = current_elevation + offset
                    
                self.field_entries["Elevation (m)"].delete(0, tk.END)
                self.field_entries["Elevation (m)"].insert(0, f"{adjusted_elevation:.6f}")
            elif gnss_model == "N/A":
                # Ne rien faire si le modèle est N/A
                pass
        except ValueError:
            pass

    def load_file(self):
        """Charge un fichier .sum"""
        # Désactiver temporairement l'attribut topmost de la fenêtre principale
        self.window.attributes('-topmost', False)
        
        file_path = filedialog.askopenfilename(
            filetypes=[("SUM Files", "*.sum")],
            parent=self.window  # Définir la fenêtre parent
        )
        
        # Réactiver l'attribut topmost
        self.window.attributes('-topmost', True)
        
        if file_path:
            parsed_data = self.parse_sum_file(file_path)
            if parsed_data:
                # Stocker l'élévation originale du fichier .sum
                if parsed_data["Elevation (m)"]:
                    self.original_elevation = float(parsed_data["Elevation (m)"])

                # Mettre à jour tous les champs
                for field, entry in self.field_entries.items():
                    if isinstance(entry, ttk.Combobox):
                        if field == "Description of Occupation":
                            entry.set(parsed_data[field] or "Base")
                        elif field == "Reference Point":
                            entry.set(parsed_data[field] or "APC")
                    elif field in parsed_data and parsed_data[field] is not None:
                        entry.delete(0, tk.END)
                        entry.insert(0, parsed_data[field])

                # Appliquer l'offset SEULEMENT si EMLID INREACH RS2 est sélectionné
                current_model = self.field_entries["GNSS Model"].get()
                if current_model == "EMLID INREACH RS2":
                    offset = 0.136
                    adjusted_elevation = self.original_elevation + offset
                    self.field_entries["Elevation (m)"].delete(0, tk.END)
                    self.field_entries["Elevation (m)"].insert(0, f"{adjusted_elevation:.6f}")

    def add_to_treeview(self):
        """Ajoute les données du formulaire à la table"""
        # Vérifier si les coordonnées sont en format DMS
        lat = self.field_entries["Latitude (DD)"].get()
        lon = self.field_entries["Longitude (DD)"].get()
        
        if '°' in lat or '°' in lon:
            error_window = messagebox.showerror(
                "Error", 
                "Please convert coordinates to Decimal Degrees before adding to table.",
                parent=self.window
            )
            return
            
        # Vérifier si un modèle GNSS valide est sélectionné
        gnss_model = self.field_entries["GNSS Model"].get()
        if gnss_model == "N/A":
            error_window = messagebox.showerror(
                "Error", 
                "Please select a valid GNSS model (not N/A) before adding to table.",
                parent=self.window
            )
            return
            
        try:
            # Vérifier si les coordonnées sont des nombres valides
            float(lat)
            float(lon)
            
            # Si toutes les validations sont passées, ajouter à la table
            values = []
            for field in self.fields:
                value = self.field_entries[field].get()
                # Formater les coordonnées à 9 décimales
                if field in ["Latitude (DD)", "Longitude (DD)"]:
                    try:
                        value = f"{float(value):.9f}"
                    except ValueError:
                        pass
                values.append(value)
            self.tree.insert("", "end", values=values)
            
        except ValueError:
            error_window = messagebox.showerror(
                "Error", 
                "Invalid coordinate format. Please ensure coordinates are in Decimal Degrees.",
                parent=self.window
            )
            return

    def delete_entry(self):
        """Supprime les entrées sélectionnées de la table"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No entry selected to delete.")
            return
        for item in selected_items:
            self.tree.delete(item)

    def clear_fields(self):
        """Efface tous les champs du formulaire"""
        for field, entry in self.field_entries.items():
            if isinstance(entry, ttk.Combobox):
                if field == "Description of Occupation":
                    entry.set("Base")
                elif field == "GNSS Model":
                    entry.set("N/A")
                elif field == "Reference Point":
                    entry.set("APC")
                    entry.state(["disabled"])
            elif field == "Solution":
                entry.delete(0, tk.END)
                entry.insert(0, "PPP")
            elif field == "Datum":
                entry.delete(0, tk.END)
                entry.insert(0, "ITRF20")
            else:
                entry.delete(0, tk.END)

    def save_to_excel(self):
        """Sauvegarde les données de la table dans un fichier Excel"""
        if not self.tree.get_children():
            messagebox.showerror("Error", "No data to save!")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            data = []
            for item in self.tree.get_children():
                data.append(self.tree.item(item)['values'])
            
            df = pd.DataFrame(data, columns=self.fields)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Data saved to Excel successfully!")

    def convert_dms_to_dd(self, field_name):
        """Convertit les coordonnées DMS en degrés décimaux"""
        dms_value = self.field_entries[field_name].get()
        try:
            dd_value = dms_to_dd(dms_value)
            self.field_entries[field_name].delete(0, tk.END)
            # Limiter à 9 décimales
            self.field_entries[field_name].insert(0, f"{dd_value:.9f}")
        except Exception as e:
            messagebox.showerror("Error", f"Invalid DMS format: {str(e)}")

    def parse_sum_file(self, file_path):
        """Parse un fichier .sum et extrait les données"""
        parsed_data = {
            "Date (UTC)": None, "File": None, "GNSS Model": None,
            "Description of Occupation": None, "Time Start (UTC)": None,
            "Time End (UTC)": None, "Duration": None, "Interval": None,
            "Latitude (DD)": None, "Longitude (DD)": None, "UTM N (m)": None,
            "UTM E (m)": None, "Elevation (m)": None, "Reference Point": "APC",
            "Sigma UTM N (m)": None, "Sigma UTM E (m)": None,
            "Sigma Elev. (m)": None, "Datum": "ITRF20", "Solution": "PPP",
        }

        try:
            with open(file_path, "r") as file:
                content = file.readlines()

            for line in content:
                if line.startswith("MKR"):
                    parsed_data["Date (UTC)"] = line.split("MKR")[1].strip()
                elif line.startswith("RNX"):
                    parsed_data["File"] = line.split("RNX")[1].strip()
                elif line.startswith("BEG"):
                    date_time = line.split("BEG")[1].strip()
                    parsed_data["Date (UTC)"] = date_time.split()[0]
                    parsed_data["Time Start (UTC)"] = date_time
                elif line.startswith("END"):
                    parsed_data["Time End (UTC)"] = line.split("END")[1].strip()
                elif line.startswith("INT"):
                    parsed_data["Interval"] = line.split("INT")[1].strip()
                elif "POS LAT" in line:
                    parts = line.split()
                    parsed_data["Latitude (DD)"] = f"{parts[7]}° {parts[8]}' {parts[9]}\""
                    parsed_data["Sigma UTM N (m)"] = parts[11]
                elif "POS LON" in line:
                    parts = line.split()
                    parsed_data["Longitude (DD)"] = f"{parts[7]}° {parts[8]}' {parts[9]}\""
                    parsed_data["Sigma UTM E (m)"] = parts[11]
                elif "POS HGT" in line:
                    parts = line.split()
                    parsed_data["Elevation (m)"] = parts[5]
                    parsed_data["Sigma Elev. (m)"] = parts[7]
                elif "PRJ TYPE ZONE    EASTING     NORTHING" in line:
                    utm_data_line = content[content.index(line) + 1].strip()
                    utm_data_parts = utm_data_line.split()
                    parsed_data["UTM E (m)"] = utm_data_parts[3]
                    parsed_data["UTM N (m)"] = utm_data_parts[4]

            if parsed_data["Time Start (UTC)"] and parsed_data["Time End (UTC)"]:
                fmt = "%Y-%m-%d %H:%M:%S.%f"
                start_time = datetime.strptime(parsed_data["Time Start (UTC)"], fmt)
                end_time = datetime.strptime(parsed_data["Time End (UTC)"], fmt)
                duration_seconds = int((end_time - start_time).total_seconds())
                parsed_data["Duration"] = f"{duration_seconds // 3600:02}:{(duration_seconds % 3600) // 60:02}:{duration_seconds % 60:02}"

            return parsed_data
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing file: {e}")
            return None

    def create_menu(self):
        """Crée le menu de l'application"""
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project", command=self.open_project)
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.window.destroy)

        # Statistics Menu
        stats_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Statistics", menu=stats_menu)
        stats_menu.add_command(label="Calculate Mean", command=self.calculate_mean)

    def new_project(self):
        """Crée un nouveau projet en réinitialisant tous les champs"""
        # Demander confirmation
        if messagebox.askyesno("New Project", "Are you sure you want to start a new project? All unsaved data will be lost."):
            # Effacer tous les champs
            self.clear_fields()
            # Effacer la table
            for item in self.tree.get_children():
                self.tree.delete(item)
            # Réinitialiser l'élévation originale
            if hasattr(self, 'original_elevation'):
                delattr(self, 'original_elevation')

    def open_project(self):
        # Désactiver temporairement l'attribut topmost
        self.window.attributes('-topmost', False)
        
        file_path = filedialog.askopenfilename(
            defaultextension=".project",
            filetypes=[("Project files", "*.project"), ("All files", "*.*")],
            parent=self.window
        )
        
        # Réactiver l'attribut topmost
        self.window.attributes('-topmost', True)
        
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
                    # Charger les données dans les champs
                    for field, value in data['fields'].items():
                        if field in self.field_entries:
                            self.field_entries[field].delete(0, tk.END)
                            self.field_entries[field].insert(0, value)
                    
                    # Charger les données dans la table
                    self.tree.delete(*self.tree.get_children())
                    for row in data['table_data']:
                        self.tree.insert("", tk.END, values=row)
                    
                messagebox.showinfo("Success", "Project loaded successfully!", parent=self.window)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load project: {str(e)}", parent=self.window)

    def save_project(self):
        # Désactiver temporairement l'attribut topmost
        self.window.attributes('-topmost', False)
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".project",
            filetypes=[("Project files", "*.project"), ("All files", "*.*")],
            parent=self.window
        )
        
        # Réactiver l'attribut topmost
        self.window.attributes('-topmost', True)
        
        if file_path:
            try:
                # Collecter les données des champs
                field_data = {
                    field: entry.get() for field, entry in self.field_entries.items()
                }
                
                # Collecter les données de la table
                table_data = []
                for item in self.tree.get_children():
                    table_data.append(self.tree.item(item)['values'])
                
                # Créer le dictionnaire de données
                data = {
                    'fields': field_data,
                    'table_data': table_data
                }
                
                # Sauvegarder dans le fichier
                with open(file_path, 'w') as f:
                    json.dump(data, f, indent=4)
                    
                messagebox.showinfo("Success", "Project saved successfully!", parent=self.window)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save project: {str(e)}", parent=self.window)

    def on_occupation_change(self, event):
        """Gère le changement de type d'occupation"""
        # Déclencher le changement de modèle GNSS pour mettre à jour le point de référence
        self.on_gnss_model_change(None)


class PosToExcelConverter:
    def __init__(self, window):
        self.window = window
        self.window.title("POS to Excel Converter")
        self.window.geometry("800x600")
        
        # Input/Output Frame
        input_frame = ttk.LabelFrame(self.window, text="File Selection", padding=10)
        input_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(input_frame, text="Input Directory:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.input_dir = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_dir, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=5)

        ttk.Label(input_frame, text="Output File:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.output_file = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.output_file, width=40).grid(row=1, column=1, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=5)

        # Button frame
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        self.preview_button = ttk.Button(button_frame, text="Preview Data", command=self.preview_data)
        self.preview_button.pack(side=tk.LEFT, padx=5)
        
        self.save_button = ttk.Button(button_frame, text="Save to Excel", command=self.save_to_excel, state='disabled')
        self.save_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = ttk.Button(button_frame, text="Clear", command=self.clear_data)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # Treeview Frame
        tree_frame = ttk.LabelFrame(self.window, text="POS Data Preview", padding=10)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create Treeview with scrollbars
        self.columns = [
            'Filename', 'Date', 'Time', 'Latitude', 'Longitude', 'Height',
            'Q', 'Ns', 'Sdn', 'Sde', 'Sdu', 'Sdne', 'Sdnu', 'Sdeu', 'Age', 'Ratio'
        ]
        
        self.tree = ttk.Treeview(tree_frame, columns=self.columns, show='headings', height=15)
        
        # Configure columns
        column_widths = {
            'Filename': 150, 'Date': 100, 'Time': 100, 'Latitude': 120,
            'Longitude': 120, 'Height': 100, 'Q': 50, 'Ns': 50,
            'Sdn': 80, 'Sde': 80, 'Sdu': 80, 'Sdne': 80,
            'Sdnu': 80, 'Sdeu': 80, 'Age': 80, 'Ratio': 80
        }

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[col], anchor=tk.CENTER)

        # Add scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Configure grid weights
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def browse_input(self):
        directory = filedialog.askdirectory()
        if directory:
            self.input_dir.set(directory)

    def browse_output(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.output_file.set(file)

    def preview_data(self):
        input_dir = self.input_dir.get()
        
        if not input_dir:
            messagebox.showerror("Error", "Please select input directory")
            return

        # Clear existing items
        self.clear_data()

        all_data = []
        try:
            for file in os.listdir(input_dir):
                if file.endswith(".pos"):
                    with open(os.path.join(input_dir, file), 'r') as f:
                        for line in f:
                            if not line.startswith('%'):
                                values = line.strip().split()
                                if len(values) >= 15:
                                    data = {
                                        'Filename': file,
                                        'Date': values[0],
                                        'Time': values[1],
                                        'Latitude': values[2],
                                        'Longitude': values[3],
                                        'Height': values[4],
                                        'Q': values[5],
                                        'Ns': values[6],
                                        'Sdn': values[7],
                                        'Sde': values[8],
                                        'Sdu': values[9],
                                        'Sdne': values[10],
                                        'Sdnu': values[11],
                                        'Sdeu': values[12],
                                        'Age': values[13],
                                        'Ratio': values[14]
                                    }
                                    all_data.append(data)
                                    self.tree.insert("", tk.END, values=tuple(data.values()))

            if all_data:
                self.save_button.configure(state='normal')
                self.data_to_save = all_data
                messagebox.showinfo("Preview", f"Found {len(all_data)} records")
            else:
                messagebox.showwarning("Warning", "No data found in POS files")

        except Exception as e:
            messagebox.showerror("Error", f"Error reading files: {str(e)}")

    def clear_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.save_button.configure(state='disabled')
        self.data_to_save = None

    def save_to_excel(self):
        if not hasattr(self, 'data_to_save') or not self.data_to_save:
            messagebox.showerror("Error", "No data to save! Please preview data first.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.data_to_save)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Saved {len(self.data_to_save)} records to Excel!")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving to Excel: {str(e)}")


class DMSConverter:
    def __init__(self, window):
        self.window = window
        self.window.title("DMS Converter")
        self.window.geometry("400x200")

        # Create main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # DMS Input
        ttk.Label(main_frame, text="DMS Input:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.dms_input = ttk.Entry(main_frame, width=40)
        self.dms_input.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(main_frame, text="Format: 73° 9' 18.99435\"").grid(row=1, column=1, sticky=tk.W)

        # Convert Button
        ttk.Button(main_frame, text="Convert", command=self.convert_dms).grid(row=2, column=1, pady=10)

        # Decimal Degrees Output
        ttk.Label(main_frame, text="Decimal Degrees:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.dd_output = ttk.Entry(main_frame, width=40, state='readonly')
        self.dd_output.grid(row=3, column=1, padx=5, pady=5)

    def convert_dms(self):
        dms = self.dms_input.get()
        try:
            dd_value = dms_to_dd(dms)
            self.dd_output.configure(state='normal')
            self.dd_output.delete(0, tk.END)
            self.dd_output.insert(0, f"{dd_value:.9f}")
            self.dd_output.configure(state='readonly')
        except ValueError as e:
            messagebox.showerror("Error", str(e))


class R27Converter:
    def __init__(self, window):
        self.window = window
        self.window.title("F16 to R27 Converter")
        self.window.geometry("600x300")

        # Create main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Source Directory
        ttk.Label(main_frame, text="Source Folder:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.source_folder = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.source_folder, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_source).grid(row=0, column=2, padx=5)

        # Destination Directory
        ttk.Label(main_frame, text="Destination Folder:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.destination_folder = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.destination_folder, width=40).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_destination).grid(row=1, column=2, padx=5)

        # Convert Button
        ttk.Button(main_frame, text="Convert Files", command=self.convert).grid(row=2, column=1, pady=20)

    def browse_source(self):
        directory = filedialog.askdirectory()
        if directory:
            self.source_folder.set(directory)

    def browse_destination(self):
        directory = filedialog.askdirectory()
        if directory:
            self.destination_folder.set(directory)

    def convert(self):
        source = self.source_folder.get()
        destination = self.destination_folder.get()

        if not source or not destination:
            messagebox.showerror("Error", "Please select both source and destination directories")
            return

        try:
            converted_count = 0
            for file in os.listdir(source):
                if file.endswith(".F16"):
                    source_path = os.path.join(source, file)
                    dest_path = os.path.join(destination, file.replace(".F16", ".R27"))
                    shutil.copy(source_path, dest_path)
                    converted_count += 1

            if converted_count > 0:
                messagebox.showinfo("Success", f"Successfully converted {converted_count} files!")
            else:
                messagebox.showinfo("Info", "No .F16 files found in the source directory")

        except Exception as e:
            messagebox.showerror("Error", f"Error during conversion: {str(e)}")


def dms_to_dd(dms_str):
    # Extract the degrees, minutes, and seconds using regex
    match = re.match(r"(-?\d+)°\s*(\d+)'?\s*(\d+(\.\d+)?)\"?", dms_str)
    if not match:
        raise ValueError(f"Invalid DMS format: {dms_str}")
    
    degrees = float(match.group(1))
    minutes = float(match.group(2))
    seconds = float(match.group(3))
    
    # Convert to decimal degrees
    dd = abs(degrees) + (minutes / 60) + (seconds / 3600)
    
    # Apply the sign of the degrees
    if degrees < 0:
        dd = -dd
    
    return dd