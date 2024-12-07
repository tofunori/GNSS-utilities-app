import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import shutil
import re
import matplotlib.pyplot as plt
from datetime import datetime
import matplotlib.dates as mdates


class UnifiedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unified Application")
        self.root.geometry("600x500")  # Made window slightly taller to accommodate title

        # Title Frame
        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, pady=20)

        # Main Title
        title_label = ttk.Label(title_frame, 
                              text="GNSS Processing tools",  # Add your title here
                              font=("Arial", 24, "bold"))
        title_label.pack()

        # Subtitle or description (optional)
        subtitle_label = ttk.Label(title_frame, 
                                 text="Data Processing and conversion utilities",  # Add your subtitle here
                                 font=("Arial", 12))
        subtitle_label.pack(pady=5)

        # Separator
        ttk.Separator(root, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)

        # Menu Label
        ttk.Label(root, text="Select a Tool", font=("Arial", 18)).pack(pady=20)

        # Tool Buttons
        ttk.Button(root, text="GNSS Data Viewer", command=self.open_gnss_window).pack(pady=10)
        ttk.Button(root, text="POS to Excel Converter", command=self.open_pos_window).pack(pady=10)
        ttk.Button(root, text="DMS Converter", command=self.open_dms_window).pack(pady=10)
        ttk.Button(root, text="F16 to R27 Converter", command=self.open_r27_window).pack(pady=10)

    def open_gnss_window(self):
        gnss_window = tk.Toplevel(self.root)
        gnss_window.title("GNSS Data Viewer with GNSS Selection")
        gnss_window.geometry("1200x800")
        gnss_window.transient(self.root)  # Set the window as transient
        gnss_window.focus_set()  # Set focus to the new window
        self.root.lower()  # Lower the main window

        fields = [
            "Date (UTC)", "File", "GNSS Model", "Description of Occupation",
            "Time Start (UTC)", "Time End (UTC)", "Duration", "Interval",
            "Latitude (DD)", "Longitude (DD)", "UTM N (m)", "UTM E (m)", "Elevation (m)",
            "Reference Point", "Sigma UTM N (m)", "Sigma UTM E (m)", "Sigma Elev. (m)", 
            "Datum", "Solution"
        ]

        # Frame for Data Entry
        form_frame = ttk.LabelFrame(gnss_window, text="Data Entry", padding=10)
        form_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)

        field_entries = {}

        def parse_sum_file(file_path):
            parsed_data = {
                "Date (UTC)": None, "File": None, "GNSS Model": None,
                "Description of Occupation": None, "Time Start (UTC)": None,
                "Time End (UTC)": None, "Duration": None, "Interval": None,
                "Latitude (DD)": None, "Longitude (DD)": None, "UTM N (m)": None,
                "UTM E (m)": None, "Elevation (m)": None, "Reference Point": "ARP",
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

        def validate_dd_format(value):
            try:
                float(value)
                return True
            except ValueError:
                return False

        def on_reference_point_change(event):
            ref_point = field_entries["Reference Point"].get()
            gnss_model = field_entries["GNSS Model"].get()
            
            try:
                current_elevation = float(field_entries["Elevation (m)"].get())
                if gnss_model == "EMLID INREACH RS2":
                    l1_offset = 0.135
                    l2_offset = 0.137
                    mean_offset = (l1_offset + l2_offset) / 2
                    
                    if ref_point == "APC":
                        adjusted_elevation = current_elevation + mean_offset
                    else:  # ARP
                        adjusted_elevation = current_elevation - mean_offset
                        
                    field_entries["Elevation (m)"].delete(0, tk.END)
                    field_entries["Elevation (m)"].insert(0, f"{adjusted_elevation:.6f}")
            except ValueError:
                pass

        def on_gnss_model_change(event):
            selected_model = field_entries["GNSS Model"].get()
            if selected_model == "EMLID INREACH RS2":
                field_entries["Reference Point"]["state"] = "readonly"
                if not field_entries["Reference Point"].get():
                    field_entries["Reference Point"].set("ARP")
            elif selected_model == "FOIF A30":
                field_entries["Reference Point"].set("APC")
                field_entries["Reference Point"]["state"] = "disabled"
            else:
                field_entries["Reference Point"].set("ARP")
                field_entries["Reference Point"]["state"] = "disabled"
            
            on_reference_point_change(None)

        def load_file():
            file_path = filedialog.askopenfilename(filetypes=[("SUM Files", "*.sum")])
            if file_path:
                parsed_data = parse_sum_file(file_path)
                if parsed_data:
                    for field, entry in field_entries.items():
                        if isinstance(entry, ttk.Combobox):
                            if field == "Description of Occupation":
                                entry.set(parsed_data[field] or "Base")
                            elif field == "Reference Point":
                                entry.set(parsed_data[field] or "ARP")
                        elif field in parsed_data and parsed_data[field] is not None:
                            entry.delete(0, tk.END)
                            entry.insert(0, parsed_data[field])

                    gnss_model = field_entries["GNSS Model"].get()
                    if gnss_model == "EMLID INREACH RS2":
                        field_entries["Reference Point"]["state"] = "readonly"
                    elif gnss_model == "FOIF A30":
                        field_entries["Reference Point"].set("APC")
                        field_entries["Reference Point"]["state"] = "disabled"
                    else:
                        field_entries["Reference Point"]["state"] = "disabled"

        def add_to_treeview():
            latitude = field_entries["Latitude (DD)"].get()
            longitude = field_entries["Longitude (DD)"].get()
            gnss_model = field_entries["GNSS Model"].get()

            if not validate_dd_format(latitude):
                messagebox.showwarning("Validation Error", "Latitude (DD) must be in decimal degrees format.")
                return
            if not validate_dd_format(longitude):
                messagebox.showwarning("Validation Error", "Longitude (DD) must be in decimal degrees format.")
                return
            if gnss_model == "N/A":
                messagebox.showwarning("Validation Error", "Please select a valid GNSS Model.")
                return

            values = [field_entries[field].get() for field in fields]
            tree.insert("", "end", values=values)

        def delete_entry():
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showwarning("Warning", "No entry selected to delete.")
                return
            for item in selected_items:
                tree.delete(item)

        def save_to_excel():
            if not tree.get_children():
                messagebox.showerror("Error", "No data to save!")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                data = []
                for item in tree.get_children():
                    data.append(tree.item(item)['values'])
                
                df = pd.DataFrame(data, columns=fields)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Data saved to Excel successfully!")

        def clear_fields():
            for field, entry in field_entries.items():
                if isinstance(entry, ttk.Combobox):
                    if field == "Description of Occupation":
                        entry.set("Base")
                    elif field == "GNSS Model":
                        entry.set("N/A")
                    elif field == "Reference Point":
                        entry.set("ARP")
                        entry.state(["disabled"])
                elif field == "Solution":
                    entry.delete(0, tk.END)
                    entry.insert(0, "PPP")
                elif field == "Datum":
                    entry.delete(0, tk.END)
                    entry.insert(0, "ITRF20")
                else:
                    entry.delete(0, tk.END)

        def convert_dms_to_dd(field_name):
            dms_value = field_entries[field_name].get()
            try:
                dd_value = dms_to_dd(dms_value)
                field_entries[field_name].delete(0, tk.END)
                field_entries[field_name].insert(0, dd_value)
            except Exception as e:
                messagebox.showerror("Error", f"Invalid DMS format: {str(e)}")

        # Create form fields
        for i, field in enumerate(fields):
            ttk.Label(form_frame, text=field).grid(row=i, column=0, padx=5, pady=5)
            if field == "GNSS Model":
                entry = ttk.Combobox(form_frame, values=["N/A", "EMLID INREACH RS2", "FOIF A30"], state="readonly")
                entry.set("N/A")
                entry.bind('<<ComboboxSelected>>', on_gnss_model_change)
            elif field == "Description of Occupation":
                entry = ttk.Combobox(form_frame, values=["Base", "Rover"], state="readonly")
                entry.set("Base")
            elif field == "Reference Point":
                entry = ttk.Combobox(form_frame, values=["ARP", "APC"], state="disabled")
                entry.set("ARP")
                entry.bind('<<ComboboxSelected>>', on_reference_point_change)
            elif field == "Solution":
                entry = ttk.Entry(form_frame, width=30)
                entry.insert(0, "PPP")
            elif field == "Datum":
                entry = ttk.Entry(form_frame, width=30)
                entry.insert(0, "ITRF20")
            else:
                entry = ttk.Entry(form_frame, width=30)
            entry.grid(row=i, column=1, padx=5, pady=5)
            field_entries[field] = entry

            # Add Convert DMS buttons
            if field in ["Latitude (DD)", "Longitude (DD)"]:
                ttk.Button(
                    form_frame, 
                    text="Convert DMS",
                    command=lambda f=field: convert_dms_to_dd(f)
                ).grid(row=i, column=2, padx=5, pady=5)

        # Buttons frame
        button_frame = ttk.Frame(gnss_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(button_frame, text="Import .sum File", command=load_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Add to Table", command=add_to_treeview).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Entry", command=delete_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Fields", command=clear_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save to Excel", command=save_to_excel).pack(side=tk.LEFT, padx=5)

        # Table Frame
        table_frame = ttk.LabelFrame(gnss_window, text="GNSS Data Table", padding=10)
        table_frame = ttk.LabelFrame(gnss_window, text="GNSS Data Table", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create Treeview
        tree = ttk.Treeview(table_frame, columns=fields, show="headings", height=15)
        for col in fields:
            tree.heading(col, text=col)
            tree.column(col, width=120)

        # Add scrollbars
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

    def open_pos_window(self):
        pos_window = tk.Toplevel(self.root)
        pos_window.transient(self.root)
        pos_window.focus_set()
        self.root.lower()
        PosToExcelConverter(pos_window)

    def open_dms_window(self):
        dms_window = tk.Toplevel(self.root)
        dms_window.transient(self.root)
        dms_window.focus_set()
        self.root.lower()
        DMSConverter(dms_window)

    def open_r27_window(self):
        r27_window = tk.Toplevel(self.root)
        r27_window.transient(self.root)
        r27_window.focus_set()
        self.root.lower()
        R27Converter(r27_window)


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


import re

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
    
    return f"{dd:.9f}"

# Example usage
dms_str = "-79° 58' 39.88108\""
print(dms_to_dd(dms_str))  # Should print -79.977744744

dms_str = "79° 58' 39.88108\""
print(dms_to_dd(dms_str))  # Should print 79.977744744


if __name__ == "__main__":
    root = tk.Tk()
    app = UnifiedApp(root)
    root.mainloop()