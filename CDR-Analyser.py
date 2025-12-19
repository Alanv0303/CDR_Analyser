import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from datetime import datetime
import os
import threading
import numpy as np
from tkcalendar import DateEntry

class CDRAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CDR Data Analyzer")
        self.root.geometry("1200x800")
        self.root.configure(bg="#f0f0f0")
       
        self.df = None
        self.filtered_df = None
        self.results = {}
        self.current_figure = None
        self.canvas = None
        self.column_mappings = {}
       
        # Create a style
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TButton", background="#4CAF50", foreground="black", font=("Arial", 10))
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 14, "bold"))
       
        # Create main frames
        self.create_frames()
       
        # Create widgets
        self.create_input_widgets()
        self.create_analysis_widgets()
        self.create_results_widgets()
       
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
       
    def create_frames(self):
        # Main container with padding
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
       
        # Top frame for input controls
        self.input_frame = ttk.LabelFrame(self.main_frame, text="Input Settings", padding="10")
        self.input_frame.pack(fill=tk.X, pady=5)
       
        # Middle frame for analysis options
        self.analysis_frame = ttk.LabelFrame(self.main_frame, text="Analysis Options", padding="10")
        self.analysis_frame.pack(fill=tk.X, pady=5)
       
        # Bottom frame for results (split into two parts)
        self.results_frame = ttk.Frame(self.main_frame, padding="10")
        self.results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
       
        # Left frame for results controls
        self.results_controls_frame = ttk.LabelFrame(self.results_frame, text="Results", padding="10")
        self.results_controls_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
       
        # Right frame for visualization
        self.visualization_frame = ttk.LabelFrame(self.results_frame, text="Visualization", padding="10")
        self.visualization_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
       
    def create_input_widgets(self):
        # File input row
        file_frame = ttk.Frame(self.input_frame)
        file_frame.pack(fill=tk.X, pady=5)
       
        ttk.Label(file_frame, text="CDR File:").pack(side=tk.LEFT, padx=(0, 5))
       
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
       
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).pack(side=tk.LEFT)
       
        # Date range row
        date_frame = ttk.Frame(self.input_frame)
        date_frame.pack(fill=tk.X, pady=5)
       
        ttk.Label(date_frame, text="Date Range:").pack(side=tk.LEFT, padx=(0, 5))
       
        self.start_date_var = tk.StringVar(value="2024-03-01")
        self.end_date_var = tk.StringVar(value="2024-03-17")
       
        ttk.Label(date_frame, text="From:").pack(side=tk.LEFT, padx=(10, 5))
        DateEntry(date_frame, width=12, textvariable=self.start_date_var, date_pattern='y-mm-dd').pack(side=tk.LEFT, padx=(0, 10))
       
        ttk.Label(date_frame, text="To:").pack(side=tk.LEFT, padx=(10, 5))
        DateEntry(date_frame, width=12, textvariable=self.end_date_var, date_pattern='y-mm-dd').pack(side=tk.LEFT)
       
        # Load data button
        ttk.Button(self.input_frame, text="Load Data", command=self.load_data).pack(pady=10)
       
    def create_analysis_widgets(self):
        # Analysis options
        options_frame = ttk.Frame(self.analysis_frame)
        options_frame.pack(fill=tk.X, pady=5)
       
        # Location analysis
        self.location_var = tk.StringVar(value="main_city")
        ttk.Label(options_frame, text="Location Analysis:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Radiobutton(options_frame, text="Main City", variable=self.location_var, value="main_city").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(options_frame, text="Sub City", variable=self.location_var, value="sub_city").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(options_frame, text="Cell ID", variable=self.location_var, value="cell_id").pack(side=tk.LEFT, padx=5)
       
        # Top N records
        count_frame = ttk.Frame(self.analysis_frame)
        count_frame.pack(fill=tk.X, pady=5)
       
        ttk.Label(count_frame, text="Top Results:").pack(side=tk.LEFT, padx=(0, 5))
        self.top_n_var = tk.StringVar(value="10")
        ttk.Spinbox(count_frame, from_=5, to=50, increment=5, textvariable=self.top_n_var, width=5).pack(side=tk.LEFT)
       
        # Analysis type
        analysis_type_frame = ttk.Frame(self.analysis_frame)
        analysis_type_frame.pack(fill=tk.X, pady=5)
       
        ttk.Label(analysis_type_frame, text="Analysis Type:").pack(side=tk.LEFT, padx=(0, 5))
        self.analysis_type_var = tk.StringVar(value="location")
        ttk.Radiobutton(analysis_type_frame, text="Location Analysis", variable=self.analysis_type_var, value="location").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(analysis_type_frame, text="Called Numbers", variable=self.analysis_type_var, value="numbers").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(analysis_type_frame, text="Call Volume by Date", variable=self.analysis_type_var, value="date").pack(side=tk.LEFT, padx=5)
       
        # Analyze button
        ttk.Button(self.analysis_frame, text="Analyze Data", command=self.analyze_data).pack(pady=10)
       
    def create_results_widgets(self):
        # Results list
        self.results_listbox = tk.Listbox(self.results_controls_frame, width=30, height=20)
        self.results_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)
       
        # Export results button
        export_frame = ttk.Frame(self.results_controls_frame)
        export_frame.pack(fill=tk.X, pady=5)
       
        ttk.Button(export_frame, text="Export Results", command=self.export_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Export Chart", command=self.export_chart).pack(side=tk.LEFT, padx=5)
       
        # Placeholder for chart
        self.fig_placeholder = ttk.Label(self.visualization_frame, text="Analysis results will appear here")
        self.fig_placeholder.pack(fill=tk.BOTH, expand=True)
       
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CDR Data File",
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)
           
    def load_data(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a file first")
            return
           
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "Selected file does not exist")
            return
           
        try:
            self.status_var.set("Loading data...")
            self.root.update_idletasks()
           
            # Start in a separate thread to prevent UI freezing
            threading.Thread(target=self._load_data_thread, args=(file_path,), daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
            self.status_var.set("Ready")
           
    def _load_data_thread(self, file_path):
        try:
            # Determine file type and load accordingly
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                # Load Excel file
                self.df = pd.read_excel(file_path)
            elif file_ext == '.csv':
                # Load CSV file - try multiple encodings
                encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'utf-16']
                loaded = False
                
                for encoding in encodings:
                    try:
                        self.df = pd.read_csv(file_path, sep=None, engine='python', encoding=encoding, encoding_errors='replace')
                        loaded = True
                        break
                    except (UnicodeDecodeError, Exception):
                        continue
                
                if not loaded:
                    raise ValueError("Could not decode file with any standard encoding")
            else:
                # Try CSV as default with multiple encodings
                encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'utf-16']
                loaded = False
                
                for encoding in encodings:
                    try:
                        self.df = pd.read_csv(file_path, sep=None, engine='python', encoding=encoding, encoding_errors='replace')
                        loaded = True
                        break
                    except (UnicodeDecodeError, Exception):
                        continue
                
                if not loaded:
                    raise ValueError("Could not decode file with any standard encoding")
           
            # Update UI in the main thread
            self.root.after(0, self._update_ui_after_load)
        except Exception as e:
            # Capture exception message before passing to lambda
            error_msg = f"Error loading data: {str(e)}\n\nPlease ensure:\n1. File is not open in another program\n2. File format is correct\n3. File is not corrupted\n4. Try saving as Excel (.xlsx) if CSV fails"
            self.root.after(0, lambda msg=error_msg: self._show_error(msg))
            self.root.after(0, lambda: self.status_var.set("Failed to load"))
           
    def _update_ui_after_load(self):
        if self.df is not None and not self.df.empty:
            try:
                # Show column mapping dialog
                self.root.after(0, self._show_column_mapping_dialog)
            except Exception as e:
                error_msg = f"Error processing data: {str(e)}"
                self._show_error(error_msg)
                self.status_var.set("Error processing data")
        else:
            self.status_var.set("Failed to load data - file is empty")
            messagebox.showerror("Error", "The file appears to be empty or could not be read")
           
    def _show_column_mapping_dialog(self):
        # Create a dialog for column mapping
        mapping_dialog = tk.Toplevel(self.root)
        mapping_dialog.title("Map CDR Columns")
        mapping_dialog.geometry("600x700")
        mapping_dialog.transient(self.root)
        mapping_dialog.grab_set()
       
        # Required column mappings
        required_fields = [
            ("Date Column", "date_col"),
            ("Time Column (optional)", "time_col"),
            ("Phone Number (B Party) Column", "phone_col"),
            ("Main Location Column", "main_loc_col"),
            ("Sub Location Column (optional)", "sub_loc_col"),
            ("Cell ID/Address Column (optional)", "cell_id_col")
        ]
       
        # Store the mapping variables
        self.column_mapping = {}
       
        # Create frame with scrollbar
        main_frame = ttk.Frame(mapping_dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
       
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
       
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
       
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
       
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
       
        # Add column detection info
        ttk.Label(scroll_frame, text="Detected Columns:", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
       
        for i, col in enumerate(self.df.columns):
            ttk.Label(scroll_frame, text=f"{i+1}. {col}", anchor="w").grid(row=i+1, column=0, columnspan=2, sticky="w", padx=5)
       
        # Add separator
        ttk.Separator(scroll_frame, orient="horizontal").grid(row=len(self.df.columns)+1, column=0, columnspan=2, sticky="ew", pady=10)
       
        # Add mapping fields
        ttk.Label(scroll_frame, text="Map CDR Fields:", font=("Arial", 12, "bold")).grid(row=len(self.df.columns)+2, column=0, columnspan=2, sticky="w", pady=(0, 10))
       
        row_offset = len(self.df.columns) + 3
       
        for i, (label, var_name) in enumerate(required_fields):
            ttk.Label(scroll_frame, text=label + ":", anchor="w").grid(row=row_offset+i, column=0, sticky="w", padx=5, pady=5)
           
            # Create combobox with column choices
            combo_var = tk.StringVar()
            combo = ttk.Combobox(scroll_frame, textvariable=combo_var, width=30, state='readonly')
            combo['values'] = [''] + list(self.df.columns)  # Add empty option for optional fields
           
            # Try to auto-select appropriate columns based on column name
            if "date" in var_name:
                for col in self.df.columns:
                    if "date" in col.lower():
                        combo_var.set(col)
                        break
            elif "time" in var_name:
                for col in self.df.columns:
                    if "time" in col.lower() and "date" not in col.lower():
                        combo_var.set(col)
                        break
            elif "phone" in var_name:
                for col in self.df.columns:
                    if any(keyword in col.lower() for keyword in ["party", "number", "b party", "called"]):
                        combo_var.set(col)
                        break
            elif "main_loc" in var_name:
                for col in self.df.columns:
                    if "main" in col.lower() and "city" in col.lower():
                        combo_var.set(col)
                        break
            elif "sub_loc" in var_name:
                for col in self.df.columns:
                    if "sub" in col.lower() and "city" in col.lower():
                        combo_var.set(col)
                        break
            elif "cell_id" in var_name:
                for col in self.df.columns:
                    if "cell" in col.lower() or "address" in col.lower():
                        combo_var.set(col)
                        break
                       
            combo.grid(row=row_offset+i, column=1, sticky="ew", padx=5, pady=5)
            self.column_mapping[var_name] = combo_var
       
        # Sample data preview
        ttk.Separator(scroll_frame, orient="horizontal").grid(row=row_offset+len(required_fields), column=0, columnspan=2, sticky="ew", pady=10)
       
        ttk.Label(scroll_frame, text="Data Preview (First 5 rows):", font=("Arial", 12, "bold")).grid(
            row=row_offset+len(required_fields)+1, column=0, columnspan=2, sticky="w", pady=(0, 10))
       
        # Add data preview (first 5 rows)
        preview_frame = ttk.Frame(scroll_frame)
        preview_frame.grid(row=row_offset+len(required_fields)+2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        preview_text = tk.Text(preview_frame, height=10, width=70, wrap=tk.NONE)
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=preview_text.yview)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=preview_text.xview)
        preview_text.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        
        preview_text.grid(row=0, column=0, sticky="nsew")
        preview_scroll_y.grid(row=0, column=1, sticky="ns")
        preview_scroll_x.grid(row=1, column=0, sticky="ew")
        
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
       
        # Format the preview data
        preview_str = self.df.head(5).to_string()
        preview_text.insert(tk.END, preview_str)
        preview_text.config(state=tk.DISABLED)
       
        # Add buttons
        button_frame = ttk.Frame(mapping_dialog)
        button_frame.pack(fill=tk.X, pady=10, padx=10)
       
        ttk.Button(button_frame, text="Cancel", command=mapping_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Confirm Mapping", command=lambda: self._process_column_mapping(mapping_dialog)).pack(side=tk.RIGHT, padx=5)
   
    def _process_column_mapping(self, dialog):
        # Get values from mapping
        mapping = {key: var.get() for key, var in self.column_mapping.items()}
       
        # Validate required fields
        if not mapping['date_col'] or not mapping['phone_col'] or not mapping['main_loc_col']:
            messagebox.showerror("Error", "Date, Phone Number, and Main Location columns are required")
            return
           
        try:
            # Process datetime
            if mapping['time_col'] and mapping['time_col'] != '':
                # If we have separate date and time columns
                self.df['DateTime'] = pd.to_datetime(
                    self.df[mapping['date_col']].astype(str) + ' ' + self.df[mapping['time_col']].astype(str),
                    errors='coerce'
                )
            else:
                # If date column contains time information or we only have date
                self.df['DateTime'] = pd.to_datetime(self.df[mapping['date_col']], errors='coerce')
            
            # Check if datetime conversion was successful
            if self.df['DateTime'].isna().all():
                messagebox.showerror("Error", "Could not parse dates. Please check the date format in your file.")
                return
               
            # Store column mappings for later use
            self.column_mappings = mapping
           
            # Close dialog
            dialog.destroy()
           
            # Update status
            messagebox.showinfo("Success", f"Successfully loaded {len(self.df)} records\n\nDate range: {self.df['DateTime'].min()} to {self.df['DateTime'].max()}")
            self.status_var.set(f"Loaded {len(self.df)} records")
           
        except Exception as e:
            error_msg = f"Error processing dates: {str(e)}"
            self._show_error(error_msg)
           
    def _show_error(self, message):
        messagebox.showerror("Error", message)
        self.status_var.set("Error")
           
    def analyze_data(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load data first")
            return
        
        if not self.column_mappings:
            messagebox.showerror("Error", "Column mappings not set. Please reload the data.")
            return
           
        try:
            self.status_var.set("Analyzing data...")
            self.root.update_idletasks()
           
            # Start analysis in a separate thread to prevent UI freezing
            threading.Thread(target=self._analyze_data_thread, daemon=True).start()
        except Exception as e:
            error_msg = f"Analysis failed: {str(e)}"
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Ready")
           
    def _analyze_data_thread(self):
        try:
            # Get date range
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
           
            # Filter data for the specified date range
            self.filtered_df = self.df[(self.df['DateTime'] >= start_date) & (self.df['DateTime'] <= end_date)].copy()
           
            if self.filtered_df.empty:
                self.root.after(0, lambda: messagebox.showwarning("Warning", f"No data found in the specified date range\n\nAvailable date range: {self.df['DateTime'].min()} to {self.df['DateTime'].max()}"))
                self.root.after(0, lambda: self.status_var.set("No data in date range"))
                return
               
            # Get analysis type
            analysis_type = self.analysis_type_var.get()
            location_type = self.location_var.get()
            top_n = int(self.top_n_var.get())
           
            # Perform the analysis
            if analysis_type == "location":
                self._analyze_location(location_type, top_n)
            elif analysis_type == "numbers":
                self._analyze_numbers(top_n)
            else:  # date analysis
                self._analyze_date_volume()
               
            # Update UI in the main thread
            self.root.after(0, self._update_ui_after_analysis)
        except Exception as e:
            # Capture exception message before passing to lambda
            error_msg = f"Analysis error: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self._show_error(msg))
            self.root.after(0, lambda: self.status_var.set("Analysis failed"))
           
    def _analyze_location(self, location_type, top_n):
        try:
            # Map location type to column name
            if location_type == "main_city":
                column = self.column_mappings["main_loc_col"]
            elif location_type == "sub_city":
                column = self.column_mappings["sub_loc_col"]
                if not column:
                    raise ValueError("Sub location column not mapped")
            elif location_type == "cell_id":
                column = self.column_mappings["cell_id_col"]
                if not column:
                    raise ValueError("Cell ID column not mapped")
            else:
                column = self.column_mappings["main_loc_col"]
           
            # Calculate location counts
            location_counts = self.filtered_df[column].value_counts().reset_index()
            location_counts.columns = ['Location', 'Count']
           
            # Store results
            self.results = {
                "type": "location",
                "subtype": location_type,
                "data": location_counts.head(top_n),
                "title": f"Top {top_n} Common Locations"
            }
        except Exception as e:
            error_msg = f"Error in location analysis: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self._show_error(msg))
       
    def _analyze_numbers(self, top_n):
        try:
            # Use the mapped phone column
            phone_column = self.column_mappings['phone_col']
           
            # Calculate most called numbers
            number_counts = self.filtered_df[phone_column].value_counts().reset_index()
            number_counts.columns = ['Phone Number', 'Count']
           
            # Store results
            self.results = {
                "type": "numbers",
                "data": number_counts.head(top_n),
                "title": f"Top {top_n} Most Called Numbers"
            }
        except Exception as e:
            error_msg = f"Error in number analysis: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self._show_error(msg))
       
    def _analyze_date_volume(self):
        try:
            # Create date column
            self.filtered_df['Date'] = self.filtered_df['DateTime'].dt.date
           
            # Calculate call volume by date
            date_counts = self.filtered_df['Date'].value_counts().sort_index().reset_index()
            date_counts.columns = ['Date', 'Call Count']
           
            # Store results
            self.results = {
                "type": "date",
                "data": date_counts,
                "title": "Call Volume by Date"
            }
        except Exception as e:
            error_msg = f"Error in date analysis: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self._show_error(msg))
       
    def _update_ui_after_analysis(self):
        if not self.results or 'data' not in self.results:
            self.status_var.set("Analysis failed")
            return
           
        # Update results listbox
        self.results_listbox.delete(0, tk.END)
       
        data = self.results["data"]
        if self.results["type"] == "location":
            for idx, row in data.iterrows():
                self.results_listbox.insert(tk.END, f"{row['Location']}: {row['Count']} calls")
        elif self.results["type"] == "numbers":
            for idx, row in data.iterrows():
                self.results_listbox.insert(tk.END, f"{row['Phone Number']}: {row['Count']} calls")
        else:  # date
            for idx, row in data.iterrows():
                self.results_listbox.insert(tk.END, f"{row['Date']}: {row['Call Count']} calls")
               
        # Create visualization
        self._create_visualization()
       
        self.status_var.set(f"Analyzed {len(self.filtered_df)} records")
       
    def _create_visualization(self):
        # Clear previous plot
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None
           
        if self.fig_placeholder:
            self.fig_placeholder.destroy()
            self.fig_placeholder = None
        
        if self.current_figure:
            plt.close(self.current_figure)
            self.current_figure = None
           
        # Create new figure
        fig, ax = plt.subplots(figsize=(10, 6))
       
        data = self.results["data"]
       
        if self.results["type"] == "location":
            # For location analysis, use horizontal bar chart
            y_col = 'Location'
            x_col = 'Count'
            sns.barplot(x=x_col, y=y_col, data=data, ax=ax, palette='viridis')
           
        elif self.results["type"] == "numbers":
            # For number analysis, use horizontal bar chart
            y_col = 'Phone Number'
            x_col = 'Count'
            sns.barplot(x=x_col, y=y_col, data=data, ax=ax, palette='viridis')
           
        else:  # date analysis
            # For date analysis, use line chart
            sns.lineplot(x='Date', y='Call Count', data=data, ax=ax, marker='o')
            plt.xticks(rotation=45, ha='right')
           
        ax.set_title(self.results["title"], fontsize=14, fontweight='bold')
        plt.tight_layout()
       
        # Display in the UI
        self.canvas = FigureCanvasTkAgg(fig, master=self.visualization_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
       
        # Store figure for export
        self.current_figure = fig
       
    def export_results(self):
        if not hasattr(self, 'results') or not self.results or 'data' not in self.results:
            messagebox.showwarning("Warning", "No analysis results to export")
            return
           
        file_path = filedialog.asksaveasfilename(
            title="Save Results",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
            defaultextension=".csv"
        )
       
        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    self.results["data"].to_excel(file_path, index=False)
                else:
                    self.results["data"].to_csv(file_path, index=False)
                messagebox.showinfo("Success", f"Results exported to {file_path}")
            except Exception as e:
                error_msg = f"Failed to export results: {str(e)}"
                messagebox.showerror("Error", error_msg)
               
    def export_chart(self):
        if not hasattr(self, 'current_figure') or self.current_figure is None:
            messagebox.showwarning("Warning", "No chart to export")
            return
           
        file_path = filedialog.asksaveasfilename(
            title="Save Chart",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), ("All files", "*.*")],
            defaultextension=".png"
        )
       
        if file_path:
            try:
                self.current_figure.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Chart exported to {file_path}")
            except Exception as e:
                error_msg = f"Failed to export chart: {str(e)}"
                messagebox.showerror("Error", error_msg)
 
if __name__ == "__main__":
    root = tk.Tk()
    app = CDRAnalyzerApp(root)
    root.mainloop()
    