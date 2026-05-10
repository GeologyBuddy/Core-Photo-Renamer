import os
import sys
import webbrowser
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from PIL import Image, ImageTk  # For icons

REQUIRED_COLUMNS = { "Hole ID": None, "Box Number": None, "From (m)": None, "To (m)": None }

# ----------------------------
# BulkRenamerApp GUI Application
# ----------------------------
class BulkRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GB•CoreNamer 2026")
        self.root.geometry("800x600")

        # Determine the base path for loading resources
        if hasattr(sys, '_MEIPASS'):  # Running as bundled .exe?
            base_path = getattr(sys, '_MEIPASS')
        else:
            base_path = os.path.abspath(".")

        # Load logo images using the correct path
        try:
            logo_path = os.path.join(base_path, "logo.png")
            _logo_pil = Image.open(logo_path).resize((128, 128), Image.Resampling.LANCZOS)
            self.logo_img: ImageTk.PhotoImage | None = ImageTk.PhotoImage(_logo_pil)
            self.root.iconphoto(True, self.logo_img)  # type: ignore[arg-type]
        except Exception as e:
            print(f"Error loading logo: {e}")
            self.logo_img = None

        try:
            second_logo_path = os.path.join(base_path, "logo.png")
            _second_pil = Image.open(second_logo_path).resize((64, 64), Image.Resampling.LANCZOS)
            self.second_logo_img: ImageTk.PhotoImage | None = ImageTk.PhotoImage(_second_pil)
        except Exception as e:
            print(f"Error loading second logo: {e}")
            self.second_logo_img = None

        # Initialize variables for file/folder selection and renaming
        self.excel_path = ""
        self.folder_path = ""
        self.rename_history = []
        self.hole_id = ""
        self.column_mapping: dict = {}

        # UI widget placeholders (assigned in setup_ui)
        self.label_excel: tk.Label
        self.btn_excel: tk.Button
        self.excel_display: tk.Label
        self.label_folder: tk.Label
        self.btn_folder: tk.Button
        self.folder_display: tk.Label
        self.tree: ttk.Treeview
        self.progress: ttk.Progressbar
        self.btn_rename: tk.Button
        self.btn_undo: tk.Button

        # Create menu and UI elements
        self.create_menu()
        self.setup_ui()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New", command=self.new_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", command=self.undo_rename)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def setup_ui(self):
        # Header Frame: logo, software name, explanation.
        header_frame = tk.Frame(self.root, bg="#f0f0f0")
        header_frame.pack(side="top", fill="x", padx=5, pady=5)

        if self.second_logo_img:
            second_logo_label = tk.Label(header_frame, image=self.second_logo_img, bg="#f0f0f0")
            second_logo_label.pack(side="left", padx=5)

        name_label = tk.Label(
            header_frame,
            text="GB•CoreNamer 2026",
            font=("Calisto MT", 14, "bold"),
            bg="#f0f0f0"
        )
        name_label.pack(side="top", padx=5, anchor="w")

        explanation_label = tk.Label(
            header_frame,
            text="Easily rename your core photos in bulk with this powerful tool.\n"
                 "Note: Renaming works only for photos with four core boxes.",
            justify="left",
            font=("Verdana", 10),
            bg="#f0f0f0"
        )
        explanation_label.pack(side="top", padx=5, anchor="w")

        # File selection frame (Excel file)
        file_frame = tk.LabelFrame(self.root, text="Interval File", bg="#f0f0f0", padx=10, pady=5)
        file_frame.pack(fill="x", padx=5, pady=5)
        self.label_excel = tk.Label(file_frame, text="Select Interval File:", bg="#f0f0f0")
        self.label_excel.grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.btn_excel = tk.Button(file_frame, text="Browse", command=self.load_excel, bg="#4CAF50", fg="black")
        self.btn_excel.grid(row=0, column=1, padx=5, pady=2)
        self.excel_display = tk.Label(file_frame, text="No file selected", bg="#f0f0f0")
        self.excel_display.grid(row=0, column=2, padx=5, pady=2, sticky="w")

        # Folder selection frame (Image folder)
        folder_frame = tk.LabelFrame(self.root, text="Image Folder", bg="#f0f0f0", padx=10, pady=5)
        folder_frame.pack(fill="x", padx=5, pady=5)
        self.label_folder = tk.Label(folder_frame, text="Select Image Folder:", bg="#f0f0f0")
        self.label_folder.grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.btn_folder = tk.Button(folder_frame, text="Browse", command=self.load_folder, bg="#4CAF50", fg="black")
        self.btn_folder.grid(row=0, column=1, padx=5, pady=2)
        self.folder_display = tk.Label(folder_frame, text="No folder selected", bg="#f0f0f0")
        self.folder_display.grid(row=0, column=2, padx=5, pady=2, sticky="w")

        # Main frame for preview and buttons
        main_frame = tk.Frame(self.root)
        main_frame.pack(side="top", fill="both", expand=True, padx=5, pady=5)

        # Preview table with scrollbars
        preview_frame = tk.Frame(main_frame)
        preview_frame.pack(side="top", fill="both", expand=True)
        self.tree = ttk.Treeview(preview_frame, columns=("Old Filename", "New Filename"), show="headings")
        self.tree.heading("Old Filename", text="Old Filename")
        self.tree.heading("New Filename", text="New Filename")
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(side="top", fill="x", pady=5)

        # Buttons frame for renaming and undo
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(side="top", fill="x", pady=5)
        self.btn_rename = tk.Button(btn_frame, text="Rename Files", command=self.rename_files, state="disabled",
                                    bg="#008CBA", fg="black")
        self.btn_rename.pack(side="left", padx=5)
        self.btn_undo = tk.Button(btn_frame, text="Undo Rename", command=self.undo_rename, state="disabled",
                                  bg="#f44336", fg="black")
        self.btn_undo.pack(side="left", padx=5)

    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title("About")
        about_win.resizable(False, False)
        about_frame = tk.Frame(about_win)
        about_frame.pack(padx=10, pady=10)
        if self.logo_img:
            logo_label = tk.Label(about_frame, image=self.logo_img)
            logo_label.grid(row=0, column=0, rowspan=3, padx=5, pady=5)
        name_label = tk.Label(about_frame, text="GB•CoreNamer 2026", font=("Calisto MT", 14, "bold"))
        name_label.grid(row=0, column=1, sticky="w", padx=5, pady=(5, 0))
        info = (
            "Version# 26.4.26.001\n"
            "Built on March 26, 2026\n\n"
            "Created and Developed by Mehmet Duyan\n"
            "Copyright © Mehmet Duyan\n\n"
            "GB•CoreNamer is licensed under the GNU General Public License\n"
            "https://www.gnu.org/licenses"

        )
        info_label = tk.Label(about_frame, text=info, justify="left")
        info_label.grid(row=1, column=1, sticky="w", padx=5)
        links_frame = tk.Frame(about_win)
        links_frame.pack(padx=10, pady=(0, 15))
        tk.Button(
            links_frame, text="🌐  Official Website",
            command=lambda: webbrowser.open_new("https://geologybuddy.com"),
            bg="#4CAF50", fg="black", cursor="hand2", relief="flat", padx=10, pady=4
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            links_frame, text="🐙  GitHub",
            command=lambda: webbrowser.open_new("https://github.com/GeologyBuddy"),
            bg="#24292e", fg="white", cursor="hand2", relief="flat", padx=10, pady=4
        ).pack(side="left")
        self.center_window(about_win)

    @staticmethod
    def center_window(win: tk.Toplevel) -> None:
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
        )
        if self.excel_path:
            self.excel_display.config(text=os.path.basename(self.excel_path))
            try:
                df = self.read_table_file(self.excel_path)
                self.column_mapping = self.prompt_column_mapping(df.columns.tolist())
                mapped_cols = [self.column_mapping.get(col) for col in REQUIRED_COLUMNS]
                if None in mapped_cols:
                    raise ValueError("All required columns must be mapped.")
                self.hole_id = str(df.iloc[0][self.column_mapping["Hole ID"]])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file: {e}")
                return
            self.update_preview()
            self.check_ready()

    @staticmethod
    def read_table_file(path: str) -> pd.DataFrame:
        """Reads Excel (.xlsx, .xls) or CSV files into a DataFrame."""
        if path.lower().endswith(".csv"):
            df: pd.DataFrame = pd.read_csv(path)  # type: ignore[assignment]
        else:
            df = pd.read_excel(path)
        return df

    def load_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_display.config(text=os.path.basename(self.folder_path))
            self.update_preview()
            self.check_ready()

    def prompt_column_mapping(self, excel_columns):
        mapping_win = tk.Toplevel(self.root)
        mapping_win.title("Map Data Columns")
        mapping_win.geometry("350x300")
        mapping_win.resizable(False, False)
        mapping_win.grab_set()

        frame = tk.Frame(mapping_win)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        tk.Label(frame, text="Match required fields to your data columns:", font=("Arial", 10, "bold")).pack(
            pady=(0, 10))

        mappings = {}
        dropdown_vars = {}

        for field in REQUIRED_COLUMNS:
            row_frame = tk.Frame(frame)
            row_frame.pack(fill="x", pady=5)

            tk.Label(row_frame, text=field, width=15, anchor="w").pack(side="left")
            var = tk.StringVar()
            dropdown = ttk.Combobox(row_frame, textvariable=var, values=excel_columns, state="readonly", width=25)
            dropdown.pack(side="right", padx=5)
            dropdown_vars[field] = var

        def confirm():
            for col_name, col_var in dropdown_vars.items():
                selected = col_var.get()
                if not selected:
                    messagebox.showerror("Error", f"Please select a column for '{col_name}'", parent=mapping_win)
                    return
                mappings[col_name] = selected
            mapping_win.destroy()

        confirm_btn = tk.Button(frame, text="Confirm", command=confirm, bg="#4CAF50", fg="white")
        confirm_btn.pack(pady=15)

        self.root.wait_window(mapping_win)
        return mappings

    def check_ready(self):
        if self.excel_path and self.folder_path and self.hole_id:
            self.btn_rename.config(state="normal")

    def update_preview(self):
        if not self.excel_path or not self.folder_path:
            return
        self.tree.delete(*self.tree.get_children())
        try:
            df = self.read_table_file(self.excel_path)
            files = sorted(os.listdir(self.folder_path))
            folder_base = os.path.basename(self.folder_path).lower()
            if "dry" in folder_base:
                photo_type = "Dry"
            elif "wet" in folder_base:
                photo_type = "Wet"
            else:
                photo_type = ""
            total_rows = len(df)
            full_groups = total_rows // 4
            for group_idx in range(full_groups):
                group = df.iloc[group_idx * 4: group_idx * 4 + 4]
                hole_id = str(group.iloc[0][self.column_mapping["Hole ID"]])
                start_box = int(group.iloc[0][self.column_mapping["Box Number"]])
                end_box = int(group.iloc[-1][self.column_mapping["Box Number"]])
                from_val = float(group.iloc[0][self.column_mapping["From (m)"]])
                to_val = float(group.iloc[-1][self.column_mapping["To (m)"]])
                box_range_str = f"Bx{start_box:03d}-{end_box:03d}"
                meter_interval_str = f"{from_val:06.1f}m-{to_val:06.1f}m"
                new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}.jpg"
                old_filename = files[group_idx] if group_idx < len(files) else "N/A"
                self.tree.insert("", "end", values=(old_filename, new_filename))
            remainder = total_rows % 4
            if remainder:
                group = df.iloc[full_groups * 4:]
                hole_id = str(group.iloc[0][self.column_mapping["Hole ID"]])
                start_box = int(group.iloc[0][self.column_mapping["Box Number"]])
                end_box = int(group.iloc[-1][self.column_mapping["Box Number"]])
                from_val = float(group.iloc[0][self.column_mapping["From (m)"]])
                to_val = float(group.iloc[-1][self.column_mapping["To (m)"]])
                box_range_str = f"Bx{start_box:03d}-{end_box:03d}"
                meter_interval_str = f"{from_val:06.1f}m-{to_val:06.1f}m"
                new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}.jpg"
                idx = full_groups
                old_filename = files[idx] if idx < len(files) else "N/A"
                self.tree.insert("", "end", values=(old_filename, new_filename))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during preview: {e}")

    def rename_files(self):
        def rename_task():
            try:
                df = self.read_table_file(self.excel_path)
                files = sorted(os.listdir(self.folder_path))
                self.rename_history = []
                total_rows = len(df)
                full_groups = total_rows // 4
                remainder = total_rows % 4
                total_groups = full_groups + (1 if remainder else 0)
                self.progress["maximum"] = total_groups
                folder_base = os.path.basename(self.folder_path).lower()
                if "dry" in folder_base:
                    photo_type = "Dry"
                elif "wet" in folder_base:
                    photo_type = "Wet"
                else:
                    photo_type = ""
                group_counter = 0
                for group_idx in range(full_groups):
                    group = df.iloc[group_idx * 4: group_idx * 4 + 4]
                    hole_id = str(group.iloc[0][self.column_mapping["Hole ID"]])
                    start_box = int(group.iloc[0][self.column_mapping["Box Number"]])
                    end_box = int(group.iloc[-1][self.column_mapping["Box Number"]])
                    from_val = float(group.iloc[0][self.column_mapping["From (m)"]])
                    to_val = float(group.iloc[-1][self.column_mapping["To (m)"]])
                    box_range_str = f"Bx{start_box:03d}-{end_box:03d}"
                    meter_interval_str = f"{from_val:06.1f}m-{to_val:06.1f}m"
                    new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}.jpg"
                    if group_counter >= len(files):
                        self.root.after(0, lambda: messagebox.showwarning("Warning",
                                                                          "Not enough files in the folder for all composite photos."))
                        break
                    old_filename = files[group_counter]
                    old_path = os.path.join(self.folder_path, old_filename)
                    new_path = os.path.join(self.folder_path, new_filename)
                    counter = 1
                    while os.path.exists(new_path):
                        new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}_{counter}.jpg"
                        new_path = os.path.join(self.folder_path, new_filename)
                        counter += 1
                    os.rename(old_path, new_path)
                    self.rename_history.append((new_path, old_path))
                    self.root.after(0, lambda: self.progress.step(1))
                    group_counter += 1
                if remainder:
                    group = df.iloc[full_groups * 4:]
                    hole_id = str(group.iloc[0][self.column_mapping["Hole ID"]])
                    start_box = int(group.iloc[0][self.column_mapping["Box Number"]])
                    end_box = int(group.iloc[-1][self.column_mapping["Box Number"]])
                    from_val = float(group.iloc[0][self.column_mapping["From (m)"]])
                    to_val = float(group.iloc[-1][self.column_mapping["To (m)"]])
                    box_range_str = f"Bx{start_box:03d}-{end_box:03d}"
                    meter_interval_str = f"{from_val:06.1f}m-{to_val:06.1f}m"
                    new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}.jpg"
                    if group_counter >= len(files):
                        self.root.after(0, lambda: messagebox.showwarning("Warning",
                                                                          "Not enough files in the folder for the last composite photo."))
                    else:
                        old_filename = files[group_counter]
                        old_path = os.path.join(self.folder_path, old_filename)
                        new_path = os.path.join(self.folder_path, new_filename)
                        counter = 1
                        while os.path.exists(new_path):
                            new_filename = f"{hole_id}_{box_range_str}_{meter_interval_str}_{photo_type}_{counter}.jpg"
                            new_path = os.path.join(self.folder_path, new_filename)
                            counter += 1
                        os.rename(old_path, new_path)
                        self.rename_history.append((new_path, old_path))
                        self.root.after(0, lambda: self.progress.step(1))
                self.root.after(0, lambda: messagebox.showinfo("Success", "Files renamed successfully."))
                self.root.after(0, lambda: self.btn_undo.config(state="normal"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}"))

        threading.Thread(target=rename_task).start()

    def undo_rename(self):
        if not self.rename_history:
            messagebox.showinfo("Info", "No rename actions to undo.")
            return
        for new_path, old_path in reversed(self.rename_history):
            os.rename(new_path, old_path)
        messagebox.showinfo("Success", "Undo completed.")
        self.btn_undo.config(state="disabled")
        self.update_preview()

    def new_file(self):
        self.excel_path = ""
        self.folder_path = ""
        self.rename_history = []
        self.hole_id = ""
        self.excel_display.config(text="No file selected")
        self.folder_display.config(text="No folder selected")
        self.btn_rename.config(state="disabled")
        self.btn_undo.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.progress["value"] = 0


if __name__ == "__main__":
    main_root = tk.Tk()
    app = BulkRenamerApp(main_root)
    try:
        main_root.mainloop()
    except KeyboardInterrupt:
        print("Application closed by user.")