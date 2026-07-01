#from bearbetningsskript import process_all, get_default_config

#process_all(input_folder, output_folder, config)

import tkinter as tk
from tkinter import filedialog, messagebox
import json

from bearbetningsskript import process_all, get_default_config


# -------------------------
# LIST EDITOR COMPONENT
# -------------------------

class ListEditor(tk.Frame):
    def __init__(self, parent, title):
        super().__init__(parent)

        tk.Label(self, text=title, font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.listbox = tk.Listbox(self, height=6)
        self.listbox.pack(fill="both", expand=True, padx=2, pady=2)

        entry_frame = tk.Frame(self)
        entry_frame.pack(fill="x")

        self.entry = tk.Entry(entry_frame)
        self.entry.pack(side="left", fill="x", expand=True)

        tk.Button(entry_frame, text="Add", command=self.add_item).pack(side="left")
        tk.Button(entry_frame, text="Remove", command=self.remove_selected).pack(side="left")

    def add_item(self):
        value = self.entry.get().strip()
        if value:
            self.listbox.insert(tk.END, value)
            self.entry.delete(0, tk.END)

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)

    def set_values(self, values):
        self.listbox.delete(0, tk.END)
        for v in values:
            self.listbox.insert(tk.END, v)

    def get_values(self):
        return list(self.listbox.get(0, tk.END))


# -------------------------
# MAIN GUI APP
# -------------------------

class CertificateGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Splitter")
        self.root.geometry("900x750")

        self.config = get_default_config()

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.build_ui()
        self.load_config_into_ui()

    # -------------------------
    # UI LAYOUT
    # -------------------------

    def build_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=10)

        # INPUT
        tk.Label(frame, text="Input folder").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.input_path, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Browse", command=self.select_input).grid(row=0, column=2)

        # OUTPUT
        tk.Label(frame, text="Output folder").grid(row=1, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.output_path, width=60).grid(row=1, column=1, padx=5)
        tk.Button(frame, text="Browse", command=self.select_output).grid(row=1, column=2)

        # CONFIG EDITORS
        self.degree_editor = ListEditor(self.root, "START_PAGE_DEGREE_MARKERS")
        self.degree_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.award_editor = ListEditor(self.root, "START_PAGE_STRONG_AWARD_MARKERS")
        self.award_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.special_editor = ListEditor(self.root, "SPECIAL_START_PAGE_IDENTIFIERS")
        self.special_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.nonstart_editor = ListEditor(self.root, "NON_START_PAGE_PATTERNS")
        self.nonstart_editor.pack(fill="both", expand=True, padx=10, pady=5)

        # RUN BUTTON
        tk.Button(
            self.root,
            text="Run processing",
            command=self.run_processing,
            bg="#2d7d46",
            fg="white",
            height=2
        ).pack(fill="x", padx=10, pady=10)

        # LOG box
        self.log_box = tk.Text(self.root, height=8)
        self.log_box.pack(fill="both", padx=10, pady=10)

    # -------------------------
    # CONFIG HANDLING
    # -------------------------

    def load_config_into_ui(self):
        self.degree_editor.set_values(self.config["start_page_degree_markers"])
        self.award_editor.set_values(self.config["start_page_strong_award_markers"])
        self.special_editor.set_values(self.config["special_start_page_identifiers"])
        self.nonstart_editor.set_values(self.config["non_start_page_patterns"])

    def collect_config(self):
        return {
            "start_page_degree_markers": self.degree_editor.get_values(),
            "start_page_strong_award_markers": self.award_editor.get_values(),
            "special_start_page_identifiers": self.special_editor.get_values(),
            "non_start_page_patterns": self.nonstart_editor.get_values(),
            # Keep weak markers from default config unless you expose them
            "start_page_weak_markers": self.config["start_page_weak_markers"],
        }

    # -------------------------
    # FILE PICKERS
    # -------------------------

    def select_input(self):
        path = filedialog.askdirectory(title="Select input folder")
        if path:
            self.input_path.set(path)

    def select_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_path.set(path)

    # -------------------------
    # LOGGING
    # -------------------------

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    # -------------------------
    # RUN PROCESS
    # -------------------------

    def run_processing(self):
        input_folder = self.input_path.get().strip()
        output_folder = self.output_path.get().strip()

        if not input_folder:
            messagebox.showwarning("Missing", "Select input folder")
            return

        if not output_folder:
            messagebox.showwarning("Missing", "Select output folder")
            return

        config = self.collect_config()

        self.log("Starting processing...")
        self.log(f"Input: {input_folder}")
        self.log(f"Output: {output_folder}")

        try:
            process_all(input_folder, output_folder, config)
            self.log("Finished successfully.")
            messagebox.showinfo("Done", "Processing complete")

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))


# -------------------------
# ENTRY
# -------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateGUI(root)
    root.mainloop()