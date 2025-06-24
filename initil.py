import subprocess
import sys

# Check if a python library is installed
def is_module_installed(module_name):
    try:
        __import__(module_name)
        return True
    except ImportError:
        return False

# Install a python library using pip
def install_module(module_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])

# Check and install pandas if not present
if not is_module_installed("pandas"):
    install_module("pandas")

# Check and install openpyxl if not present
if not is_module_installed("openpyxl"):
    install_module("openpyxl")

# Check and install tkinter if not present
if not is_module_installed("tkinter"):
    install_module("tkinter")

import tkinter as tk
from tkinter import font
from tkinter import filedialog
import tkinter.messagebox as msg

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.geometry("500x300")
        self.pack()

        self.font = font.Font(size=20)

        # Initialize instance variables to store file paths as memory
        self.dir_path = ""
        self.hod_rank_file = ""
        self.mo_rank_file = ""
        self.mopex_staff_listing = ""
        self.gdfm_file = ""
        self.removed_mo_file = ""

        self.create_widgets()

    def create_widgets(self):
        #Creating buttons to open up subwindows
        self.hod_rank_compiler_button = tk.Button(self, text="Compile HOD Ranking", command=self.open_hod_rank_compiler_window, font=self.font)
        self.hod_rank_compiler_button.pack(pady=20)

        self.run_button = tk.Button(self, text="Generate Match", command=self.open_match_algo_window, font=self.font)
        self.run_button.pack(pady=20)

        self.remove_mos_button = tk.Button(self, text="Remove MOs + Blacklist", command=self.open_remove_mos_window, font=self.font)
        self.remove_mos_button.pack(pady=20)

        #Formatting to center buttons and text
        self.columnconfigure(0, weight=1)

    def open_hod_rank_compiler_window(self):
        self.hod_rank_compiler_window = tk.Toplevel(self.master)
        self.hod_rank_compiler_window.title("Compile HOD Ranking")
        self.hod_rank_compiler_window.geometry("1200x300")

        # Raw Data Folder selection
        tk.Label(self.hod_rank_compiler_window, text="Raw Data Folder:", font=self.font).grid(row=0, column=0)
        self.raw_data_entry = tk.Entry(self.hod_rank_compiler_window, font=self.font, width=50)
        self.raw_data_entry.grid(row=0, column=1)
        tk.Button(self.hod_rank_compiler_window, text="Browse", command=self.browse_raw_data_folder, font=self.font).grid(row=0, column=2)

        # HOD Ranking File selection
        tk.Label(self.hod_rank_compiler_window, text="HOD Ranking File:", font=self.font).grid(row=1, column=0)
        self.hod_rank_entry = tk.Entry(self.hod_rank_compiler_window, font=self.font, width=50)
        self.hod_rank_entry.grid(row=1, column=1)
        tk.Button(self.hod_rank_compiler_window, text="Browse", command=self.browse_hod_rank_file, font=self.font).grid(row=1, column=2)

        # MOPEX Staff Listing File selection
        tk.Label(self.hod_rank_compiler_window, text="MOPEX Staff Listing File:", font=self.font).grid(row=2, column=0)
        self.mopex_stafflist_entry = tk.Entry(self.hod_rank_compiler_window, font=self.font, width=50)
        self.mopex_stafflist_entry.grid(row=2, column=1)
        tk.Button(self.hod_rank_compiler_window, text="Browse", command=self.browse_mopex_staff_listing, font=self.font).grid(row=2, column=2)

        # Run button for HOD Rank Compiler
        tk.Button(self.hod_rank_compiler_window, text="Run", command=self.run_hod_rank_compiler, font=self.font).grid(row=3, column=1, pady=20)

    def browse_raw_data_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.raw_data_entry.delete(0, tk.END)
            self.raw_data_entry.insert(0, folder_selected)
            self.dir_path = folder_selected  # Store in memory

    def browse_hod_rank_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.hod_rank_entry.delete(0, tk.END)
            self.hod_rank_entry.insert(0, file_selected)
            self.hod_rank_file = file_selected  # Store in memory
    
    def browse_mopex_staff_listing(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.mopex_stafflist_entry.delete(0, tk.END)
            self.mopex_stafflist_entry.insert(0, file_selected)
            self.mopex_staff_listing = file_selected  # Store in memory

    def run_hod_rank_compiler(self):
        if not self.dir_path:
            msg.showerror("Error", "Raw Data Folder is required!")
        elif not self.hod_rank_file:
            msg.showerror("Error", "HOD Rank file is required!")
        elif not self.mopex_staff_listing:
            msg.showerror("Error, MOPEX staff listing file is required!")
        else:
            try:
                subprocess.check_call(["python", "HOD_Rank_Compiler.py", "--dir", self.dir_path, "--HODRankFile", self.hod_rank_file, "--MOPEXStaffList", self.mopex_staff_listing])
                msg.showinfo("Success", "The script has completed successfully!")
            except subprocess.CalledProcessError:
                msg.showerror("Error", "An error has occurred, please recheck the files again.")
    
    def open_match_algo_window(self):
        self.match_algo_window = tk.Toplevel(self.master)
        self.match_algo_window.title("Generate Match")
        self.match_algo_window.geometry("1200x300")

        # HOD Ranking File selection
        tk.Label(self.match_algo_window, text="HOD Ranking File:", font=self.font).grid(row=0, column=0)
        self.hod_rank_entry = tk.Entry(self.match_algo_window, font=self.font, width=50)
        self.hod_rank_entry.grid(row=0, column=1)
        tk.Button(self.match_algo_window, text="Browse", command=self.browse_hod_rank_file_match, font=self.font).grid(row=0, column=2)

        # MO Ranking File selection
        tk.Label(self.match_algo_window, text="MO Ranking File:", font=self.font).grid(row=1, column=0)
        self.mo_rank_entry = tk.Entry(self.match_algo_window, font=self.font, width=50)
        self.mo_rank_entry.grid(row=1, column=1)
        tk.Button(self.match_algo_window, text="Browse", command=self.browse_mo_rank_file_match, font=self.font).grid(row=1, column=2)

        # GDFM Intake File selection
        tk.Label(self.match_algo_window, text="GDFM Intake File:", font=self.font).grid(row=2, column=0)
        self.gdfm_entry = tk.Entry(self.match_algo_window, font=self.font, width=50)
        self.gdfm_entry.grid(row=2, column=1)
        tk.Button(self.match_algo_window, text="Browse", command=self.browse_gdfm_file, font=self.font).grid(row=2, column=2)

        # Run button for Match Algo
        tk.Button(self.match_algo_window, text="Run", command=self.run_match_algo, font=self.font).grid(row=3, column=1, pady=20)

    def browse_hod_rank_file_match(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.hod_rank_entry.delete(0, tk.END)
            self.hod_rank_entry.insert(0, file_selected)
            self.hod_rank_file = file_selected  # Store in memory

    def browse_mo_rank_file_match(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.mo_rank_entry.delete(0, tk.END)
            self.mo_rank_entry.insert(0, file_selected)
            self.mo_rank_file = file_selected  # Store in memory

    def browse_gdfm_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.gdfm_entry.delete(0, tk.END)
            self.gdfm_entry.insert(0, file_selected)
            self.gdfm_file = file_selected  # Store in memory

    def run_match_algo(self):
        if not self.hod_rank_file:
            msg.showerror("Error", "HOD Rank file is required!")
        elif not self.mo_rank_file:
            msg.showerror("Error", "MO Rank file is required!")
        elif not self.gdfm_file:
            msg.showerror("Error", "GDFM Intake file is required!")
        else:
            try:
                subprocess.check_call(["python", "Match_Algo.py", 
                                      "--HODRankFile", self.hod_rank_file, 
                                      "--MORankFile", self.mo_rank_file, 
                                      "--GDFM", self.gdfm_file])
                msg.showinfo("Success", "The script has completed successfully!")
            except subprocess.CalledProcessError:
                msg.showerror("Error", "An error has occurred, please recheck the files again.")

    def open_remove_mos_window(self):
        self.remove_mos_window = tk.Toplevel(self.master)
        self.remove_mos_window.title("Remove MOs + Blacklist")
        self.remove_mos_window.geometry("1200x300")

        # MO Ranking File selection
        tk.Label(self.remove_mos_window, text="MO Ranking File:", font=self.font).grid(row=0, column=0)
        self.mo_rank_entry = tk.Entry(self.remove_mos_window, font=self.font, width=50)
        self.mo_rank_entry.grid(row=0, column=1)
        tk.Button(self.remove_mos_window, text="Browse", command=self.browse_mo_rank_file_remove, font=self.font).grid(row=0, column=2)

        # Remove MOs File selection
        tk.Label(self.remove_mos_window, text="Remove MOs File:", font=self.font).grid(row=1, column=0)
        self.removed_mo_entry = tk.Entry(self.remove_mos_window, font=self.font, width=50)
        self.removed_mo_entry.grid(row=1, column=1)
        tk.Button(self.remove_mos_window, text="Browse", command=self.browse_removed_mo_file, font=self.font).grid(row=1, column=2)

        # Run button for Remove MOs
        tk.Button(self.remove_mos_window, text="Run", command=self.run_remove_mos, font=self.font).grid(row=2, column=1, pady=20)

    def browse_mo_rank_file_remove(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.mo_rank_entry.delete(0, tk.END)
            self.mo_rank_entry.insert(0, file_selected)
            self.mo_rank_file = file_selected

    def browse_removed_mo_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_selected:
            self.removed_mo_entry.delete(0, tk.END)
            self.removed_mo_entry.insert(0, file_selected)
            self.removed_mo_file = file_selected

    def run_remove_mos(self):
        if not self.mo_rank_file:
            msg.showerror("Error", "MO Rank file is required!")
        elif not self.removed_mo_file:
            msg.showerror("Error", "Remove MOs file is required!")
        else:
            try:
                subprocess.check_call(["python", "Remove_MOs.py",
                                      "--MORankFile", self.mo_rank_file,
                                      "--RemovedMOFile", self.removed_mo_file])
                msg.showinfo("Success", "The script has completed successfully!")
            except subprocess.CalledProcessError:
                msg.showerror("Error", "An error has occurred, please recheck the files again.")

root = tk.Tk()
app = Application(master=root)
app.mainloop()