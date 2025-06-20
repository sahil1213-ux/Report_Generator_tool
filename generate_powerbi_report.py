import os
import shutil
import subprocess
import customtkinter as ctk # type: ignore
from tkinter import filedialog
from datetime import datetime

# Constants
PBI_TEMPLATE_PATH = "template/generate_dash.pbit"
INPUT_CSV_PATH = "data/input.csv"
POWERBI_EXECUTABLE = r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin\PBIDesktop.exe"  # Update path if needed

# GUI Setup
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Power BI AutoReport Generator")
app.geometry("600x400")

# Status log output
log_box = ctk.CTkTextbox(app, width=550, height=200)
log_box.pack(pady=20)

def log(message):
    timestamp = datetime.now().strftime("[%H:%M:%S]")
    log_box.insert("end", f"{timestamp} {message}\n")
    log_box.see("end")

def select_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        log("No file selected.")
        return

    try:
        shutil.copyfile(file_path, INPUT_CSV_PATH)
        log(f"✅ Replaced input.csv with: {os.path.basename(file_path)}")
    except Exception as e:
        log(f"❌ Error copying file: {e}")

def launch_powerbi():
    if not os.path.exists(PBI_TEMPLATE_PATH):
        log("❌ Power BI template not found!")
        return

    try:
        subprocess.Popen([POWERBI_EXECUTABLE, PBI_TEMPLATE_PATH], shell=True)
        log("🚀 Power BI launched with template.")
    except Exception as e:
        log(f"❌ Failed to open Power BI: {e}")

# Buttons
btn_select = ctk.CTkButton(app, text="Select CSV", command=select_csv)
btn_select.pack(pady=10)

btn_launch = ctk.CTkButton(app, text="Launch Power BI", command=launch_powerbi)
btn_launch.pack(pady=10)

app.mainloop()