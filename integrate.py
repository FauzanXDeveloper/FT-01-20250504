import sys
import os
import subprocess
import customtkinter
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

customtkinter.set_default_color_theme("Themes/lavender.json")  # Options: "System", "Dark", "Light"

def run_ccris():
    # Use sys.executable to ensure we launch with the current interpreter.
    script_path = os.path.join(os.getcwd(), "main.py")
    subprocess.Popen([sys.executable, script_path])

def run_ctos():
    script_path = os.path.join(os.getcwd(), "app.py")
    subprocess.Popen([sys.executable, script_path])

app = ctk.CTk()
app.title("Report Launcher")
app.geometry("300x200")

btn_ccris = ctk.CTkButton(app, text="Run CCRIS", command=run_ccris)
btn_ccris.pack(padx=20, pady=(40,10))

btn_ctos = ctk.CTkButton(app, text="Run CTOS", command=run_ctos)
btn_ctos.pack(padx=20, pady=10)

app.mainloop()