import os
import json
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import winshell
from win32com.client import Dispatch

# ------------------------------------------------------------
# Step 1: Ask for installation directory FIRST
# ------------------------------------------------------------

def ask_install_path():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Installation Setup",
        "Please select an installation directory for the Insurgency Sandstorm VR Express Setup"
    )

    path = filedialog.askdirectory(
        title="Select installation directory for Sandstorm VR Express Setup"
    )

    root.destroy()

    if not path:
        messagebox.showerror(
            "Installation aborted",
            "You must select an installation directory."
        )
        sys.exit(1)

    return os.path.abspath(os.path.join(path, "sandstormVRsetup"))


INSTALL_DIR = ask_install_path()
os.makedirs(INSTALL_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(INSTALL_DIR, "sandstorm_vr_setup.json")

# ------------------------------------------------------------
# Config handling
# ------------------------------------------------------------

DEFAULT_PATHS = {
    "sandstorm_dir": "",
    "uevr_path": ""
}

def load_paths():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    else:
        save_paths(DEFAULT_PATHS)
        return DEFAULT_PATHS.copy()

def save_paths(paths):
    with open(CONFIG_FILE, "w") as f:
        json.dump(paths, f, indent=4)

# ------------------------------------------------------------
# UI callbacks
# ------------------------------------------------------------

def select_sandstorm_dir():
    path = filedialog.askdirectory(title="Select Insurgency Sandstorm Directory")
    if path:
        sandstorm_dir_var.set(path)

def select_uevr_path():
    path = filedialog.askopenfilename(title="Select UEVR path")
    if path:
        uevr_path_var.set(path)

def install():
    paths = {
        "sandstorm_dir": sandstorm_dir_var.get(),
        "uevr_path": uevr_path_var.get()
    }

    if not paths["sandstorm_dir"] or not paths["uevr_path"]:
        messagebox.showerror("Error", "Please set both paths before installing.")
        return

    save_paths(paths)

    try:
        rust_exe = install_rust_binary(INSTALL_DIR)
        create_desktop_shortcut(rust_exe)

        subprocess.Popen(
            [rust_exe],
            cwd=INSTALL_DIR,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        messagebox.showinfo(
            "Installation complete",
            "Sandstorm VR Setup installed successfully.\n\n"
            "You can now start it anytime via the desktop shortcut."
        )

        root.destroy()

    except Exception as e:
        messagebox.showerror("Installation failed", str(e))

def get_embedded_rust_exe():
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS  # PyInstaller
        return os.path.join(base, "payload", "sandstorm-vr-setup.exe")
    else:
        return os.path.join(os.path.dirname(__file__), "sandstorm-vr-setup.exe")


def install_rust_binary(install_dir):
    src = get_embedded_rust_exe()
    dst = os.path.join(install_dir, "sandstorm-vr-setup.exe")

    if not os.path.exists(src):
        messagebox.showerror("Installer Error", "sandstorm-vr-setup.exe not found.")
        sys.exit(1)

    shutil.copy2(src, dst)
    return dst

def create_desktop_shortcut(target_exe):
    desktop = winshell.desktop()
    shortcut_path = os.path.join(desktop, "Sandstorm VR Setup.lnk")

    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_exe
    shortcut.WorkingDirectory = os.path.dirname(target_exe)
    shortcut.IconLocation = target_exe
    shortcut.save()


# ------------------------------------------------------------
# Tkinter UI
# ------------------------------------------------------------

root = tk.Tk()
root.title("Insurgency Sandstorm VR Express Setup")
root.geometry("520x260")
root.resizable(False, False)

paths = load_paths()

sandstorm_dir_var = tk.StringVar(value=paths["sandstorm_dir"])
uevr_path_var = tk.StringVar(value=paths["uevr_path"])

tk.Label(root, text="Insurgency Sandstorm Directory (Your Steam Install):").pack(pady=(10, 0))
tk.Entry(root, textvariable=sandstorm_dir_var, width=65, state="readonly").pack()
tk.Button(root, text="Browse", command=select_sandstorm_dir).pack(pady=5)

tk.Label(root, text="UEVRInjector executable path \n(Download here: https://github.com/praydog/UEVR/releases/tag/1.05):").pack(pady=(10, 0))
tk.Entry(root, textvariable=uevr_path_var, width=65, state="readonly").pack()
tk.Button(root, text="Browse", command=select_uevr_path).pack(pady=5)

tk.Button(
    root,
    text="Install",
    command=install,
    bg="#007a33",
    fg="white",
    width=20
).pack(pady=20)

root.mainloop()
