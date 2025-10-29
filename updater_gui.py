import os
import requests
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess

# Configurações do repositório
GITHUB_USER = "Kvsl11"  # <-- substitua pelo seu nome de usuário
REPO_NAME = "Auto-Ficha-OPE"

# URLs base
BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/main/"
VERSION_FILE = "version.txt"
SCRIPT_FILE = "main.py"

# Caminho local
LOCAL_SCRIPT = os.path.join(os.getcwd(), SCRIPT_FILE)
LOCAL_VERSION_FILE = os.path.join(os.getcwd(), "version.txt")

def get_remote_version():
    try:
        r = requests.get(BASE_URL + VERSION_FILE, timeout=10)
        if r.status_code == 200:
            return r.text.strip()
    except Exception as e:
        print(f"Erro ao obter versão remota: {e}")
    return None

def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, "r") as f:
            return f.read().strip()
    return None

def download_new_version(version):
    r = requests.get(BASE_URL + f"Version_{version}.py", stream=True)
    total = int(r.headers.get('content-length', 0))
    with open(LOCAL_SCRIPT, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    with open(LOCAL_VERSION_FILE, "w") as f:
        f.write(version)

def check_update():
    local = get_local_version()
    remote = get_remote_version()
    if remote and local != remote:
        if messagebox.askyesno("Atualização disponível", f"Versão {remote} encontrada. Deseja atualizar agora?"):
            download_new_version(remote)
            messagebox.showinfo("Atualizado", f"App atualizado para a versão {remote}.")
            subprocess.Popen(["python", LOCAL_SCRIPT])
            exit()
    else:
        subprocess.Popen(["python", LOCAL_SCRIPT])
        exit()

root = tk.Tk()
root.withdraw()
check_update()
