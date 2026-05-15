import os
import urllib.request
import zipfile
import subprocess
import shutil

APP_DIR = "Offline_App"
PYTHON_DIR = os.path.join(APP_DIR, "python")
PYTHON_ZIP = "python-3.11.9-embed-amd64.zip"
PYTHON_URL = f"https://www.python.org/ftp/python/3.11.9/{PYTHON_ZIP}"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"
GET_PIP_FILE = "get-pip.py"

def download_file(url, dest):
    print(f"Downloading {url} to {dest}...")
    urllib.request.urlretrieve(url, dest)
    print("Download complete.")

def build_offline_app():
    # 1. Create Directories
    if not os.path.exists(APP_DIR):
        os.makedirs(APP_DIR)
    if not os.path.exists(PYTHON_DIR):
        os.makedirs(PYTHON_DIR)

    # 2. Download and Extract Python Embeddable
    zip_path = os.path.join(APP_DIR, PYTHON_ZIP)
    if not os.path.exists(zip_path):
        download_file(PYTHON_URL, zip_path)
    
    print("Extracting Python...")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(PYTHON_DIR)
    
    # 3. Patch python311._pth to enable site-packages
    pth_file = os.path.join(PYTHON_DIR, "python311._pth")
    with open(pth_file, "r") as f:
        content = f.read()
    content = content.replace("#import site", "import site")
    with open(pth_file, "w") as f:
        f.write(content)

    # 4. Download get-pip.py
    pip_path = os.path.join(APP_DIR, GET_PIP_FILE)
    if not os.path.exists(pip_path):
        download_file(GET_PIP_URL, pip_path)

    # 5. Install pip
    print("Installing pip...")
    python_exe = os.path.join(PYTHON_DIR, "python.exe")
    subprocess.run([python_exe, pip_path], check=True)

    # 6. Install requirements
    print("Installing requirements...")
    pip_exe = os.path.join(PYTHON_DIR, "Scripts", "pip.exe")
    # if pip.exe doesn't exist, try python.exe -m pip
    subprocess.run([python_exe, "-m", "pip", "install", "-r", "requirements.txt"], check=True)

    # 7. Copy Project Files
    print("Copying project files...")
    items_to_copy = [
        "app.py",
        "config.json",
        "Column_Mapper.xlsx",
        "requirements.txt",
        "scripts",
        "data",
        "dashboard",
        "backup_v1",
        "DOCUMENTATION.md"
    ]

    for item in items_to_copy:
        src = item
        dst = os.path.join(APP_DIR, item)
        if not os.path.exists(src):
            print(f"Warning: {src} does not exist, skipping.")
            continue
        
        if os.path.isdir(src):
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)

    # 8. Create Start Application.bat
    print("Creating launcher...")
    bat_content = """@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Write-Host 'Starting DB Merging Application...' -ForegroundColor Cyan; try { .\\python\\python.exe -m streamlit run app.py } catch { Write-Host 'An error occurred:' -ForegroundColor Red; $_.Exception.Message }; Write-Host 'Press any key to exit...' -ForegroundColor Yellow; $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
"""
    bat_path = os.path.join(APP_DIR, "Start Application.bat")
    with open(bat_path, "w") as f:
        f.write(bat_content)
    
    print("--------------------------------------------------")
    print("Build complete! The 'Offline_App' folder is ready.")
    print("--------------------------------------------------")

if __name__ == "__main__":
    build_offline_app()
