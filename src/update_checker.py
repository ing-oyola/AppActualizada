import os
import requests
import sys
from packaging import version
import subprocess

class AutoUpdater:
    def __init__(self):
        self.github_user = "ing-oyola"  # Tu usuario de GitHub
        self.repo_name = "AppActualizada"  # Tu repositorio
        
        # Lee la versión actual del archivo version.txt
        with open("version.txt", "r") as f:
            self.current_version = f.read().strip()
            
        self.github_api_url = f"https://api.github.com/repos/{self.github_user}/{self.repo_name}/releases/latest"
    
    def check_for_updates(self):
        try:
            response = requests.get(self.github_api_url)
            response.raise_for_status()
            latest_version = response.json()['tag_name'].replace('v', '')
            
            if version.parse(latest_version) > version.parse(self.current_version):
                return True, latest_version
            return False, latest_version
            
        except Exception as e:
            print(f"Error al verificar actualizaciones: {e}")
            return False, None
    
    def download_update(self, version):
        try:
            response = requests.get(self.github_api_url)
            assets = response.json()['assets']
            
            for asset in assets:
                if asset['name'].endswith('.exe'):
                    download_url = asset['browser_download_url']
                    print("Descargando actualización...")
                    r = requests.get(download_url)
                    
                    update_file = "update_temp.exe"
                    with open(update_file, 'wb') as f:
                        f.write(r.content)
                    
                    return update_file
            return None
            
        except Exception as e:
            print(f"Error al descargar actualización: {e}")
            return None

    def apply_update(self, update_file):
        try:
            with open("updater.bat", "w") as batch:
                batch.write(f'''
@echo off
timeout /t 2 /nobreak
del "{sys.executable}"
move /y "{update_file}" "{sys.executable}"
start "" "{sys.executable}"
del "%~f0"
                ''')
            
            subprocess.Popen("updater.bat")
            sys.exit()
            
        except Exception as e:
            print(f"Error al aplicar actualización: {e}")
            if os.path.exists(update_file):
                os.remove(update_file)