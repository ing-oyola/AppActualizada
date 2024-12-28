from base_app import BaseApp
from temp_handler import temp_handler
import requests
import os
import sys
import shutil
from packaging import version
import subprocess
import zipfile

class AutoUpdater:
    def __init__(self):
        self.github_user = "ing-oyola"
        self.repo_name = "AppActualizada"
        self.current_version = BaseApp.get_version()
        self.github_api_url = f"https://api.github.com/repos/{self.github_user}/{self.repo_name}/releases/latest"
        self.base_dir = BaseApp.get_base_path()
        self.temp_dir = temp_handler.get_temp_dir()
    
    def check_for_updates(self):
        """Verifica si hay actualizaciones disponibles"""
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
        """Descarga la nueva versión"""
        try:
            response = requests.get(self.github_api_url)
            assets = response.json()['assets']
            
            for asset in assets:
                if asset['name'].endswith('.zip'):
                    download_url = asset['browser_download_url']
                    print("Descargando actualización...")
                    r = requests.get(download_url)
                    
                    # Usar la ruta base de BaseApp
                    update_file = os.path.join(self.temp_dir, "update_temp.zip")
                    with open(update_file, 'wb') as f:
                        f.write(r.content)
                    
                    # Hacer backup de los archivos de datos
                    self.backup_data_files()
                    
                    # Extraer el ZIP
                    self.extract_update(update_file)
                    return os.path.join(self.base_dir, "update_temp", "AppActualizada.exe")
            return None
            
        except Exception as e:
            print(f"Error al descargar actualización: {e}")
            return None

    def backup_data_files(self):
        """Hace una copia de seguridad de los archivos de datos"""
        backup_dir = os.path.join(self.base_dir, "backup_data")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        try:
            data_dir = os.path.join(self.base_dir, "data")
            if os.path.exists(data_dir):
                # Copiar cada archivo requerido
                for filename in BaseApp.REQUIRED_FILES:
                    src_file = os.path.join(data_dir, filename)
                    if os.path.exists(src_file):
                        dst_file = os.path.join(backup_dir, filename)
                        shutil.copy2(src_file, dst_file)
        except Exception as e:
            print(f"Error al hacer backup: {e}")

    def restore_backup(self):
        """Restaura los archivos de datos desde el backup"""
        backup_dir = os.path.join(self.base_dir, "backup_data")
        data_dir = os.path.join(self.base_dir, "data")
        
        if os.path.exists(backup_dir):
            for filename in BaseApp.REQUIRED_FILES:
                backup_file = os.path.join(backup_dir, filename)
                if os.path.exists(backup_file):
                    shutil.copy2(backup_file, os.path.join(data_dir, filename))

    def extract_update(self, zip_path):
        """Extrae la actualización"""
        extract_path = os.path.join(self.temp_dir, "update_temp")
        if os.path.exists(extract_path):
            shutil.rmtree(extract_path)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        
        os.remove(zip_path)

    def apply_update(self, update_file):
        """Aplica la actualización"""
        try:
            batch_path = os.path.join(self.temp_dir, "updater.bat")
            executable_path = BaseApp.get_app_file_path()
            
            with open(batch_path, "w") as batch:
                batch.write(f'''
@echo off
timeout /t 2 /nobreak
del "{executable_path}"
move /y "{update_file}" "{executable_path}"
rmdir /s /q "{os.path.dirname(update_file)}"
start "" "{executable_path}"
del "%~f0"
                ''')
            
            subprocess.Popen(batch_path)
            sys.exit()
            
        except Exception as e:
            print(f"Error al aplicar actualización: {e}")
            self.restore_backup()  # Restaurar backup si falla la actualización
            if os.path.exists(update_file):
                os.remove(update_file)