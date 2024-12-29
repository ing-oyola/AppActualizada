from base_app import BaseApp
from temp_handler import temp_handler
import requests
import os
import sys
import shutil
from packaging import version
import subprocess
import zipfile

import logging
import os
import tkinter as tk
from tkinter import ttk

class AutoUpdater:
    def __init__(self):
        self.github_user = "ing-oyola"
        self.repo_name = "AppActualizada"
        self.current_version = BaseApp.get_version()
        self.github_api_url = f"https://api.github.com/repos/{self.github_user}/{self.repo_name}/releases/latest"
        self.base_dir = BaseApp.get_base_path()
        self.temp_dir = temp_handler.get_temp_dir()
        
        # Configurar logging
        log_path = os.path.join(self.base_dir, 'update.log')
        self.logger = logging.getLogger('updater')
        self.logger.setLevel(logging.INFO)
        handler = logging.FileHandler(log_path)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(handler)

    def download_update(self, version, progress_window=None):
        """Descarga la nueva versión"""
        try:
            self.logger.info(f"Iniciando descarga de versión {version}")
            
            # Actualizar interfaz si existe
            status_label = None
            progress_bar = None
            if progress_window:
                status_label = progress_window.nametowidget("status_label")
                progress_bar = progress_window.nametowidget("progress_bar")

            def update_status(message):
                self.logger.info(message)
                if status_label:
                    status_label.config(text=message)
                    progress_window.update()

            update_status("Conectando con GitHub...")
            response = requests.get(self.github_api_url)
            assets = response.json()['assets']
            
            for asset in assets:
                if asset['name'].endswith('.zip'):
                    download_url = asset['browser_download_url']
                    update_status(f"Descargando desde {download_url}")
                    
                    r = requests.get(download_url, stream=True)
                    total_size = int(r.headers.get('content-length', 0))
                    
                    if progress_bar:
                        progress_bar['mode'] = 'determinate'
                        progress_bar['maximum'] = total_size

                    update_dir = os.path.join(self.temp_dir, "update")
                    os.makedirs(update_dir, exist_ok=True)
                    update_zip = os.path.join(update_dir, "update.zip")

                    downloaded = 0
                    with open(update_zip, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if progress_bar:
                                    progress_bar['value'] = downloaded
                                    progress_window.update()
                                update_status(f"Descargando... {(downloaded/total_size)*100:.1f}%")

                    update_status("Extrayendo archivos...")
                    exe_path = self.extract_update(update_zip)
                    
                    if exe_path:
                        update_status("Descarga completada exitosamente")
                        self.logger.info("Descarga y extracción completadas con éxito")
                        return exe_path
                    else:
                        update_status("Error en la extracción")
                        self.logger.error("Error durante la extracción")
                        return None

            self.logger.error("No se encontró archivo ZIP en el release")
            return None

        except Exception as e:
            self.logger.error(f"Error en la descarga: {str(e)}", exc_info=True)
            if progress_window:
                progress_window.destroy()
            return None
         
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
                        print(f"Backup realizado de: {filename}")
        except Exception as e:
            print(f"Error al hacer backup: {e}")

    def extract_update(self, zip_path):
        """Extrae la actualización y retorna la ruta al ejecutable"""
        try:
            extract_dir = os.path.join(os.path.dirname(zip_path), "extracted")
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir)
            
            print(f"Extrayendo en: {extract_dir}")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # Eliminar el ZIP
            os.remove(zip_path)
            
            # Buscar el ejecutable
            exe_path = os.path.join(extract_dir, "AppActualizada.exe")
            if os.path.exists(exe_path):
                print(f"Ejecutable encontrado: {exe_path}")
                return exe_path
            else:
                print(f"No se encontró el ejecutable en: {exe_path}")
                return None
            
        except Exception as e:
            print(f"Error al extraer actualización: {e}")
            import traceback
            traceback.print_exc()
            return None

    def apply_update(self, new_exe_path):
        """Aplica la actualización"""
        try:
            if not new_exe_path or not os.path.exists(new_exe_path):
                print(f"Ejecutable no encontrado: {new_exe_path}")
                return False
            
            print("Preparando actualización...")
            batch_path = os.path.join(self.temp_dir, "updater.bat")
            current_exe = BaseApp.get_app_file_path()
            
            batch_content = f'''
@echo on
echo Iniciando actualización...
timeout /t 2 /nobreak
echo Cerrando aplicación actual...
taskkill /F /IM AppActualizada.exe /T
echo Eliminando versión anterior...
del "{current_exe}"
echo Copiando nueva versión...
copy "{new_exe_path}" "{current_exe}"
echo Limpiando temporales...
rmdir /s /q "{os.path.dirname(os.path.dirname(new_exe_path))}"
echo Iniciando nueva versión...
start "" "{current_exe}"
echo Eliminando batch...
del "%~f0"
'''
            print("Creando batch de actualización...")
            with open(batch_path, "w") as batch:
                batch.write(batch_content)
            
            print("Ejecutando actualización...")
            subprocess.Popen(batch_path, shell=True)
            print("Cerrando aplicación...")
            sys.exit()
            
        except Exception as e:
            print(f"Error al aplicar actualización: {e}")
            self.restore_backup()
            import traceback
            traceback.print_exc()
            return False