
import time
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
        
        # Solo configurar logging en desarrollo
        if not BaseApp.is_production():
            log_path = os.path.join(self.base_dir, 'update.log')
            self.logger = logging.getLogger('updater')
            self.logger.setLevel(logging.INFO)
            handler = logging.FileHandler(log_path)
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            self.logger.addHandler(handler)
        else:
            # Logger nulo para producción
            self.logger = logging.getLogger('null')
            self.logger.addHandler(logging.NullHandler())

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
            self.logger.info(f"Assets encontrados: {[asset['name'] for asset in assets]}")
            
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

                    # Descargar el ZIP
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

                    # Verificar contenido del ZIP
                    update_status("Verificando contenido del ZIP...")
                    with zipfile.ZipFile(update_zip, 'r') as zip_ref:
                        files = zip_ref.namelist()
                        self.logger.info("Contenido del ZIP:")
                        for file in files:
                            self.logger.info(f"- {file}")
                        
                        # Verificar version.txt dentro del ZIP
                        if 'version.txt' in files:
                            with zip_ref.open('version.txt') as version_file:
                                zip_version = version_file.read().decode('utf-8').strip()
                                self.logger.info(f"Versión en ZIP: {zip_version}")
                        else:
                            self.logger.warning("No se encontró version.txt en el ZIP")

                    # Extraer archivos
                    update_status("Extrayendo archivos...")
                    extract_dir = os.path.join(update_dir, "extracted")
                    if os.path.exists(extract_dir):
                        shutil.rmtree(extract_dir)
                    os.makedirs(extract_dir)
                    
                    with zipfile.ZipFile(update_zip, 'r') as zip_ref:
                        zip_ref.extractall(extract_dir)
                    
                    # Verificar archivos extraídos
                    exe_path = os.path.join(extract_dir, "AppActualizada.exe")
                    version_path = os.path.join(extract_dir, "version.txt")
                    
                    if os.path.exists(exe_path):
                        self.logger.info(f"Ejecutable encontrado: {exe_path}")
                        
                        # Modificar esta parte del código
                        if os.path.exists(version_path):
                            with open(version_path, 'r') as f:
                                extracted_version = f.read().strip()
                                self.logger.info(f"Versión extraída: {extracted_version}")
                                # Normalizar las versiones para comparación
                                extracted_ver = extracted_version.replace('v', '')
                                expected_ver = version.replace('v', '')
                                if extracted_ver != expected_ver:
                                    self.logger.warning(f"¡Advertencia! Versión extraída ({extracted_version}) no coincide con la esperada (v{version})")
                        
                        update_status("Descarga completada exitosamente")
                        return exe_path
                    else:
                        update_status("Error: No se encontró el ejecutable")
                        self.logger.error(f"Ejecutable no encontrado en: {exe_path}")
                        return None

            update_status("Error: No se encontró el archivo ZIP en el release")
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
                self.logger.error(f"Ejecutable no encontrado: {new_exe_path}")
                return False
            
            self.logger.info(f"Preparando actualización desde: {new_exe_path}")
            current_exe = BaseApp.get_app_file_path()
            current_dir = os.path.dirname(current_exe)
            new_dir = os.path.dirname(new_exe_path)
            self.logger.info(f"Ejecutable actual: {current_exe}")
            
            batch_path = os.path.join(current_dir, "updater.bat")
            self.logger.info(f"Creando batch en: {batch_path}")
            
            batch_content = f'''
    @echo off
    cd /d "{current_dir}"

    :: Cerrar la aplicación actual
    taskkill /F /IM AppActualizada.exe >nul 2>&1
    timeout /t 3 /nobreak >nul

    :: Eliminar archivos anteriores
    del /F /Q "{current_exe}" >nul 2>&1
    del /F /Q "{os.path.join(current_dir, 'version.txt')}" >nul 2>&1

    :: Copiar nuevos archivos
    copy /Y "{new_exe_path}" "{current_exe}" >nul 2>&1
    copy /Y "{os.path.join(new_dir, 'version.txt')}" "{os.path.join(current_dir, 'version.txt')}" >nul 2>&1

    :: Verificar la copia
    if exist "{current_exe}" (
        if not exist "{os.path.join(current_dir, 'version.txt')}" (
            exit /b 1
        )
    ) else (
        exit /b 1
    )

    :: Limpiar temporales
    rmdir /S /Q "{os.path.dirname(os.path.dirname(new_exe_path))}" >nul 2>&1

    :: Iniciar nueva versión
    start /b "" "{current_exe}"
    (goto) 2>nul & del "%~f0"
    '''
            # Escribir el batch
            self.logger.info("Escribiendo archivo batch...")
            with open(batch_path, 'w', encoding='utf-8') as batch:
                batch.write(batch_content)
            
            # Ejecutar el batch
            self.logger.info("Ejecutando batch de actualización...")
            CREATE_NO_WINDOW = 0x08000000
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            subprocess.Popen(
                f'cmd /c "{batch_path}"',
                creationflags=CREATE_NO_WINDOW,
                startupinfo=startupinfo,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            
            # Esperar un momento antes de cerrar
            self.logger.info("Esperando antes de cerrar...")
            time.sleep(3)
            self.logger.info("Cerrando aplicación para completar actualización...")
            sys.exit(0)
            
        except Exception as e:
            self.logger.error(f"Error al aplicar actualización: {str(e)}", exc_info=True)
            return False
