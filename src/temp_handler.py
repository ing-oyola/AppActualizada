import os
import shutil
import tempfile
import atexit
import time

class TempFileHandler:
    def __init__(self):
        self.temp_dirs = set()
        atexit.register(self.cleanup_temp_files)
    
    def get_temp_dir(self):
        """Crea un directorio temporal y registra su ubicación"""
        temp_dir = tempfile.mkdtemp()
        self.temp_dirs.add(temp_dir)
        return temp_dir
    
    def cleanup_temp_files(self):
        """Limpia los archivos temporales al cerrar la aplicación"""
        for temp_dir in self.temp_dirs.copy():
            try:
                if os.path.exists(temp_dir):
                    # Esperar un momento antes de intentar eliminar
                    time.sleep(0.5)
                    shutil.rmtree(temp_dir, ignore_errors=True)
                self.temp_dirs.remove(temp_dir)
            except Exception:
                pass  # Ignorar errores al limpiar archivos temporales

# Crear una instancia global
temp_handler = TempFileHandler()