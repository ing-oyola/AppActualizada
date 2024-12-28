import os
import sys
from typing import List, Optional

class BaseApp:
    REQUIRED_FILES = [
        "db_maestrospdv.xlsx",
        "Form.Maestro.Neg.xlsx"
    ]

    @classmethod
    def get_base_path(cls) -> str:
        """Obtiene la ruta base de la aplicación"""
        if getattr(sys, 'frozen', False):
            # Si es un ejecutable empaquetado
            return os.path.dirname(sys.executable)
        else:
            # Si está en desarrollo
            return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    @classmethod
    def get_data_path(cls, filename: str) -> str:
        """Obtiene la ruta completa de un archivo en el directorio data"""
        data_dir = os.path.join(cls.get_base_path(), "data")
        file_path = os.path.join(data_dir, filename)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"No se encontró el archivo {filename}")
        return file_path
    
    @classmethod
    def get_version(cls) -> str:
        """Lee la versión actual desde version.txt"""
        version_path = os.path.join(cls.get_base_path(), "version.txt")
        try:
            with open(version_path, "r") as f:
                return f.read().strip()
        except FileNotFoundError:
            return "v0.0.0"
    
    @classmethod
    def verify_data_files(cls) -> List[str]:
        """Verifica que existan todos los archivos requeridos"""
        missing_files = []
        data_dir = os.path.join(cls.get_base_path(), "data")
        
        # Verificar que exista el directorio data
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        
        # Verificar cada archivo requerido
        for filename in cls.REQUIRED_FILES:
            file_path = os.path.join(data_dir, filename)
            if not os.path.exists(file_path):
                missing_files.append(filename)
        
        return missing_files
    
    @classmethod
    def is_production(cls) -> bool:
        """Determina si la aplicación está corriendo en producción"""
        return getattr(sys, 'frozen', False)
    
    @classmethod
    def get_app_file_path(cls) -> Optional[str]:
        """Obtiene la ruta del ejecutable en producción"""
        if cls.is_production():
            return sys.executable
        return None