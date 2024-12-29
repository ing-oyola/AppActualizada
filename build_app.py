import PyInstaller.__main__
import shutil
import os
import time
import zipfile
import subprocess

def retry_remove(path, max_attempts=3, delay=1):
    """Intenta eliminar un archivo o directorio varias veces"""
    for attempt in range(max_attempts):
        try:
            if os.path.isfile(path):
                os.unlink(path)
            else:
                shutil.rmtree(path)
            return True
        except PermissionError:
            if attempt < max_attempts - 1:
                print(f"Intento {attempt + 1}: No se pudo eliminar {path}, reintentando en {delay} segundos...")
                # Intentar liberar el archivo ejecutando el recolector de basura
                import gc
                gc.collect()
                time.sleep(delay)
            else:
                print(f"No se pudo eliminar {path} después de {max_attempts} intentos.")
                return False

def clean_dist():
    """Limpia el directorio dist y archivos temporales con manejo de errores"""
    directories_to_clean = ['dist', 'build', '__pycache__']
    
    # Cerrar cualquier proceso que pueda estar usando los archivos
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'analyze_portfolios.exe'], 
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL)
    except Exception:
        pass

    # Limpiar directorios
    for directory in directories_to_clean:
        if os.path.exists(directory):
            print(f"Limpiando directorio: {directory}")
            if not retry_remove(directory):
                print(f"ADVERTENCIA: No se pudo eliminar {directory}. Continuando...")

    # Limpiar archivos .spec
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            try:
                os.remove(file)
            except PermissionError:
                print(f"No se pudo eliminar {file}. Continuando...")

def create_release_zip():
    """Crea un archivo ZIP del release para GitHub"""
    zip_name = 'AppActualizada.zip'
    if os.path.exists(zip_name):
        retry_remove(zip_name)
    
    try:
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Verificar que el ejecutable existe
            exe_path = os.path.join('dist', 'AppActualizada.exe')
            if not os.path.exists(exe_path):
                # Si no existe, buscar por el otro nombre
                exe_path = os.path.join('dist', 'analyze_portfolios.exe')
                if not os.path.exists(exe_path):
                    raise FileNotFoundError("No se encontró el archivo ejecutable en /dist")
            
            print(f"Agregando ejecutable: {exe_path}")
            # Añadir el ejecutable
            zipf.write(exe_path, 'AppActualizada.exe')
            
            # Añadir version.txt
            print("Agregando version.txt")
            zipf.write('dist/version.txt', 'version.txt')
            
            # Añadir la carpeta data y sus archivos
            print("Agregando archivos de data")
            for root, dirs, files in os.walk('dist/data'):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.join('data', file)
                    print(f"Agregando: {arcname}")
                    zipf.write(file_path, arcname)
        
        print(f"ZIP creado exitosamente: {zip_name}")
        
    except Exception as e:
        print(f"Error al crear el ZIP: {e}")
        import traceback
        traceback.print_exc()
        raise

def build_app():
    print("Iniciando proceso de build...")
    
    # Limpiar directorios anteriores
    clean_dist()
    
    # Configurar los argumentos para PyInstaller
    args = [
        'src/main_app.py',              # Archivo principal
        '--onefile',                    # Crear un solo ejecutable
        '--noconsole',                  # Sin consola
        '--name=AppActualizada',        # Cambiado para coincidir
        '--clean',                      # Limpieza antes de construir
        # Incluir archivos necesarios
        '--add-data=version.txt;.',     # version.txt en raíz
        '--add-data=data;data',         # Carpeta data completa
        '--add-data=src/base_app.py;.', # Incluir base_app.py
        '--add-data=src/temp_handler.py;.',  # Incluir temp_handler.py
    ]
    
    try:
        print("Ejecutando PyInstaller...")
        PyInstaller.__main__.run(args)
        
        print("Copiando archivos adicionales...")
        # Copiar version.txt
        shutil.copy2('version.txt', 'dist/version.txt')
        
        # Copiar la carpeta data
        if os.path.exists('dist/data'):
            retry_remove('dist/data')
        shutil.copytree('data', 'dist/data')
        
        print("Creando ZIP para release...")
        create_release_zip()
        
        print("\nBuild completado exitosamente!")
        print("Archivos generados:")
        print(" - dist/AppActualizada.exe")
        print(" - AppActualizada.zip (listo para subir a GitHub)")
        
    except Exception as e:
        print(f"\nError durante el build: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        build_app()
    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario.")
    except Exception as e:
        print(f"\nError inesperado: {str(e)}")
        raise