import sys
from cx_Freeze import setup, Executable

# Dépendances
build_exe_options = {
    "packages": ["os", "numpy", "pandas", "tkinter", "PIL"],
    "excludes": [],
    "include_files": [
        "tools.py",
        "PPK batch processor.py",
        "Drone_GNSS_app_v1.3.py",
        # Ajoutez tous les autres fichiers nécessaires à votre application
    ]
}

# Configuration de base
setup(
    name="PPK Batch Processor",
    version="1.1",
    description="Traitement par lots des données PPK",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "PPK batch processor.py",
            base="Win32GUI" if sys.platform == "win32" else None,
            icon="assets/icon.ico",  # Remplacez avec le vrai chemin
            shortcut_name="PPK Batch Processor",
            shortcut_dir="DesktopFolder"
        )
    ]
) 