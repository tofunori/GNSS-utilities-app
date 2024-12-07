@echo off
echo Nettoyage des anciens builds...
rmdir /s /q dist
del /q *.spec
echo Installation/Mise à jour des dépendances...
pip install -U pyinstaller
echo Construction avec PyInstaller...
pyinstaller --clean ^
    --windowed ^
    --name "PPK Batch Processor" ^
    --add-data "tools.py;." ^
    --add-data "Drone_GNSS_app_v1.3.py;." ^
    --icon="assets/icon.ico" ^
    "PPK batch processor.py"
echo Build terminé!
pause