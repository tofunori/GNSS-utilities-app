@echo off
echo Nettoyage des anciens fichiers...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "Output" rmdir /s /q Output

echo Installation des dépendances...
pip install pyinstaller

echo Vérification de l'icône...
if not exist "assets\icon.ico" (
    echo ERREUR: L'icône n'existe pas dans assets\icon.ico
    pause
    exit /b 1
)

echo Construction de l'application...
pyinstaller --clean ^
    --windowed ^
    --name "PPK_Batch_Processor" ^
    --add-data "tools.py;." ^
    --add-data "Drone_GNSS_app_v1.3.py;." ^
    --add-data "assets/*;assets/" ^
    --icon="assets\icon.ico" ^
    "PPK batch processor.py"

echo Vérification des fichiers générés...
if exist "dist" (
    echo Le dossier dist existe
    dir dist
) else (
    echo ERREUR: Le dossier dist n'existe pas
    pause
    exit /b 1
)

echo Création de l'installateur...
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
) else (
    echo Erreur: Inno Setup n'est pas installé
    pause
    exit /b 1
)

echo Construction terminée !
if exist "Output\PPK_Batch_Processor_Setup.exe" (
    echo L'installateur a été créé avec succès
    start Output
) else (
    echo Erreur: L'installateur n'a pas été créé
)
pause