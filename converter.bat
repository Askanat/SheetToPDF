@echo off
chcp 65001 >nul
SETLOCAL EnableExtensions
cls

echo.
echo ---------------------------------------------------------
echo Ce logiciel est sous licence MIT. Pour plus d'informations, consultez le fichier LICENSE.
echo Retrouvez-moi sur:
echo Gitea: https://gitea.askanat.com/
echo LinkedIn: www.linkedin.com/in/florian-vaissiere-2bab64122
echo ---------------------------------------------------------
echo.

echo  ███████╗██╗  ██╗███████╗███████╗████████╗    ████████╗ ██████╗  ██████╗ ██╗     
echo ██╔════╝██║  ██║██╔════╝██╔════╝╚══██╔══╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     
echo ███████╗███████║█████╗  █████╗     ██║          ██║   ██║   ██║██║   ██║██║     
echo ╚════██║██╔══██║██╔══╝  ██╔══╝     ██║          ██║   ██║   ██║██║   ██║██║     
echo ███████║██║  ██║███████╗███████╗   ██║          ██║   ╚██████╔╝╚██████╔╝███████╗
echo ╚══════╝╚═╝  ╚═╝╚══════╝╚══════╝   ╚═╝          ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝
echo.

echo ---------------------------------------------------------
echo Bienvenue dans l'assistant de conversion Excel vers PDF!
echo ---------------------------------------------------------
echo.

:: Création d'un fichier de logs pour les erreurs
set "logFile=%~dp0conversion_errors.log"

:: Vérification de la présence de Python 3
echo Vérification de Python 3 en cours...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python 3 n'est pas détecté. Installation en cours...
    echo Cette opération peut prendre quelques minutes. Veuillez patienter...
    powershell -Command "(new-object net.webclient).DownloadFile('https://www.python.org/ftp/python/3.10.4/python-3.10.4-amd64.exe', 'python-installer.exe')"
    python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 >nul 2>>"%logFile%"
    if %errorlevel% neq 0 (
        echo Une erreur est survenue lors de l'installation de Python 3. Veuillez consulter le fichier de logs pour plus de détails.
        goto End
    )
    del python-installer.exe
    echo Python 3 a été installé avec succès.
) else (
    echo Python 3 est déjà installé sur cet ordinateur.
)

echo.

:: Vérification et installation de pip
echo Vérification de pip...
python -m ensurepip >nul 2>>"%logFile%"
echo.
echo Mise à jour de pip...
python -m pip install --upgrade pip >nul 2>>"%logFile%"
echo.

:: Installation des bibliothèques nécessaires
echo Installation des bibliothèques nécessaires...
python -m pip install pypiwin32 >nul 2>>"%logFile%"
echo.
echo Les bibliothèques nécessaires ont été installées.
echo.

:: Interaction avec l'utilisateur pour les chemins des fichiers
echo ----------------------------------------------------------------
echo Veuillez glisser et déposer le fichier Excel à convertir ici puis appuyez sur Entrée.
set /p source_path="Chemin du fichier source: "
echo.

echo ----------------------------------------------------------------
echo Veuillez glisser et déposer le dossier où vous souhaitez sauvegarder les PDFs puis appuyez sur Entrée.
echo Si vous souhaitez utiliser le répertoire courant, laissez simplement vide et appuyez sur Entrée.
set /p output_folder="Chemin du dossier de sortie (facultatif): "
echo.

:: Vérification du chemin de sortie
if "%output_folder%"=="" set output_folder=%cd%

:: Exécution du script Python avec les chemins
echo ----------------------------------------------------------------
echo Conversion en cours, veuillez patienter...
python SheetToPDF.py "%source_path%" "%output_folder%" >nul 2>>"%logFile%"
if %errorlevel% neq 0 (
    echo Une erreur est survenue pendant la conversion. Veuillez consulter le fichier de logs pour plus de détails.
    goto End
)

echo.
echo La conversion est terminée. Vérifiez vos PDFs dans le dossier spécifié.
echo Merci d'avoir utilisé cet assistant de conversion.

:End
pause
