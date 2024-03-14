@echo off
chcp 65001 >nul
SETLOCAL EnableExtensions EnableDelayedExpansion
title Conversion Toolkit
cls
echo.
echo ---------------------------------------------------------
echo Ce logiciel est sous licence MIT. Pour plus d'informations, consultez le fichier LICENSE.
echo Retrouvez-moi sur:
echo Gitea: https://gitea.askanat.com/
echo LinkedIn: www.linkedin.com/in/florian-vaissiere-2bab64122
echo ---------------------------------------------------------
echo.

echo  ██████╗ ██████╗ ███╗   ██╗██╗   ██╗███████╗██████╗ ███████╗██╗ ██████╗ ███╗   ██╗   
echo ██╔════╝██╔═══██╗████╗  ██║██║   ██║██╔════╝██╔══██╗██╔════╝██║██╔═══██╗████╗  ██║  
echo ██║     ██║   ██║██╔██╗ ██║██║   ██║█████╗  ██████╔╝███████╗██║██║   ██║██╔██╗ ██║   
echo ██║     ██║   ██║██║╚██╗██║╚██╗ ██╔╝██╔══╝  ██╔══██╗╚════██║██║██║   ██║██║╚██╗██║    
echo ╚██████╗╚██████╔╝██║ ╚████║ ╚████╔╝ ███████╗██║  ██║███████║██║╚██████╔╝██║ ╚████║     
echo  ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝  ╚═══╝  ╚══════╝╚═╝  ╚═╝╚══════╝╚═╝ ╚═════╝ ╚═╝  ╚═══╝  

echo     ████████╗ ██████╗  ██████╗ ██╗     ██╗  ██╗██╗████████╗
echo     ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██║ ██╔╝██║╚══██╔══╝
echo        ██║   ██║   ██║██║   ██║██║     █████╔╝ ██║   ██║   
echo        ██║   ██║   ██║██║   ██║██║     ██╔═██╗ ██║   ██║   
echo        ██║   ╚██████╔╝╚██████╔╝███████╗██║  ██╗██║   ██║   
echo        ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝   ╚═╝ 

echo.
echo ---------------------------------------------------------
echo Vérification des prérequis!
echo ---------------------------------------------------------
echo.

:: Création d'un fichier de logs pour les erreurs
set "logFile=%~dp0conversionLogs.log"

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
python -m pip install pywin32 comtypes pdf2docx >nul 2>>"%logFile%"
echo.
echo Les bibliothèques nécessaires ont été installées.

echo.

echo ---------------------------------------------------------
echo Bienvenue dans l'assistant de conversion de documents!
echo ---------------------------------------------------------

echo.

:main_loop
echo Veuillez glisser et déposer un ou plusieurs fichiers séparé par un espace, ou un dossier, puis appuyez sur Entrée.
set /p input="Entrée: "
echo.
set /p output_folder="Chemin du dossier de sortie (facultatif): "
if "!output_folder!"=="" (
    set "output_folder=%cd%\ResultatConversion"
)

:: Crée le dossier s'il n'existe pas déjà
if not exist "!output_folder!" (
    mkdir "!output_folder!"
)

set "input=!input:"=!"
if "!input!"=="" goto ask_continue

:: Détermine si l'entrée est un dossier
if exist "!input!\*" (
    echo Dossier détecté.
    set "source_folder=!input!"
    goto process_folder
) else if exist "!input!" (
    echo Fichier détecté.
    set "output_folder=%cd%\ResultatConversion"
    if not exist "!output_folder!" mkdir "!output_folder!"
    call :process_file "!input!"
    goto ask_continue
) else (
    echo Ni un dossier valide ni un fichier détecté.
    goto ask_continue
)

:: Traitement de plusieurs fichiers ou d'un seul fichier
set "output_folder=%cd%"
echo.
:process_files
for %%i in (!input!) do (
    call :process_file "%%~fi"
)
goto ask_continue

:process_folder
echo Traitement du dossier: !source_folder!
for /r "%source_folder%" %%f in (*.*) do (
    call :process_file "%%f"
)
goto ask_continue

:process_file
set "file_path=%~1"
echo ----------------------------------------------------------------
echo Traitement de: %file_path%

:: Appel du script Python
echo %date% %time% - Début de la conversion de: %file_path% >> "%logFile%"
python desk_tool_converter.py "%file_path%" "%output_folder%" >nul 2>>"%logFile%"

if !errorlevel! neq 0 (
    echo %date% %time% - Une erreur est survenue pendant la conversion de: %file_path% >> "%logFile%"
    echo Une erreur est survenue pendant la conversion de: %file_path%. Consultez le fichier de logs pour plus de détails.
) else (
    echo %date% %time% - La conversion de: %file_path% est terminée. >> "%logFile%"
    echo La conversion de: %file_path% est terminée. Vérifiez vos documents dans le dossier de sortie : "%output_folder%".
)
goto :eof

:ask_continue
echo.
echo Voulez-vous convertir d'autres fichiers ou dossiers ? (O/N)
set /p continue="Réponse: "
if /i "!continue!"=="O" goto main_loop

:end
echo ----------------------------------------------------------------
echo Merci d'avoir utilisé cet assistant de conversion.
echo ----------------------------------------------------------------
pause