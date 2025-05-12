@echo off
chcp 65001 >nul
color a
REM filepath: c:\Users\kowkow\Desktop\kBackup\kBackup.bat
REM ==================================================================
REM Script Batch pour Sauvegardes Automatiques
REM ==================================================================

:: Variables globales
set "LOG_FILE=%USERPROFILE%\Desktop\kBackup_log.txt"
set "BACKUP_DIR=C:\_kBackup"
set "SCRIPT_DIR=C:\_kBackup\script"
set "SOURCE_DIR="
set "FREQUENCY="
set "TASK_NAME=kBackupTask"
set "BACKUP_TIME="

if "%1"=="run_backup" (
    goto run_backup
)


:: Fonction principale
:main
cls
echo ============================================================
echo.
echo               Script de Sauvegarde Automatique
echo.
echo ============================================================
echo.
echo [1] Créer une nouvelle tâche de sauvegarde
echo [2] Liste des tâches planifiées
echo [3] Quitter
echo.
echo ============================================================
echo.
set /p "CHOICE=Choisissez une option : "

if "%CHOICE%"=="1" goto create_task
if "%CHOICE%"=="2" goto list_tasks
if "%CHOICE%"=="3" (
    echo Merci d'avoir utilisé kBackup.
    echo Appuyez sur une touche pour quitter...
    pause >nul
    exit /b
)
goto main

:: Créer une nouvelle tâche de sauvegarde
:create_task
cls
echo ============================================================
echo.
echo               Création d'une Tâche de Sauvegarde
echo.
echo ============================================================
echo.

:: Vérifier et créer le répertoire pour le script
if not exist "%SCRIPT_DIR%" (
    mkdir "%SCRIPT_DIR%"
    if errorlevel 1 (
        echo [ERREUR] Impossible de créer le répertoire pour le script. >> "%LOG_FILE%"
        pause
        goto main
    )
)

:: Copier le script dans le répertoire de destination
copy "%~f0" "%SCRIPT_DIR%\kBackup.bat" >nul
if errorlevel 1 (
    echo [ERREUR] Impossible de copier le script dans le répertoire cible. >> "%LOG_FILE%"
    pause
    goto main
)

:: Demander le chemin du répertoire source
set /p "SOURCE_DIR=Entrez le chemin du répertoire source : "
if not exist "%SOURCE_DIR%" (
    echo [ERREUR] Le chemin source est invalide. >> "%LOG_FILE%"
    pause
    goto main
)

:: Vérifier et créer le répertoire de sauvegarde par défaut
if not exist "%BACKUP_DIR%" (
    mkdir "%BACKUP_DIR%"
    if errorlevel 1 (
        echo [ERREUR] Impossible de créer le répertoire de sauvegarde par défaut. >> "%LOG_FILE%"
        pause
        goto main
    )
)

:: Demander la fréquence de sauvegarde
echo.
echo [1] Quotidienne (Daily)
echo [2] Hebdomadaire (Weekly)
echo [3] Mensuelle (Monthly)
echo [4] Toutes les fréquences (Daily/Weekly/Monthly)
echo.
set /p "FREQUENCY=Choisissez la fréquence de sauvegarde (1/2/3/4) : "

if "%FREQUENCY%"=="1" (
    set "SCHEDULE=Daily"
    if not exist "%BACKUP_DIR%\%SCHEDULE%" (
        mkdir "%BACKUP_DIR%\%SCHEDULE%"
    )
) else if "%FREQUENCY%"=="2" (
    set "SCHEDULE=Weekly"
    if not exist "%BACKUP_DIR%\%SCHEDULE%" (
        mkdir "%BACKUP_DIR%\%SCHEDULE%"
    )
) else if "%FREQUENCY%"=="3" (
    set "SCHEDULE=Monthly"
    if not exist "%BACKUP_DIR%\%SCHEDULE%" (
        mkdir "%BACKUP_DIR%\%SCHEDULE%"
    )
) else if "%FREQUENCY%"=="4" (
    set "SCHEDULE=All"
    if not exist "%BACKUP_DIR%\Daily" (
        mkdir "%BACKUP_DIR%\Daily"
    )
    if not exist "%BACKUP_DIR%\Weekly" (
        mkdir "%BACKUP_DIR%\Weekly"
    )
    if not exist "%BACKUP_DIR%\Monthly" (
        mkdir "%BACKUP_DIR%\Monthly"
    )
) else (
    echo [ERREUR] Fréquence invalide. >> "%LOG_FILE%"
    echo [DEBUG] Fréquence invalide : %FREQUENCY%
    pause
    goto main
)

:: Demander l'heure de la sauvegarde
set /p "HOURS_TIME=Choisissez l'heure de sauvegarde (0 à 23) : "


for /L %%X in (0,1,23) do (
    REM Si l'heure est correcte, on passe à la suite grâce au GOTO
    if "%HOURS_TIME%"=="%%X" (

        REM Ajouter un zéro initial si l'heure est inférieure ou égale à 9
        if %%X LEQ 9 (
            set "%%X=0%%X"
        ) else (
            set "%%X=%%X"
        )

        set "BACKUP_TIME=%%X:00"

        goto suite3
    )
)

:suite3

if "%SCHEDULE%"=="All" (
    for %%F in (Daily Weekly Monthly) do (

        if not exist "%BACKUP_DIR%\%%F" (
            mkdir "%BACKUP_DIR%\%%F"
        )

        schtasks /create /tn "%TASK_NAME%_%%F" /tr "cmd /c \"%SCRIPT_DIR%\kBackup.bat run_backup %SOURCE_DIR% %%F\"" /sc %%F /st %BACKUP_TIME% /f

        if errorlevel 1 (
            pause
            goto main
        )

        echo [SUCCES] Tâche planifiée pour %%F créée avec succès.
    )
) else (
    schtasks /create /tn "%TASK_NAME%" /tr "cmd /c \"%SCRIPT_DIR%\kBackup.bat run_backup %SOURCE_DIR% %SCHEDULE%\"" /sc %SCHEDULE% /st %BACKUP_TIME% /f
    if errorlevel 1 (
        pause
    )
    echo [SUCCES] Tâche planifiée créée avec succès.
    goto main
)
pause
goto main


:list_tasks
cls
echo ============================================================
echo.
echo               Liste des Tâches Planifiées
echo.
echo ============================================================
echo.
schtasks /query /fo TABLE | findstr /I "kBackupTask" > "%TEMP%\tasks.txt"

:: Vérifier si des tâches ont été trouvées
if not exist "%TEMP%\tasks.txt" (
    echo [INFO] Aucune tâche planifiée trouvée.
    pause
    goto main
)

:: Afficher les tâches sous forme de tableau
setlocal enabledelayedexpansion
set "INDEX=0"
for /f "tokens=1,* delims= " %%A in ('type "%TEMP%\tasks.txt"') do (

    if not "%%A"=="Nom" (
        set /a INDEX+=1
        set "TASK_!INDEX!=%%A"
        echo [!INDEX!] %%A %%B
    )
)

:: Demander à l'utilisateur de choisir une tâche à supprimer
echo.
echo ============================================================
echo.
set /p "TASK_INDEX=Entrez le numéro de la tâche à supprimer (ou appuyez sur Entrée pour revenir au menu) : "

:: Vérifier si l'utilisateur n'a rien entré
if "%TASK_INDEX%"=="" (
    echo [INFO] Aucune tâche sélectionnée. Retour au menu principal.
    pause
    goto main
)

:: Vérifier si l'index est valide
if not defined TASK_%TASK_INDEX% (
    echo [ERREUR] Index invalide ou aucune tâche correspondante.
    pause
    goto list_tasks
)

:: Supprimer la tâche sélectionnée
set "TASK_NAME=!TASK_%TASK_INDEX%!"
schtasks /delete /tn "!TASK_NAME!" /f
if errorlevel 1 (
    echo [ERREUR] Impossible de supprimer la tâche : !TASK_NAME!.
    pause
    goto list_tasks
)

echo [SUCCES] Tâche supprimée avec succès : !TASK_NAME!.
pause
goto main




:: Exécuter la sauvegarde
:run_backup
echo [INFO] Exécution de la sauvegarde...
echo [DEBUG] Chemin de sauvegarde : %BACKUP_DIR%
echo [DEBUG] Chemin source : %2

:: Vérifier si un argument est passé pour le chemin source
if "%2"=="" (
    echo [ERREUR] Aucun chemin source n'a été fourni. >> "%LOG_FILE%"
    pause
    exit /b
)

if "%3"=="" (
    echo [ERREUR] Aucune fréquence de sauvegarde n'a été fourni. >> "%LOG_FILE%"
    pause
    exit /b
)

set "SOURCE_DIR=%~2"
set "FREQUENCY=%~3"

if not exist "%SOURCE_DIR%" (
    echo [ERREUR] Le répertoire source n'existe pas : %SOURCE_DIR% >> "%LOG_FILE%"
    pause
    exit /b
)

:: Vérifier si le répertoire source est vide
dir "%SOURCE_DIR%" /b >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Le répertoire source est vide : %SOURCE_DIR% >> "%LOG_FILE%"
    echo [ERREUR] Le répertoire source est vide : %SOURCE_DIR%
    pause
    exit /b
)

:: Vérifier si le répertoire de sauvegarde existe, sinon le créer
if not exist "%BACKUP_DIR%" (
    echo [INFO] Le répertoire de sauvegarde n'existe pas. Création en cours...
    mkdir "%BACKUP_DIR%"
    if errorlevel 1 (
        echo [ERREUR] Impossible de créer le répertoire de sauvegarde par défaut : %BACKUP_DIR% >> "%LOG_FILE%"
        pause
        exit /b
    )
)

:: Obtenir un horodatage simple pour le dossier de sauvegarde
set "DATE=%date:/=-%"
set "TIMESTAMP=%DATE% KBACKUP"

set "DEST_DIR=%BACKUP_DIR%\%FREQUENCY%\%TIMESTAMP%"
mkdir "%DEST_DIR%"
if errorlevel 1 (
    echo [ERREUR] Impossible de créer le répertoire de destination : %DEST_DIR% >> "%LOG_FILE%"
    pause
    exit /b
)

xcopy "%SOURCE_DIR%" "%DEST_DIR%" /E /I /H /C /Y >nul
if errorlevel 1 (
    echo [ERREUR] La sauvegarde a échoué. >> "%LOG_FILE%"
    pause
    exit /b
)
echo [SUCCES] Sauvegarde effectuée avec succès. >> "%LOG_FILE%"
pause
exit /b