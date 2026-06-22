@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title Recherche cabinets Geometres-Experts
cd /d "%~dp0"

REM --- Charge les cles si le fichier cles.bat existe ---
if exist "cles.bat" (
    call "cles.bat"
    echo [OK] Cles chargees depuis cles.bat
) else (
    echo [i] Pas de fichier cles.bat : options --recrutement et --lettres desactivees.
)
echo.

REM --- Choix de la zone ---
echo Quelle zone veux-tu traiter ?
echo    1 = France entiere (long : ~2150 cabinets)
echo    2 = Une region (ex: nouvelle-aquitaine, occitanie...)
echo    3 = Un ou plusieurs departements (ex: 87 19 23)
echo.
set "CHOIX="
set /p CHOIX="Ton choix [1/2/3] : "

set "ZONE="
if "%CHOIX%"=="1" set "ZONE=--tout"
if "%CHOIX%"=="2" (
    set "REG="
    set /p REG="Nom de la region : "
    set "ZONE=--region !REG!"
)
if "%CHOIX%"=="3" (
    set "DEPS="
    set /p DEPS="Departement(s) (ex: 87 19 23) : "
    set "ZONE=--departements !DEPS!"
)
if "!ZONE!"=="" (
    echo Choix invalide. Abandon.
    pause
    exit /b
)

REM --- Construit les options selon les cles disponibles ---
REM (dossier.py scrape toujours les sites : pas besoin de --specialites)
set "OPTIONS="
if defined FT_CLIENT_ID if not "%FT_CLIENT_ID%"=="PAR_xxxxx" set "OPTIONS=%OPTIONS% --recrutement"

echo.
echo Analyser intelligemment le contenu des sites web (plus lent) ?
echo    1 = Non (extraction simple)
echo    2 = Oui, avec un LLM LOCAL (LM Studio doit etre lance)
set "ANALYSE="
set /p ANALYSE="Ton choix [1/2] : "
if "%ANALYSE%"=="2" set "OPTIONS=%OPTIONS% --analyse-site --local"

echo.
echo Generer les lettres de motivation ?
echo    1 = Non (dossiers seulement)
echo    2 = Oui, avec un LLM LOCAL (LM Studio doit etre lance)
echo    3 = Oui, avec l'API Anthropic (cle requise)
set "LET="
set /p LET="Ton choix [1/2/3] : "
if "%LET%"=="2" set "OPTIONS=%OPTIONS% --lettres --local"
if "%LET%"=="3" set "OPTIONS=%OPTIONS% --lettres"

echo.
echo Lancement : python dossier.py %ZONE% %OPTIONS%
echo.

python dossier.py %ZONE% %OPTIONS%
if errorlevel 1 (
    echo.
    echo [!] "python" a echoue, nouvel essai avec "py"...
    py dossier.py %ZONE% %OPTIONS%
)

echo.
echo ====================================================================
echo  Termine. Resultats dans le dossier : sortie\
echo   - sortie\dossiers\           (un dossier .md par cabinet)
echo   - sortie\dossiers_cabinets.xlsx  (recapitulatif)
echo   - sortie\lettres\            (si lettres generees)
echo ====================================================================
echo.
pause
