@echo off
color 0E
echo ========================================================
echo     PURGE DES 267 PREMIERES LETTRES OBSOLETES
echo ========================================================
echo.
echo Ce script va supprimer les 267 premiers cabinets du cache
echo pour forcer l'IA a les refaire, SANS toucher aux 950 autres.
echo.
echo ATTENTION : Fermez la fenetre noire de l'IA (lancer.bat) 
echo avant de continuer !
echo.
pause
python purger_anciennes_lettres.py
