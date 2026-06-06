@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Anonymiseur de documents

rem ==========================================================================
rem  Lanceur interactif pour anonymiseur.py
rem  - Mode SCAN   : genere un fichier de correspondances (un nouveau a chaque fois)
rem  - Mode ANONYM : recree l'arborescence du dossier source dans un dossier cible
rem ==========================================================================

set "SCRIPT=%~dp0anonymiseur.py"

if not exist "%SCRIPT%" (
    echo [ERREUR] Script introuvable : "%SCRIPT%"
    echo Place ce .bat dans le meme dossier que anonymiseur.py.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo [ERREUR] Python est introuvable dans le PATH.
    echo Installe Python 3.8+ ou ouvre une session ou la commande "python" fonctionne.
    pause
    exit /b 1
)

:menu
cls
echo ============================================================
echo                ANONYMISEUR DE DOCUMENTS
echo ============================================================
echo.
echo   Formats traites : .txt .md .docx .pptx .xlsx .pdf
echo.
echo   1. Scanner un dossier  (generer un fichier de correspondances)
echo   2. Anonymiser un dossier
echo   3. Quitter
echo.
set "CHOIX="
set /p "CHOIX=Votre choix [1-3] : "
if "%CHOIX%"=="1" goto scan
if "%CHOIX%"=="2" goto anon
if "%CHOIX%"=="3" exit /b 0
goto menu

rem --------------------------------------------------------------------------
:scan
cls
echo ------------------------------------------------------------
echo  MODE SCAN
echo ------------------------------------------------------------
echo  Genere un "correspondances_<dossier>.tsv" dans le sous-dossier
echo  "Scan". Un NOUVEAU fichier a chaque scan (jamais d'ecrasement).
echo  Les termes deja presents dans correspondances_reference.tsv sont
echo  EXCLUS. Aucun document n'est modifie.
echo.
set "SRC="
set /p "SRC=Dossier a scanner (glisser-deposer possible) : "
if defined SRC set SRC=%SRC:"=%
if not defined SRC goto menu
if not exist "!SRC!\" (
    echo.
    echo [ERREUR] Dossier introuvable : "!SRC!"
    echo.
    pause
    goto menu
)
set "REF=%~dp0correspondances_reference.tsv"
echo.
if exist "%REF%" (
    echo Commande : python "%SCRIPT%" "!SRC!" --scan --exclure "correspondances_reference.tsv"
    echo.
    python "%SCRIPT%" "!SRC!" --scan --exclure "%REF%"
) else (
    echo [info] correspondances_reference.tsv introuvable : aucun terme exclu.
    echo Commande : python "%SCRIPT%" "!SRC!" --scan
    echo.
    python "%SCRIPT%" "!SRC!" --scan
)
echo.
echo Termine. Le .tsv genere est dans le sous-dossier "Scan" et ne
echo contient QUE les termes absents de correspondances_reference.tsv.
echo Complete la 2e colonne, puis recopie les lignes utiles dans ton
echo fichier de reference avant de lancer le mode Anonymiser.
echo.
pause
goto menu

rem --------------------------------------------------------------------------
:anon
cls
echo ------------------------------------------------------------
echo  MODE ANONYMISATION
echo ------------------------------------------------------------
echo.
set "SRC="
set /p "SRC=Dossier SOURCE : "
if defined SRC set SRC=%SRC:"=%
if not defined SRC goto menu
if not exist "!SRC!\" (
    echo [ERREUR] Dossier source introuvable : "!SRC!"
    pause
    goto menu
)

set "OUT="
set /p "OUT=Dossier CIBLE (sortie, l'arborescence y sera recreee) : "
if defined OUT set OUT=%OUT:"=%
if not defined OUT goto menu

rem --- Fichier de correspondances : 1er .tsv du dossier du .bat (auto) -----
set "MAP="
set "NBTSV=0"
set "LISTE="
for %%F in ("%~dp0*.tsv") do (
    set /a NBTSV+=1
    if not defined MAP set "MAP=%%~fF"
    set "LISTE=!LISTE! %%~nxF"
)
if not defined MAP (
    echo [ERREUR] Aucun fichier .tsv trouve dans le dossier du lanceur :
    echo          "%~dp0"
    pause
    goto menu
)
if !NBTSV! GTR 1 (
    echo.
    echo [ATTENTION] !NBTSV! fichiers .tsv dans le dossier :!LISTE!
    echo             Le PREMIER est utilise -- verifie que c'est le bon.
)
set "NB=?"
for /f "delims=" %%C in ('python "%SCRIPT%" --compter --map "!MAP!" 2^>nul') do set "NB=%%C"
echo.
echo Fichier de correspondances detecte automatiquement :
echo    "!MAP!"
echo    -^> !NB! correspondance(s) chargee(s).

set "OPTS="

echo.
echo ------------------------------------------------------------
echo  Par DEFAUT, le traitement va :
echo    - remplacer les termes du .tsv (regex + termes exacts) ;
echo    - masquer emails, URLs, IP, MAC, ports et chemins ;
echo    - accepter le suivi de modifications Word, supprimer les commentaires ;
echo    - SUPPRIMER entierement les proprietes du document (auteur, societe,
echo      titres, proprietes custom) et neutraliser les horodatages internes ;
echo    - SUPPRIMER les images et les objets OLE embarques (Excel/Word colles) ;
echo    - anonymiser AUSSI les noms de fichiers et de dossiers ;
echo    - insensible a la casse et aux accents.
echo ------------------------------------------------------------
echo.
echo   [1] Lancer avec les options PAR DEFAUT
echo   [2] Configurer les options en detail (dry-run, casse, images, noms...)
echo.
set "MODE="
set /p "MODE=Votre choix [1-2] : "
if not "!MODE!"=="2" goto anon_run

set "ANS="
set /p "ANS=Simulation, sans rien ecrire (dry-run) ? [o/N] : "
if /i "!ANS!"=="o" set "OPTS=!OPTS! --dry-run"

set "ANS="
set /p "ANS=Respecter la casse (--case-sensitive) ? [o/N] : "
if /i "!ANS!"=="o" set "OPTS=!OPTS! --case-sensitive"

echo.
echo Par defaut, les images sont SUPPRIMEES (si gardees, leurs metadonnees
echo sont strippees mais leur contenu visuel n'est PAS anonymise).
set "ANS="
set /p "ANS=Conserver les images (--garder-images) ? [o/N] : "
if /i "!ANS!"=="o" set "OPTS=!OPTS! --garder-images"

set "ANS="
set /p "ANS=Garder le suivi de modifications Word (--garder-suivi) ? [o/N] : "
if /i "!ANS!"=="o" set "OPTS=!OPTS! --garder-suivi"

echo.
echo Par defaut, les noms de fichiers/dossiers sont AUSSI anonymises.
set "ANS="
set /p "ANS=Garder les noms d'origine (ne PAS anonymiser les noms) ? [o/N] : "
if /i "!ANS!"=="o" set "OPTS=!OPTS! --garder-noms"

echo.
echo Nettoyages techniques actifs par defaut : email, url, ip, mac, port, chemin.
set "SCRUB="
set /p "SCRUB=Lettres a DESACTIVER (E=email U=url I=ip P=port C=chemin M=mac) ou Entree : "
if defined SCRUB (
    echo !SCRUB! | findstr /i "E" >nul && set "OPTS=!OPTS! --no-email"
    echo !SCRUB! | findstr /i "U" >nul && set "OPTS=!OPTS! --no-url"
    echo !SCRUB! | findstr /i "I" >nul && set "OPTS=!OPTS! --no-ip"
    echo !SCRUB! | findstr /i "P" >nul && set "OPTS=!OPTS! --no-port"
    echo !SCRUB! | findstr /i "C" >nul && set "OPTS=!OPTS! --no-path"
    echo !SCRUB! | findstr /i "M" >nul && set "OPTS=!OPTS! --no-mac"
)

:anon_run
cls
echo ------------------------------------------------------------
echo  RECAPITULATIF
echo ------------------------------------------------------------
echo  Source : "!SRC!"
echo  Cible  : "!OUT!"
echo  Map    : "!MAP!"
echo  Options: !OPTS!
echo.
echo  Commande :
echo    python "%SCRIPT%" "!SRC!" --out "!OUT!" --map "!MAP!"!OPTS!
echo.
set "GO="
set /p "GO=Lancer maintenant ? [O/n] : "
if /i "!GO!"=="n" goto menu

echo.
python "%SCRIPT%" "!SRC!" --out "!OUT!" --map "!MAP!"!OPTS!
echo.
echo ------------------------------------------------------------
echo  Termine. Verifie le message ci-dessus et le rapport CSV
echo  dans le dossier cible (_rapport_anonymisation.csv).
echo ------------------------------------------------------------
echo.
pause
goto menu
