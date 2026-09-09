@echo off
chcp 65001 >nul
title Esplora File - crea il file .exe (Nuitka)
echo ============================================
echo   Generazione dell'eseguibile con Nuitka
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERRORE: Python non trovato nel PATH.
    echo Installa Python da https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Installazione/aggiornamento di Nuitka...
python -m pip install --upgrade nuitka
if errorlevel 1 (
    echo ERRORE durante l'installazione di Nuitka.
    pause
    exit /b 1
)

echo.
echo [2/5] Installazione di mutagen (metadati audio/video)...
python -m pip install mutagen
if errorlevel 1 (
    echo AVVISO: mutagen non installato: l'anteprima audio/video restera' disabilitata.
)

echo.
echo [3/5] Installazione di py7zr (7z senza 7-Zip)...
set "PY7ZR_FLAGS="
python -m pip install py7zr
if errorlevel 1 (
    echo AVVISO: py7zr non installato: il 7z richiedera' 7-Zip installato.
) else (
    rem NOTA: il pacchetto PyPI si chiama pybcj ma il modulo importabile e' 'bcj'
    rem (importato da py7zr.compressor), quindi l'include usa 'bcj'.
    set "PY7ZR_FLAGS=--include-package=py7zr --include-package=bcj --include-package=pyppmd --include-package=Cryptodome --include-package=inflate64 --include-package=multivolumefile --include-package=texttable --include-package=brotli --include-package=psutil --include-package=mutagen"
)

echo.
echo [4/5] Compilazione con Nuitka (puo' richiedere diversi minuti)...
echo      La prima volta Nuitka scarica anche un compilatore C
echo      (MSVC o MinGW64) se non e' gia' installato nel sistema.
echo.
echo ATTENZIONE: se l'antivirus (Bitdefender, Windows Defender, ecc.) blocca
echo             la fase finale "Failed to add resources", aggiungi questa
echo             cartella alle ESCLUSIONI dell'antivirus e rilancia il build.
python -m nuitka --onefile --windows-console-mode=disable --enable-plugin=tk-inter --output-dir=dist --output-filename=EsploraFile.exe --remove-output --assume-yes-for-downloads %PY7ZR_FLAGS% file_explorer.py
if errorlevel 1 (
    echo ERRORE durante la compilazione.
    pause
    exit /b 1
)

echo.
echo [5/5] Operazione completata!
echo.
echo L'eseguibile si trova in:  dist\EsploraFile.exe
echo (compilato in codice C: piu' veloce e protetto dalla decompilazione)
echo.
pause
