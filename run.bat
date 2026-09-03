@echo off
setlocal
title IT Infrastructure SLA Reporter

cd /d "%~dp0"

:: Sanal ortam (.venv) kontrolü - yoksa kurulumu otomatik başlat
if not exist ".venv\Scripts\python.exe" (
    echo ===========================================================================
    echo   Ilk kurulum yapilmamis! Kurulum baslatiliyor...
    echo ===========================================================================
    echo.
    call setup.bat
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [!] HATA: Kurulum tamamlanamadi. Cikis yapiliyor.
        pause
        exit /b 1
    )
    cls
)

:: Sürükle-bırak desteği kontrolü (Dosya run.bat üzerine sürüklendiyse %1 doludur)
if not "%~1"=="" (
    echo [*] Suruklenen dosya isleniyor: "%~1"
    .venv\Scripts\python.exe generate_sla_summary.py "%~1"
) else (
    .venv\Scripts\python.exe generate_sla_summary.py
)

if errorlevel 1 (
    echo.
    echo [!] Islem bir hata ile sonlandi veya kullanici tarafindan iptal edildi.
)

echo.
pause
