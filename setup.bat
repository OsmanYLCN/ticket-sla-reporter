@echo off
setlocal
title Prometeon SLA Reporter - Kurulum

echo ===========================================================================
echo   PROMETEON IT INFRASTRUCTURE - SLA SUMMARY REPORTER KURULUMU
echo ===========================================================================
echo.

cd /d "%~dp0"

REM 1. Python Kontrolu
echo [1/4] Python kontrol ediliyor...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [!] HATA: Sistemde Python bulunamadi veya PATH ortam degiskenine eklenmemis!
    echo     Lutfen Python 3.9 veya daha yeni bir surumu kurun:
    echo     https://www.python.org/downloads/
    echo     Not: Kurulum sirasinda Add Python to PATH secenegini isaretleyin.
    echo.
    pause
    exit /b 1
)

for /f "delims=" %%v in ('python --version 2^>^&1') do echo       Tespit edilen: %%v

REM 2. Klasor Kontrolu
echo.
echo [2/4] Calisma klasorleri kontrol ediliyor...
if not exist "sample_data" (
    mkdir sample_data
    echo       sample_data klasoru olusturuldu.
)
if not exist "outputs" (
    mkdir outputs
    echo       outputs klasoru olusturuldu.
)

REM 3. Sanal Ortam
echo.
echo [3/4] Izole sanal ortam hazirlaniyor...
if not exist ".venv\Scripts\python.exe" (
    echo       Sanal ortam olusturuluyor, lutfen bekleyin...
    python -m venv .venv
    if errorlevel 1 (
        echo [!] HATA: Sanal ortam olusturulamadi.
        pause
        exit /b 1
    )
    echo       Sanal ortam basariyla olusturuldu.
) else (
    echo       Mevcut sanal ortam .venv bulundu.
)

REM 4. Paketlerin Yuklenmesi
echo.
echo [4/4] Gerekli paketler yukleniyor...
.venv\Scripts\python.exe -m pip install --upgrade pip >nul 2>&1
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [!] HATA: Paketler yuklenirken bir sorun olustu!
    echo     Internet baglantinizi kontrol edip tekrar deneyin.
    echo.
    pause
    exit /b 1
)

echo.
echo ===========================================================================
echo   KURULUM BASARIYLA TAMAMLANDI!
echo ===========================================================================
echo.
echo   Kullanim:
echo   1. Ham Excel dosyalarinizi 'sample_data' klasorune koyun.
echo   2. 'run.bat' dosyasina cift tiklayarak raporlayiciyi calistirin.
echo      Veya bir Excel dosyasini dogrudan 'run.bat' uzerine surukleyip birakin!
echo.
pause
