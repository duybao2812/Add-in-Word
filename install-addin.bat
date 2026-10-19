@echo off
chcp 65001 >nul
title Cài Đặt Tự Động Word Add-in Chuẩn Hóa Thể Thức & Tra Cứu Pháp Luật
color 1F

echo =======================================================================
echo     CÀI ĐẶT 1-CLICK: MICROSOFT WORD ADD-IN PHÁP LUẬT & NĐ 30
echo =======================================================================
echo.
echo  Đang tiến hành đăng ký Word Add-in vào hệ thống Windows Office...
echo.

set "MANIFEST_PATH=%~dp0manifest.xml"

if not exist "%MANIFEST_PATH%" (
    echo [LỖI] Không tìm thấy file manifest.xml trong thư mục:
    echo "%~dp0"
    echo Vui lòng đảm bảo file install-addin.bat nằm cùng thư mục với manifest.xml.
    pause
    exit /b 1
)

:: Ghi vào Registry Office 16.0 (Word 2016, 2019, 2021, Office 365)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "GXDWordAddin" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

if %ERRORLEVEL% EQU 0 (
    echo [THÀNH CÔNG] Đã đăng ký thành công Add-in vào Microsoft Word!
    echo.
    echo  Hướng dẫn sử dụng:
    echo  1. Mở ứng dụng Microsoft Word lên.
    echo  2. Chọn tab "Home" hoặc "Insert" -^> tìm nút "Chuẩn hóa NĐ 30".
    echo  3. Bấm vào nút số 6: "Hệ thống tra cứu văn bản pháp luật" để sử dụng.
    echo.
) else (
    echo [CẢNH BÁO] Không thể ghi Registry tự động.
    echo Vui lòng chạy file này bằng quyền "Run as Administrator" (Chuột phải -^> Run as Administrator).
    echo.
)

echo =======================================================================
pause
