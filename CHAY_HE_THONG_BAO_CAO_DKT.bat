@echo off
chcp 65001 >nul
title HE THONG TAC TU BAO CAO CONG VIEC & DONG BO USMART - BO MON DIA KY THUAT & KGN (TB 940)
cd /d "%~dp0"

:MENU
cls
echo ========================================================================================
echo    HE THONG TAC TU BAO CAO CONG VIEC VA DONG BO USMART (THONG BAO 940/TB-DHGTVT)
echo    Don vi: Bo mon Dia ky thuat va Khong gian ngam (DonVi=412 - Truong DH GTVT)
echo    Truong Bo mon: GS.TS. Tran The Truyen ^| Pho Truong Bo mon: PGS.TS. Nguyen Chau Lan
echo ========================================================================================
echo.
echo  [1] Mo Giao dien Webapp Tuong tac tren Trinh duyet (http://localhost:8501)
echo  [2] Dong bo TKB Giang day (110 lop) ^& De tai NCKH/NCKHSV tu uSmart (TK lannc)
echo  [3] Xuat dong loat 42 Bao cao Ca nhan (.xlsx ^& .docx) cho 21 GV + Mau 02 + Mau 03
echo  [4] Dong bo ^& Luu nhap Bao cao Thang 11/2026 len uSmart (TK hanh1 - Nguyen Hai Ha)
echo  [5] Chay Bo Kiem thu Kiem toan Tu dong Toan trinh (Unittest Audit Suite)
echo  [0] Thoat
echo.
set /p choice="Nhap lua chon cua ban (0-5): "

if "%choice%"=="1" (
    echo Dang khoi chay Giao dien Webapp Streamlit tai http://localhost:8501 ...
    python -m streamlit run he_thong_bao_cao_dkt/app_webapp.py
    pause
    goto MENU
)
if "%choice%"=="2" (
    echo Dang dong bo Thoi khoa bieu Giang day va De tai NCKH/NCKHSV tu uSmart...
    python -m he_thong_bao_cao_dkt.tkb_giang_day_sync
    python -m he_thong_bao_cao_dkt.usmart_nckh_sync
    pause
    goto MENU
)
if "%choice%"=="3" (
    echo Dang xuat toan bo bao cao ca nhan (21 GV), Mau 02 Don vi va So giao viec Mau 03...
    python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --dong-bo-nckh --xuat-toan-bo-bo-mon --che-do-usmart khong
    pause
    goto MENU
)
if "%choice%"=="4" (
    echo Dang dong bo va luu nhap bao cao len https://usmart.utc.edu.vn/baocaocongviec/ ...
    python -m he_thong_bao_cao_dkt.cli_baocao --thang 11 --nam 2026 --ma-cb 1051 --dong-bo-nckh --xuat-toan-bo-bo-mon --usmart-user hanh1 --usmart-pass 182128 --che-do-usmart luu_nhap
    pause
    goto MENU
)
if "%choice%"=="5" (
    echo Dang chay toan bo bo kiem thu kiem toan tu dong...
    python -m unittest discover -s he_thong_bao_cao_dkt/tests -p "test_*.py" -v
    pause
    goto MENU
)
if "%choice%"=="0" exit /b 0
goto MENU
