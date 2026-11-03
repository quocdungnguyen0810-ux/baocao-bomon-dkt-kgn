@echo off
chcp 65001 >nul
title TRIỂN KHAI & CẬP NHẬT WEBAPP LÊN STREAMLIT COMMUNITY CLOUD (0 ĐỒNG)
echo ============================================================================
echo   ĐỒNG BỘ MÃ NGUỒN & CSDL LÊN GITHUB -> STREAMLIT COMMUNITY CLOUD (MIỄN PHÍ)
echo   Đơn vị: Bộ môn Địa kỹ thuật và Không gian ngầm - Trường ĐH GTVT (UTC)
echo ============================================================================
echo.

cd /d "%~dp0"

git add -A
git commit -m "Cap nhat he thong Bao cao Cong viec Bo mon DKT & KGN (TB 940)"
git push -u origin main

echo.
echo [HOÀN TẤT] Đã đẩy bản cập nhật mới nhất lên GitHub!
echo Streamlit Community Cloud (https://share.streamlit.io) sẽ tự động cập nhật Webapp sau 10-15 giây.
pause

