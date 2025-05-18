@echo off
REM Chạy tool nhập liệu Excel dạng phần mềm .exe (yêu cầu đã cài Python và các thư viện)
cd /d %~dp0
python dataExcelTool.py
pause
