@echo off
REM TSMMIS Word Report Generator - Batch Script
REM This script runs the Python document generator

setlocal enabledelayedexpansion

echo Creating TSMMIS Project Report...
echo.

cd /d "D:\MVRP_Project\TSMMIS_MATLAB"

REM Try to run with Python
if exist "generate_docx_no_deps.py" (
    python generate_docx_no_deps.py
    if !errorlevel! equ 0 (
        echo.
        echo Report generation completed successfully!
        if exist "TSMMIS_Project_Report.docx" (
            for %%F in ("TSMMIS_Project_Report.docx") do (
                echo File size: %%~zF bytes
            )
        )
    ) else (
        echo Error during report generation
        exit /b 1
    )
) else (
    echo Error: generate_docx_no_deps.py not found
    exit /b 1
)

pause
