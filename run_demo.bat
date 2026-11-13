@echo off
title ASD Clinical Screening AI - Live Demo
echo ========================================================
echo   Autism Spectrum Disorder (ASD) Prediction System
echo   Final Year B.Tech ML Capstone Project & Live Demo
echo ========================================================
echo.
cd /d "%~dp0"
echo Launching Streamlit web application on local browser...
python -m streamlit run app/app.py --server.headless=false
pause
