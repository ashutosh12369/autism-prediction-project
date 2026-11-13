# Launch ASD Clinical Screening Web Application
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Autism Spectrum Disorder (ASD) Prediction System" -ForegroundColor Green
Write-Host "  Final Year B.Tech ML Capstone Project & Live Demo" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Set-Location $PSScriptRoot
python -m streamlit run app/app.py
