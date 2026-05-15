@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Write-Host 'Starting DB Merging Application...' -ForegroundColor Cyan; try { .\python\python.exe -m streamlit run app.py } catch { Write-Host 'An error occurred:' -ForegroundColor Red; $_.Exception.Message }; Write-Host 'Press any key to exit...' -ForegroundColor Yellow; $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
