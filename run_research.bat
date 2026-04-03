@echo off
:: Change directory to the script's location
cd /d "%~dp0"

:: Execute the python script
echo [%date% %time%] Starting Investment News Research... >> run_log.txt
python -u main.py >> run_log.txt 2>&1
echo [%date% %time%] Finished execution. >> run_log.txt
