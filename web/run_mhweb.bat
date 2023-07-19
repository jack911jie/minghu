@echo off
start cmd /c "python app.py > C:\Users\admin\Desktop\app_log.txt 2>&1"
start "" "http://127.0.0.1:5000"