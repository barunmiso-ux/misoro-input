@echo off
REM Weekly: noshow/conversion + per-person(doctor/counselor) + per-disease -> cache sheet.
REM Pure-ASCII only (scheduler bat rule: Korean in REM breaks the cmd OEM parser).
cd /d "%~dp0"
if not exist logs mkdir logs
set PYTHONUTF8=1
set PY="C:\Users\user\AppData\Local\Python\pythoncore-3.14-64\python.exe"
%PY% -u noshow_publish.py --commit >> logs\noshow.log 2>&1
%PY% -u person_publish.py --commit >> logs\person.log 2>&1
%PY% -u disease_publish.py --commit >> logs\disease.log 2>&1
