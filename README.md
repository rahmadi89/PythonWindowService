# PythonWindowService
Write a single-file Windows service with python

run cmd as administrator and then run these commands:
```
mkdir C:\pyvenv\
```
```
cd c:\pyvenv\
```
```
python -m venv .
```
```
Scripts\activate.bat
```
```
pip install pywin32 pyinstaller 
```
```
pyinstaller --onefile --hidden-import win32timezone PythonServiceFramework.py
```
```
dist\PythonServiceFramework.exe install
```
```
dist\PythonServiceFramework.exe start
```
```
dist\PythonServiceFramework.exe stop
```
