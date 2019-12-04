PyInstaller -F GradeBook.py
move /Y dist\GradeBook.exe .\

del /s /q dist 
rd /s /q  dist

del /s /q __pycache__
rd /s /q __pycache__

del /s /q build
rd /s /q build
del GradeBook.spec

PAUSE