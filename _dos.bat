@echo off
title dos(py3.11.x)
color 0A
set current_dir=%cd%

echo DEVELOPER: Vu Van Viet - OAEM Lab
echo.
echo %date% %time%
echo ERROR!!! Please install python 3.11.x
echo ERROR!!! Please install gamspy
echo Cai dat gamspy bang cach go: py -m pip install gamspy
echo ERROR!!! Please install pandas
echo Cai dat pandas bang cach go: py -m pip install pandas
echo.
:loop   
set /p command="%current_dir%>"

if /i "%command%"=="exit" goto end
if /i "%command%"=="cls" (
    cls
    echo DEVELOPED BY: Vu Van Viet - OAEM Lab
    echo.
    echo %date% %time%
    echo ERROR!!! Please install python 3.11.x
	echo ERROR!!! Please install gamspy
	echo Cai dat gamspy bang cach go: py -m pip install gamspy
	echo ERROR!!! Please install pandas
	echo Cai dat pandas bang cach go: py -m pip install pandas
    echo.
    goto loop
)
if /i "%command%"=="dir" (
    echo Version 1.0
    echo.
    echo Directory of %current_dir%
    echo.
    echo 05/29/2025  03:06 PM    ^<DIR^>          .
    echo 05/29/2025  03:06 PM    ^<DIR^>          ..
    echo 05/29/2025  03:05 PM             1,234 example.txt
    echo 05/29/2025  03:04 PM             5,678 program.bat
    echo                2 File(s^)          6,912 bytes
    echo                2 Dir(s^)  15,234,567,890 bytes free
    goto loop
)
if /i "%command%"=="run" (
	echo Bat dau chay file python misocp.py ...
    py -3.11 misocp.py
    goto loop
)
if /i "%command%"=="check" (
    echo ***Checking Python version...
    py -3.11 --version >nul 2>&1
    if %errorlevel%==0 (
        echo    - Python 3.11.x is installed.
    ) else (
        echo    - Python 3.11.x is not installed.
        echo !!!Please install Python 3.11.x.
    )
    echo.
    goto loop
)
if /i "%command%"=="help" (
    echo Available commands:
    echo DIR     - List directory contents
    echo CLS     - Clear screen
    echo VER     - Show version tools
    echo EXIT    - Exit the program
    echo HELP    - Show this help
    echo RUN     - Run python
    echo CHECK   - Check version Python
    goto loop
)

if "%command%"=="" goto loop

REM Thử thực thi lệnh
%command% 2>nul || (
    echo '%command%' is not recognized as an internal or external command,
    echo operable program or batch file.
)

goto loop

:end
exit