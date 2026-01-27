@ECHO OFF

@REM :: ECHO Just update all in dist directory; bring all to dist


@REM :: 1. delete all first
IF EXIST "%~dp0dist\" (
    RMDIR /s /q "%~dp0dist"
)
MKDIR "%~dp0dist"


@REM :: 2. bring files to "dist"
COPY "%~dp0src\correlations.py" "%~dp0dist\"
COPY "%~dp0run_*.bat" "%~dp0dist\"
COPY "%~dp0requirements.txt" "%~dp0dist\"


@REM :: 3. And update files in "tests" subfolder
IF EXIST "%~dp0tests\" (
    COPY "%~dp0dist\*" "%~dp0tests\"
)

