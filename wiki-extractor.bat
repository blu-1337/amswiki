@echo off
setlocal

rem Use the known working Edge invocation pattern
rem (quotes are part of EDGE so we call %EDGE% directly)
set EDGE="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
set PROFILE=C:\Users\G1HDMGS\AppData\Local\Microsoft\Edge\User Data
set BASEURL=https://ams-wiki.in.audi.vwg/wiki/bin/view/PPService
set TOPICS="C:\Users\G1HDMGS\Documents\topics.txt"
set OUTPUT=C:\Users\G1HDMGS\Documents\wiki_output

if not exist "%OUTPUT%" mkdir "%OUTPUT%"

echo Starting export...
echo Topics file: %TOPICS%
echo Output dir:  %OUTPUT%
echo.

for /f "usebackq delims=" %%T in (%TOPICS%) do call :ProcessTopic "%%T"

echo.
echo Done.
pause
endlocal
goto :EOF

:ProcessTopic
set "TOPIC=%~1"
if "%TOPIC%"=="" goto :EOF

set "PDFFILE=%OUTPUT%\%TOPIC%.pdf"

rem If file does not exist, go download
if not exist "%PDFFILE%" goto download

rem File exists, check size
for %%A in ("%PDFFILE%") do set "SIZE=%%~zA"

rem Non-zero size -> skip
if "%SIZE%"=="0" goto redownload

echo Skipping %TOPIC% - PDF already exists (%SIZE% bytes).
set "SIZE="
goto :EOF

:redownload
echo Existing PDF for %TOPIC% is zero bytes, re-downloading...
set "SIZE="

:download
echo Exporting %TOPIC% ...
%EDGE% --headless --disable-gpu ^
    --user-data-dir="%PROFILE%" ^
    --print-to-pdf="%PDFFILE%" ^
    "%BASEURL%/%TOPIC%?skin=print"

rem Optional: small delay between requests
timeout /t 2 /nobreak >nul
goto :EOF