@echo off
rem
rem Build aoSpider collection
rem Calls build.ps1 for the actual build process
rem Usage: build.cmd [/nopause]
rem

pushd "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0build.ps1"
set buildError=%errorlevel%
popd

if "%~1"=="/nopause" goto :done
pause
:done
exit /b %buildError%
