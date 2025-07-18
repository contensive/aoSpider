
rem build and deliver to deployment folder

set appName=veronica

call build.cmd

rem upload to contensive application
c:
cd %collectionPath%
cc -a %appName% --installFile "%collectionName%.zip"
cd ..\..\scripts

rem -- done --
pause