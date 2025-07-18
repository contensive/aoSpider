
rem @echo off

rem Must be run from the projects git\project\scripts folder - everything is relative
rem run >build [versionNumber]
rem versionNumber is YY.MM.DD.build-number, like 20.5.8.1
rem

c:
cd \Git\aoSpider\scripts

set collectionName=Spider
set solutionName=Spider
set collectionPath=..\collections\aoSpider\
set binPath=..\server\Spider\bin\Debug\
set deploymentFolderRoot=C:\Deployments\aoSpider\Dev\
set msbuildLocation=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\
set NuGetLocalPackagesFolder=C:\NuGetLocalPackages\

rem @echo off
rem Setup deployment folder

set year=%date:~12,4%
set month=%date:~4,2%
if %month% GEQ 10 goto monthOk
set month=%date:~5,1%
:monthOk
set day=%date:~7,2%
if %day% GEQ 10 goto dayOk
set day=%date:~8,1%
:dayOk
set versionMajor=%year%
set versionMinor=%month%
set versionBuild=%day%
set versionRevision=1
rem
rem if deployment folder exists, delete it and make directory
rem
:tryagain
set versionNumber=%versionMajor%.%versionMinor%.%versionBuild%.%versionRevision%
if not exist "%deploymentFolderRoot%%versionNumber%" goto :makefolder
set /a versionRevision=%versionRevision%+1
goto tryagain
:makefolder
md "%deploymentFolderRoot%%versionNumber%"

pause


rem ==============================================================
rem

del "%collectionPath%Spider.zip"

pause


rem 
rem ==============================================================
rem
echo build 
rem
cd ..\server

dotnet clean %solutionName%.sln

pause


dotnet build spider/spider.csproj --no-restore --configuration Debug --no-dependencies /property:Version=%versionNumber% /property:AssemblyVersion=%versionNumber% /property:FileVersion=%versionNumber%
if errorlevel 1 (
   echo failure building spider

   pause
   exit /b %errorlevel%
)

cd ..\scripts

pause

rem ==============================================================
rem
echo Build addon collection
rem

rem copy bin folder assemblies to collection folder
copy "%binPath%*.dll" "%collectionPath%"

pause

rem create new collection zip file
c:
cd %collectionPath%
del "%collectionName%.zip" /Q
"c:\program files\7-zip\7z.exe" a "%collectionName%.zip"
xcopy "%collectionName%.zip" "%deploymentFolderRoot%%versionNumber%" /Y
xcopy "%collectionName%.zip" "c:\deployments\_current_sprint" /Y
cd ..\..\scripts


rem remove files from the collection folder
del "%collectionPath%"\*.DLL

pause

