@echo off
setlocal
set testrunner=%USERPROFILE%\Documents\dev\vbs\bin\TestRunner.wsf
set cscript=%windir%\SysWow64\cscript.exe
dir /b test_*.vbs | %cscript% %testrunner% //Job:ConsoleTestRunner /stdin+ %*
endlocal