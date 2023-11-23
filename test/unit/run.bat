@echo off
setlocal
set testrunner=..\..\bin\TestRunner.wsf
set cscript=C:\Windows\SysWow64\cscript.exe
dir /b test_*.vbs | %cscript% %testrunner% //Job:ConsoleTestRunner /stdin+ %*
endlocal