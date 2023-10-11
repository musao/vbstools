@echo off
dir /b test_*.vbs | C:\Windows\SysWOW64\cscript.exe %~dp0..\bin\TestRunner.wsf //Job:ConsoleTestRunner /stdin+ %*
