@echo off
cd %~dp0

echo The vbac supports Word, Excel and Access Files. Which file do you want? 
set /p ext=File Extension (default xlsm): 
if "%ext%" neq "" (ren "src\Ariawase.xlsm" "Ariawase.%ext%")
if "%PROCESSOR_ARCHITECTURE%" neq "x86" (
    %windir%\SysWOW64\cscript //nologo vbac.wsf combine
) else {
    cscript //nologo vbac.wsf combine
}
