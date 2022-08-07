@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\cl.exe" /nologo /MT /O2 /Gz
:: /arch:IA32

set bin_dir=..\Bin

call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\vcvars32.bat"

set "include=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.1A\include;%include%"
set "lib=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.1A\lib;%lib%"

pushd %~dp0

%cl_exe% /LD sqlite3win32stubs.cpp sqlite3win32helper.c /Fesqlite3win32.dll /link /DEF:sqlite3win32.def
copy sqlite3win32.dll %bin_dir% > nul
copy sqlite3win32stubs.obj %bin_dir%\*.cobj > nul
copy sqlite3win32helper.obj %bin_dir%\*.cobj > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd

pause