@echo off
setlocal
set cl_exe="C:\Programme\Microsoft Visual Studio 6.0\VC98\Bin\cl.exe" /nologo /MD
set bin_dir=..\Bin

pushd %~dp0

%cl_exe% /LD sqlite3win32stubs.cpp sqlite3win32helper.c /Fesqlite3win32.dll /link /DEF:sqlite3win32.def
copy sqlite3win32.dll %bin_dir% > nul
copy sqlite3win32stubs.obj %bin_dir%\*.cobj > nul
copy sqlite3win32helper.obj %bin_dir%\*.cobj > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd

pause