@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\cl.exe" /nologo /MD /O2 /Og- /Gz
set bin_dir=..\Bin
set bin_file=sqlite3win32.dll

call "C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\vcvars32.bat"

pushd %~dp0

%cl_exe% /LD sqlite3win32stubs.cpp sqlite3win32helper.c /Fe%bin_file% /link /DEF:sqlite3win32.def
copy %bin_file% %bin_dir% > nul
copy *.obj %bin_dir%\*.cobj > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd

pause