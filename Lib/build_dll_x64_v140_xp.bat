@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\x86_amd64\cl.exe" /nologo /MT /O2 /Gz
set bin_dir=..\Bin
set bin_file=sqlite3win64.dll

call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\x86_amd64\vcvarsx86_amd64.bat"

set "include=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.1A\include;%include%"
set "lib=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.1A\lib\x64;%lib%"

pushd %~dp0

%cl_exe% /LD sqlite3win32helper.c /Fe%bin_file% /link /MACHINE:X64 /SUBSYSTEM:WINDOWS,5.02 /DEF:sqlite3win64.def
copy %bin_file% %bin_dir% > nul
copy *.obj %bin_dir%\*.cobj > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd

pause