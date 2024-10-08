@echo off
setlocal
set cl_exe="C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\amd64\cl.exe" /nologo /c /Gw /Gy /GS-
set bin_dir=..\Bin
set bin_file=sqlite3win64.obj

call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\amd64\vcvars64.bat"

pushd %~dp0

%cl_exe% sqlite3win32helper.c /Fo%bin_file%
copy %bin_file% %bin_dir% > nul

:cleanup
del /q *.exp *.lib *.obj *.dll ~$*

popd

pause