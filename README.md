The initial concept and core of this was developed by 'wqweto'. (https://github.com/wqweto/vbsqlite)

Out of this a Ax-DLL COM-Wrapper (VBSQLite10) was created. (https://github.com/Kr00l/VBSQLite)

The fork of this has the following differences:
- Updated sqlite3 c source from 3011001 (2016-03-03) to 3024000 (2018-06-04).
- Staticly compiled sqlite3 into a VB6 Std-EXE instead of a VB6 ActiveX DLL.
- Win32 flags (SQLITE_WIN32_MALLOC, SQLITE_WIN32_HEAP_CREATE) are used in the c++ sources.
- Renamings in the c++ sources.
- No use of a link.exe replacement (as it caused some problem on other project compilations) to swap cobj files, instead an add-in is used to intercept the linking event. (http://www.vbforums.com/showthread.php?866321-VB6-IDE-Linker-AddIn)

What is the same:
- __stdcall convention applied in each exported function and callback routines so that it can be easily used by VB6.
