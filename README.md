The initial concept and core of this was developed by 'wqweto'. (https://github.com/wqweto/vbsqlite)

Out of this an Ax-DLL COM-Wrapper was created. (https://github.com/Kr00l/VBSQLite)

The fork of this has the following differences:
- Updated sqlite3 c source from 3011001 (2016-03-03) to 3024000 (2018-06-04).
- Win32 flags (SQLITE_WIN32_MALLOC, SQLITE_WIN32_HEAP_CREATE) are used in the c++ sources.
- Renamings in the c++ sources.

What is the same:
- __stdcall convention applied in each exported function and callback routines so that it can be easily used by VB6.
