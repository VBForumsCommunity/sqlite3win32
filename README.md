The initial concept and core of this was developed by 'wqweto'. (https://github.com/wqweto/vbsqlite)

Out of this an Ax-DLL COM-Wrapper was created. (https://github.com/Kr00l/VBSQLite)

The fork of this has the following differences:
- Updated sqlite3 c source from 3011001 (2016-03-03) to 3039002 (2022-07-21).
- Win32 flags (SQLITE_WIN32_MALLOC, SQLITE_WIN32_HEAP_CREATE) are used in the c++ sources.
- Statically linking to the regexp c extension to enable the REGEXP operator. Use the exported function sqlite3_regexp_init() for registering it.
  Example: "rc = sqlite3_regexp_init(hDB, 0, 0)" (pass zero in the second and third parameter as not used)
- Statically linking to the concat c extension to enable the CONCAT() and CONCAT_WS() function. Use the exported function sqlite3_concat_init() for registering it.
- All callback functions are _stdcall. Even those used for the overloading of the SQL functions or aggregates.

What is the same:
- __stdcall convention applied in each exported function so that it can be easily used by VB6/VBA.
