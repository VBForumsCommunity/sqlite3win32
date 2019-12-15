// # pragma warning (disable : 4018)  // signed/unsigned mismatch
# pragma warning (disable : 4022)  // pointer mismatch for actual parameter x
# pragma warning (disable : 4028)  // formal parameter x different from declaration
# pragma warning (disable : 4047)  // 'void ** ' differs in levels of indirection from 'long *'
# pragma warning (disable : 4049)  // terminating line number emission
// # pragma warning (disable : 4056)  // overflow in floating-point constant arithmetic
# pragma warning (disable : 4090)  // 'function' : different 'volatile' qualifiers
// # pragma warning (disable : 4100)  // unreferenced formal parameter
// # pragma warning (disable : 4101)  // unreferenced local variable
// # pragma warning (disable : 4115)  // named type definition in parentheses
// # pragma warning (disable : 4127)  // conditional expression is constant
// # pragma warning (disable : 4132)  // const object should be initialized
// # pragma warning (disable : 4232)  // address of dllimport 'x' is not static, identity not guaranteed
// # pragma warning (disable : 4244)  // conversion from 'int ' to 'char ', possible loss of data
// # pragma warning (disable : 4305)  // truncation from 'int ' to 'short '
# pragma warning (disable : 4550)  // expression evaluates to a function which is missing an argument list 
// # pragma warning (disable : 4701)  // 'xxx' may be used without having been initialized
// # pragma warning (disable : 4702)  // unreachable code
// # pragma warning (disable : 4706)  // assignment within conditional expression
// # pragma warning (disable : 4756)  // overflow in constant arithmetic
# pragma warning (disable : 4761)  // integral size mismatch in argument; conversion supplied

#define SQLITE_API
#define SQLITE_STDCALL __stdcall
#define SQLITE_WIN32_MALLOC 1
#define SQLITE_WIN32_HEAP_CREATE 1
#define SQLITE_THREADSAFE 2
// #define SQLITE_USE_URI 1
#define SQLITE_SOUNDEX 1
#define SQLITE_ENABLE_FTS3 1
#define SQLITE_ENABLE_FTS3_PARENTHESIS 1
#define SQLITE_ENABLE_FTS4 1
#define SQLITE_ENABLE_RTREE 1
#define SQLITE_ENABLE_UNLOCK_NOTIFY 1
#define SQLITE_ENABLE_COLUMN_METADATA 1

#include "sqlite3\sqlite3.c"
