/*
**
** Convenient concat() function which does the same as the concatenate operator || to join strings into one.
**
** The two examples below will produce the same output 'SQLite concat'.
**
** SELECT 'SQLite ' || 'concat';
**
** SELECT concat('SQLite ', 'concat');
**
*/
#include <string.h>
#include <stdlib.h>
/* #include "sqlite3ext.h" */
SQLITE_EXTENSION_INIT1

static void concat_sql_func(
  sqlite3_context *context, 
  int argc, 
  sqlite3_value **argv
){
  size_t lens[SQLITE_MAX_FUNCTION_ARG];
  size_t lenall = 0;
  size_t lencurrent = 0;
  int i = 0;
  char *all;
  for(i=0; i < argc; i++)
  {
    lens[i] = strlen(sqlite3_value_text(argv[i]));
    lenall += lens[i];
  }
  all = (char*)malloc(sizeof(argv) * lenall);
  all[lenall] = '\0';
  for(i = 0; i < argc - 1; i++)
  {
	memcpy(all + lencurrent, sqlite3_value_text(argv[i]), lens[i]);
    lencurrent += lens[i];
  }
  memcpy(all + lencurrent, sqlite3_value_text(argv[argc - 1]), lens[argc - 1]);
  sqlite3_result_text(context, all, strlen(all), SQLITE_TRANSIENT);
  free(all);
}

/*
** Invoke this routine to register the concat() function with the
** SQLite database connection.
*/
#ifdef _WIN32
__declspec(dllexport)
#endif
SQLITE_API int SQLITE_STDCALL sqlite3_concat_init(
  sqlite3 *db, 
  char **pzErrMsg, 
  const sqlite3_api_routines *pApi
){
  int rc = SQLITE_OK;
  (void)pzErrMsg;  /* Unused parameter */
  SQLITE_EXTENSION_INIT2(pApi);
  rc = sqlite3_create_function(db, "concat", -1, SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS, 0,
                                 concat_sql_func, 0, 0);
  return rc;
}
