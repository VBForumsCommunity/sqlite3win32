/*
**
** Convenient concat() function which does certainly the same as the concatenation operator || to concatenate strings into one.
**
** Unlike the concatenation operator ||, the concat() function ignores NULL arguments.
**
** The two examples below will produce the same output 'SQLite concat'.
**
** SELECT 'SQLite ' || 'concat';
**
** SELECT concat('SQLite ', 'concat');
**
** The concat_ws() function concatenates strings into one separated by a particular separator. (ws stands for "with separator")
**
** SELECT concat_ws(', ', 'SQLite', 'concat');
**
** The output in above example will be 'SQLite, concat'
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
  if( argc < 1 )
    return;
  for(i=0; i < argc; i++)
  {
    lens[i] = sqlite3_value_bytes(argv[i]);
    lenall += lens[i];
  }
  all = sqlite3_malloc(lenall + 1);
  if( all == 0 ){
    sqlite3_result_error_nomem(context);
    return;
  }
  all[lenall] = '\0';
  for(i = 0; i < (argc - 1); i++)
  {
    if( lens[i] > 0 )
      memcpy(all + lencurrent, sqlite3_value_text(argv[i]), lens[i]);
    lencurrent += lens[i];
  }
  if( lens[argc - 1] > 0 )
    memcpy(all + lencurrent, sqlite3_value_text(argv[argc - 1]), lens[argc - 1]);
  sqlite3_result_text(context, all, -1, SQLITE_TRANSIENT);
  sqlite3_free(all);
}

static void concat_ws_sql_func(
  sqlite3_context *context, 
  int argc, 
  sqlite3_value **argv
){
  size_t lens[SQLITE_MAX_FUNCTION_ARG];
  size_t lenall = 0;
  size_t lencurrent = 0;
  size_t lensep = 0;
  int i = 0;
  char *all;
  if( argc < 2 )
    return;
  lensep = sqlite3_value_bytes(argv[0]);
  for(i=0; i < (argc - 1); i++)
  {
    lens[i] = sqlite3_value_bytes(argv[i + 1]);
    lenall += lens[i];
  }
  all = sqlite3_malloc(lenall + (lensep * (argc - 2)) + 1);
  if( all == 0 ){
    sqlite3_result_error_nomem(context);
    return;
  }
  all[lenall + (lensep * (argc - 2))] = '\0';
  for(i = 0; i < (argc - 2); i++)
  {
    if( lens[i] > 0 )
      memcpy(all + lencurrent + (i * lensep), sqlite3_value_text(argv[i + 1]), lens[i]);
    if( lensep > 0 )
      memcpy(all + lencurrent + lens[i] + (i * lensep), sqlite3_value_text(argv[0]), lensep);
    lencurrent += lens[i];
  }
  if( lens[argc - 2] > 0 )
    memcpy(all + lencurrent + ((argc - 2) * lensep), sqlite3_value_text(argv[argc - 1]), lens[argc - 2]);
  sqlite3_result_text(context, all, -1, SQLITE_TRANSIENT);
  sqlite3_free(all);
}

/*
** Invoke this routine to register the concat() function with the
** SQLite database connection.
*/
#ifdef _WIN32
__declspec(dllexport)
#endif
SQLITE_API int sqlite3_concat_init(
  sqlite3 *db, 
  char **pzErrMsg, 
  const sqlite3_api_routines *pApi
){
  int rc = SQLITE_OK;
  (void)pzErrMsg;  /* Unused parameter */
  SQLITE_EXTENSION_INIT2(pApi);
  rc = sqlite3_create_function(db, "concat", -1, SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS, 0, concat_sql_func, 0, 0);
  if( rc == SQLITE_OK )
  {
    rc = sqlite3_create_function(db, "concat_ws", -1, SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS, 0, concat_ws_sql_func, 0, 0);
  }
  return rc;
}
