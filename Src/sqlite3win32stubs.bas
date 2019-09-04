Attribute VB_Name = "sqlite3win32stubs"
Option Explicit
Private Declare Function sqlite3_aggregate_context Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_aggregate_count Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_auto_extension Lib "sqlite3win32.dll" (ByVal xEntryPoint As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "sqlite3win32.dll" (ByVal pBak As Long) As Long
Private Declare Function sqlite3_backup_init Lib "sqlite3win32.dll" (ByVal pDest As Long, ByVal pzDestName As Long, ByVal pSrc As Long, ByVal pzSrcName As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "sqlite3win32.dll" (ByVal pBak As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "sqlite3win32.dll" (ByVal pBak As Long) As Long
Private Declare Function sqlite3_backup_step Lib "sqlite3win32.dll" (ByVal pBak As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_bind_blob64 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Currency, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_bind_double Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal rValue As Double) As Long
Private Declare Function sqlite3_bind_int Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal iValue As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal iValue As Currency) As Long
Private Declare Function sqlite3_bind_null Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long) As Long
Private Declare Function sqlite3_bind_parameter_count Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal pzName As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long) As Long
Private Declare Function sqlite3_bind_pointer Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pPtr As Long, ByVal pzPType As Long, ByVal lpfnDestroy As Long)
Private Declare Function sqlite3_bind_text Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_bind_text16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_bind_text64 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Currency, ByVal lpfnDestroy As Long, ByVal Encoding As Byte) As Long
Private Declare Function sqlite3_bind_value Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal pValue As Long) As Long
Private Declare Function sqlite3_bind_zeroblob Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_bind_zeroblob64 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long, ByVal n As Currency) As Long
Private Declare Function sqlite3_blob_bytes Lib "sqlite3win32.dll" (ByVal pBlob As Long) As Long
Private Declare Function sqlite3_blob_close Lib "sqlite3win32.dll" (ByVal pBlob As Long) As Long
Private Declare Function sqlite3_blob_open Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDB As Long, ByVal pzTable As Long, ByVal pzColumn As Long, ByVal iRow As Currency, ByVal Flags As Long, ByRef pBlob As Long) As Long
Private Declare Function sqlite3_blob_read Lib "sqlite3win32.dll" (ByVal pBlob As Long, ByVal pz As Long, ByVal n As Long, ByVal iOffset As Long) As Long
Private Declare Function sqlite3_blob_reopen Lib "sqlite3win32.dll" (ByVal pBlob As Long, ByVal iRow As Currency) As Long
Private Declare Function sqlite3_blob_write Lib "sqlite3win32.dll" (ByVal pBlob As Long, ByVal pz As Long, ByVal n As Long, ByVal iOffset As Long) As Long
Private Declare Function sqlite3_busy_handler Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal xBusy As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_busy_timeout Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function sqlite3_cancel_auto_extension Lib "sqlite3win32.dll" (ByVal xEntryPoint As Long) As Long
Private Declare Function sqlite3_changes Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_close Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_close_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_collation_needed Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pCollNeededArg As Long, ByVal xCollNeeded As Long) As Long
Private Declare Function sqlite3_collation_needed16 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pCollNeededArg As Long, ByVal xCollNeeded16 As Long) As Long
Private Declare Function sqlite3_column_blob Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long) As Long
Private Declare Function sqlite3_column_bytes Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long) As Long
Private Declare Function sqlite3_column_bytes16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal i As Long) As Long
Private Declare Function sqlite3_column_count Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_database_name Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_database_name16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_decltype Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_decltype16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_name Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_origin_name Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_origin_name16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_table_name Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_table_name16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_text Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_text16 Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_type Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_value Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_commit_hook Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_compileoption_get Lib "sqlite3win32.dll" (ByVal n As Long) As Long
Private Declare Function sqlite3_compileoption_used Lib "sqlite3win32.dll" (ByVal pzOptName As Long) As Long
Private Declare Function sqlite3_complete Lib "sqlite3win32.dll" (ByVal pzSQL As Long) As Long
Private Declare Function sqlite3_complete16 Lib "sqlite3win32.dll" (ByVal pzSQL As Long) As Long
Private Declare Function sqlite3_context_db_handle Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_create_collation Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long) As Long
Private Declare Function sqlite3_create_collation_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_create_collation16 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long) As Long
Private Declare Function sqlite3_create_function Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzFunc As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long) As Long
Private Declare Function sqlite3_create_function_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzFunc As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_create_function16 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzFunctionName As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long) As Long
Private Declare Function sqlite3_create_module Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal pModule As Long, ByVal pAux As Long) As Long
Private Declare Function sqlite3_create_module_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal pModule As Long, ByVal pAux As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_data_count Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_db_cacheflush Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_db_filename Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDbName As Long) As Long
Private Declare Function sqlite3_db_handle Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_db_mutex Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_db_readonly Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDbName As Long) As Long
Private Declare Function sqlite3_db_release_memory Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_db_status Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal StatusOpt As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
Private Declare Function sqlite3_declare_vtab Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzCreateTable As Long) As Long
Private Declare Function sqlite3_enable_load_extension Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal iOnOff As Long) As Long
Private Declare Function sqlite3_enable_shared_cache Lib "sqlite3win32.dll" (ByVal fEnable As Long) As Long
Private Declare Function sqlite3_errcode Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_errmsg Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_errmsg16 Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_errstr Lib "sqlite3win32.dll" (ByVal ResultCode As Long) As Long
Private Declare Function sqlite3_exec Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal lpfnCallback As Long, ByVal pArg As Long, ByVal pzErrMsg As Long) As Long
Private Declare Function sqlite3_expanded_sql Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_expired Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_extended_result_codes Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal iOnOff As Long) As Long
Private Declare Function sqlite3_file_control Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDbName As Long, ByVal Code As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_finalize Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_free Lib "sqlite3win32.dll" (ByVal pMem As Long) As Long
Private Declare Function sqlite3_free_table Lib "sqlite3win32.dll" (ByVal azResult As Long) As Long
Private Declare Function sqlite3_get_autocommit Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_get_auxdata Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal iArg As Long) As Long
Private Declare Function sqlite3_get_table Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal pazResult As Long, ByVal pnRow As Long, ByVal pnColumn As Long, ByVal pzErrMsg As Long) As Long
Private Declare Function sqlite3_global_recover Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_initialize Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_interrupt Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_keyword_check Lib "sqlite3win32.dll" (ByVal pzName As Long, ByVal pnName As Long) As Long
Private Declare Function sqlite3_keyword_count Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_keyword_name Lib "sqlite3win32.dll" (ByVal i As Long, ByVal pzName As Long, ByVal pnName As Long) As Long
Private Declare Function sqlite3_last_insert_rowid Lib "sqlite3win32.dll" (ByVal hDB As Long) As Currency
Private Declare Function sqlite3_libversion Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_libversion_number Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_limit Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal LimitID As Long, ByVal NewLimit As Long) As Long
Private Declare Function sqlite3_load_extension Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzFile As Long, ByVal pzProc As Long, ByVal pzErrMsg As Long) As Long
Private Declare Function sqlite3_malloc Lib "sqlite3win32.dll" (ByVal n As Long) As Long
Private Declare Function sqlite3_malloc64 Lib "sqlite3win32.dll" (ByVal n As Currency) As Long
Private Declare Function sqlite3_memory_alarm Lib "sqlite3win32.dll" (ByVal lpfnCallback As Long, ByVal pArg As Long, ByVal iThreshold As Currency) As Long
Private Declare Function sqlite3_memory_highwater Lib "sqlite3win32.dll" (ByVal ResetFlag As Long) As Currency
Private Declare Function sqlite3_memory_used Lib "sqlite3win32.dll" () As Currency
Private Declare Function sqlite3_msize Lib "sqlite3win32.dll" (ByVal pMem As Long) As Currency
Private Declare Function sqlite3_mutex_alloc Lib "sqlite3win32.dll" (ByVal pMtx As Long) As Long
Private Declare Function sqlite3_mutex_enter Lib "sqlite3win32.dll" (ByVal pMtx As Long) As Long
Private Declare Function sqlite3_mutex_free Lib "sqlite3win32.dll" (ByVal pMtx As Long) As Long
Private Declare Function sqlite3_mutex_leave Lib "sqlite3win32.dll" (ByVal pMtx As Long) As Long
Private Declare Function sqlite3_mutex_try Lib "sqlite3win32.dll" (ByVal pMtx As Long) As Long
Private Declare Function sqlite3_next_stmt Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal hStmt As Long) As Long
Private Declare Function sqlite3_open Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByRef hDB As Long) As Long
Private Declare Function sqlite3_open_v2 Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByRef hDB As Long, ByVal Flags As Long, ByVal pzVfs As Long) As Long
Private Declare Function sqlite3_open16 Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByRef hDB As Long) As Long
Private Declare Function sqlite3_os_end Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_os_init Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_overload_function Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzName As Long, ByVal nArg As Long) As Long
Private Declare Function sqlite3_prepare Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_prepare_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_prepare_v3 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByVal PrepFlags As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_prepare16 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_prepare16_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_prepare16_v3 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByVal PrepFlags As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
Private Declare Function sqlite3_profile Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnProfile As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_progress_handler Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal nOps As Long, ByVal lpfnProgress As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_randomness Lib "sqlite3win32.dll" (ByVal n As Long, ByVal pBuf As Long) As Long
Private Declare Function sqlite3_realloc Lib "sqlite3win32.dll" (ByVal pOld As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_realloc64 Lib "sqlite3win32.dll" (ByVal pOld As Long, ByVal n As Currency) As Long
Private Declare Function sqlite3_release_memory Lib "sqlite3win32.dll" (ByVal n As Long) As Long
Private Declare Function sqlite3_reset Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_reset_auto_extension Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_result_blob Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_blob64 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Currency, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_double Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal rValue As Double) As Long
Private Declare Function sqlite3_result_error Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_result_error_code Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal ErrCode As Long) As Long
Private Declare Function sqlite3_result_error_nomem Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_result_error_toobig Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_result_error16 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_result_int Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal iValue As Long) As Long
Private Declare Function sqlite3_result_int64 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal iValue As Currency) As Long
Private Declare Function sqlite3_result_null Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_result_subtype Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal eSubtype As Long) As Long
Private Declare Function sqlite3_result_text Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_text16 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_text16be Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_text16le Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_result_text64 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pz As Long, ByVal n As Currency, ByVal lpfnDestroy As Long, ByVal Encoding As Byte) As Long
Private Declare Function sqlite3_result_value Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pValue As Long) As Long
Private Declare Function sqlite3_result_pointer Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal pPtr As Long, ByVal pzPType As Long, ByVal lpfnDestroy As Long)
Private Declare Function sqlite3_result_zeroblob Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_result_zeroblob64 Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal n As Currency) As Long
Private Declare Function sqlite3_rollback_hook Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_rtree_geometry_callback Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzGeom As Long, ByVal lpfnGeom As Long, ByVal pContext As Long) As Long
Private Declare Function sqlite3_rtree_query_callback Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzQueryFunc As Long, ByVal lpfnQueryFunc As Long, ByVal pContext As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_set_authorizer Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal xAuth As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_set_auxdata Lib "sqlite3win32.dll" (ByVal pCtx As Long, ByVal iArg As Long, ByVal pAux As Long, ByVal lpfnDestroy As Long) As Long
Private Declare Function sqlite3_set_last_insert_rowid Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal iRow As Currency) As Long
Private Declare Function sqlite3_shutdown Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_sleep Lib "sqlite3win32.dll" (ByVal dwMilliseconds As Long) As Long
Private Declare Function sqlite3_soft_heap_limit Lib "sqlite3win32.dll" (ByVal n As Long) As Long
Private Declare Function sqlite3_soft_heap_limit64 Lib "sqlite3win32.dll" (ByVal n As Currency) As Currency
Private Declare Function sqlite3_sourceid Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_sql Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
'Private Declare Function sqlite3_str_new Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_finish Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_vappendf Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_append Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_appendall Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_appendchar Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_reset Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_errcode Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_length Lib "sqlite3win32.dll"
'Private Declare Function sqlite3_str_value Lib "sqlite3win32.dll"
Private Declare Function sqlite3_status Lib "sqlite3win32.dll" (ByVal Code As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
Private Declare Function sqlite3_status64 Lib "sqlite3win32.dll" (ByVal Code As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
Private Declare Function sqlite3_step Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stmt_busy Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stmt_readonly Lib "sqlite3win32.dll" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stmt_status Lib "sqlite3win32.dll" (ByVal hStmt As Long, ByVal Code As Long, ByVal ResetFlag As Long) As Long
Private Declare Function sqlite3_strglob Lib "sqlite3win32.dll" (ByVal pzGlobPattern As Long, ByVal pzString As Long) As Long
Private Declare Function sqlite3_stricmp Lib "sqlite3win32.dll" (ByVal pzLeft As Long, ByVal pzRight As Long) As Long
Private Declare Function sqlite3_strlike Lib "sqlite3win32.dll" (ByVal pzPattern As Long, ByVal pzStr As Long, ByVal cEsc As Long) As Long
Private Declare Function sqlite3_strnicmp Lib "sqlite3win32.dll" (ByVal pzLeft As Long, ByVal pzRight As Long, ByVal n As Long) As Long
Private Declare Function sqlite3_system_errno Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_table_column_metadata Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDbName As Long, ByVal pzTableName As Long, ByVal pzColumnName As Long, ByVal pzDataType As Long, ByVal pzCollSeq As Long, ByVal pNotNull As Long, ByVal pPrimaryKey As Long, ByVal pAutoinc As Long) As Long
Private Declare Function sqlite3_thread_cleanup Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_threadsafe Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_total_changes Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_trace Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnTrace As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_trace_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal uMask As Long, ByVal lpfnCallback As Long, ByVal pCtx As Long) As Long
Private Declare Function sqlite3_transfer_bindings Lib "sqlite3win32.dll" (ByVal hStmtFrom As Long, ByVal hStmtTo As Long) As Long
Private Declare Function sqlite3_update_hook Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_uri_boolean Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByVal pzParam As Long, ByVal bDefault As Long) As Long
Private Declare Function sqlite3_uri_int64 Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByVal pzParam As Long, ByVal bDefault As Currency) As Currency
Private Declare Function sqlite3_uri_parameter Lib "sqlite3win32.dll" (ByVal pzFilename As Long, ByVal pzParam As Long) As Long
Private Declare Function sqlite3_user_data Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_value_blob Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_bytes Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_bytes16 Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_double Lib "sqlite3win32.dll" (ByVal pValue As Long) As Double
Private Declare Function sqlite3_value_dup Lib "sqlite3win32.dll" (ByVal pOrig As Long) As Long
Private Declare Function sqlite3_value_free Lib "sqlite3win32.dll" (ByVal pOld As Long) As Long
Private Declare Function sqlite3_value_int Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_int64 Lib "sqlite3win32.dll" (ByVal pValue As Long) As Currency
Private Declare Function sqlite3_value_pointer Lib "sqlite3win32.dll" (ByVal pValue As Long, ByVal pzPType As Long) As Long
Private Declare Function sqlite3_value_numeric_type Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_nochange Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_subtype Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_text Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_text16 Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_text16be Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_text16le Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_value_type Lib "sqlite3win32.dll" (ByVal pValue As Long) As Long
Private Declare Function sqlite3_vfs_find Lib "sqlite3win32.dll" (ByVal pzVfs As Long) As Long
Private Declare Function sqlite3_vfs_register Lib "sqlite3win32.dll" (ByVal pVfs As Long, ByVal MakeDefault As Long) As Long
Private Declare Function sqlite3_vfs_unregister Lib "sqlite3win32.dll" (ByVal pVfs As Long) As Long
Private Declare Function sqlite3_vmprintf Lib "sqlite3win32.dll" (ByVal pzFormat As Long, ByVal va_list As Long) As Long
Private Declare Function sqlite3_vsnprintf Lib "sqlite3win32.dll" (ByVal n As Long, ByVal pzBuffer As Long, ByVal pzFormat As Long, ByVal va_list As Long) As Long
Private Declare Function sqlite3_vtab_collation Lib "sqlite3win32.dll" (ByVal pIdxInfo As Long, ByVal iCons As Long)
Private Declare Function sqlite3_vtab_nochange Lib "sqlite3win32.dll" (ByVal pCtx As Long) As Long
Private Declare Function sqlite3_vtab_on_conflict Lib "sqlite3win32.dll" (ByVal hDB As Long) As Long
Private Declare Function sqlite3_wal_autocheckpoint Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal nFrame As Long) As Long
Private Declare Function sqlite3_wal_checkpoint Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDB As Long) As Long
Private Declare Function sqlite3_wal_checkpoint_v2 Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal pzDB As Long, ByVal eMode As Long, ByVal pnLog As Long, ByVal pnCkpt As Long) As Long
Private Declare Function sqlite3_wal_hook Lib "sqlite3win32.dll" (ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
Private Declare Function sqlite3_win32_is_nt Lib "sqlite3win32.dll" () As Long
Private Declare Function sqlite3_win32_mbcs_to_utf8 Lib "sqlite3win32.dll" (ByVal pzFilename As Long) As Long
Private Declare Function sqlite3_win32_set_directory Lib "sqlite3win32.dll" (ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
Private Declare Function sqlite3_win32_set_directory8 Lib "sqlite3win32.dll" (ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
Private Declare Function sqlite3_win32_set_directory16 Lib "sqlite3win32.dll" (ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
Private Declare Function sqlite3_win32_sleep Lib "sqlite3win32.dll" (ByVal dwMilliseconds As Long) As Long
Private Declare Function sqlite3_win32_utf8_to_mbcs Lib "sqlite3win32.dll" (ByVal pzFilename As Long) As Long
Private Declare Function sqlite3_win32_write_debug Lib "sqlite3win32.dll" (ByVal pzBuffer As Long, ByVal nBuffer As Long) As Long

Public Function stub_sqlite3_aggregate_context(ByVal pCtx As Long, ByVal nBytes As Long) As Long
stub_sqlite3_aggregate_context = sqlite3_aggregate_context(pCtx, nBytes)
End Function

Public Function stub_sqlite3_aggregate_count(ByVal pCtx As Long) As Long
stub_sqlite3_aggregate_count = sqlite3_aggregate_count(pCtx)
End Function

Public Function stub_sqlite3_auto_extension(ByVal xEntryPoint As Long) As Long
stub_sqlite3_auto_extension = sqlite3_auto_extension(xEntryPoint)
End Function

Public Function stub_sqlite3_backup_finish(ByVal pBak As Long) As Long
stub_sqlite3_backup_finish = sqlite3_backup_finish(pBak)
End Function

Public Function stub_sqlite3_backup_init(ByVal pDest As Long, ByVal pzDestName As Long, ByVal pSrc As Long, ByVal pzSrcName As Long) As Long
stub_sqlite3_backup_init = sqlite3_backup_init(pDest, pzDestName, pSrc, pzSrcName)
End Function

Public Function stub_sqlite3_backup_pagecount(ByVal pBak As Long) As Long
stub_sqlite3_backup_pagecount = sqlite3_backup_pagecount(pBak)
End Function

Public Function stub_sqlite3_backup_remaining(ByVal pBak As Long) As Long
stub_sqlite3_backup_remaining = sqlite3_backup_remaining(pBak)
End Function

Public Function stub_sqlite3_backup_step(ByVal pBak As Long, ByVal nPage As Long) As Long
stub_sqlite3_backup_step = sqlite3_backup_step(pBak, nPage)
End Function

Public Function stub_sqlite3_bind_blob(ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_bind_blob = sqlite3_bind_blob(hStmt, i, pzData, nData, lpfnDestroy)
End Function

Public Function stub_sqlite3_bind_blob64(ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Currency, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_bind_blob64 = sqlite3_bind_blob64(hStmt, i, pzData, nData, lpfnDestroy)
End Function

Public Function stub_sqlite3_bind_double(ByVal hStmt As Long, ByVal i As Long, ByVal rValue As Double) As Long
stub_sqlite3_bind_double = sqlite3_bind_double(hStmt, i, rValue)
End Function

Public Function stub_sqlite3_bind_int(ByVal hStmt As Long, ByVal i As Long, ByVal iValue As Long) As Long
stub_sqlite3_bind_int = sqlite3_bind_int(hStmt, i, iValue)
End Function

Public Function stub_sqlite3_bind_int64(ByVal hStmt As Long, ByVal i As Long, ByVal iValue As Currency) As Long
stub_sqlite3_bind_int64 = sqlite3_bind_int64(hStmt, i, iValue)
End Function

Public Function stub_sqlite3_bind_null(ByVal hStmt As Long, ByVal i As Long) As Long
stub_sqlite3_bind_null = sqlite3_bind_null(hStmt, i)
End Function

Public Function stub_sqlite3_bind_parameter_count(ByVal hStmt As Long) As Long
stub_sqlite3_bind_parameter_count = sqlite3_bind_parameter_count(hStmt)
End Function

Public Function stub_sqlite3_bind_parameter_index(ByVal hStmt As Long, ByVal pzName As Long) As Long
stub_sqlite3_bind_parameter_index = sqlite3_bind_parameter_index(hStmt, pzName)
End Function

Public Function stub_sqlite3_bind_parameter_name(ByVal hStmt As Long, ByVal i As Long) As Long
stub_sqlite3_bind_parameter_name = sqlite3_bind_parameter_name(hStmt, i)
End Function

Public Function stub_sqlite3_bind_pointer(ByVal hStmt As Long, ByVal i As Long, ByVal pPtr As Long, ByVal pzPType As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_bind_pointer = sqlite3_bind_pointer(hStmt, i, pPtr, pzPType, lpfnDestroy)
End Function

Public Function stub_sqlite3_bind_text(ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_bind_text = sqlite3_bind_text(hStmt, i, pzData, nData, lpfnDestroy)
End Function

Public Function stub_sqlite3_bind_text16(ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_bind_text16 = sqlite3_bind_text16(hStmt, i, pzData, nData, lpfnDestroy)
End Function

Public Function stub_sqlite3_bind_text64(ByVal hStmt As Long, ByVal i As Long, ByVal pzData As Long, ByVal nData As Currency, ByVal lpfnDestroy As Long, ByVal Encoding As Byte) As Long
stub_sqlite3_bind_text64 = sqlite3_bind_text64(hStmt, i, pzData, nData, lpfnDestroy, Encoding)
End Function

Public Function stub_sqlite3_bind_value(ByVal hStmt As Long, ByVal i As Long, ByVal pValue As Long) As Long
stub_sqlite3_bind_value = sqlite3_bind_value(hStmt, i, pValue)
End Function

Public Function stub_sqlite3_bind_zeroblob(ByVal hStmt As Long, ByVal i As Long, ByVal n As Long) As Long
stub_sqlite3_bind_zeroblob = sqlite3_bind_zeroblob(hStmt, i, n)
End Function

Public Function stub_sqlite3_bind_zeroblob64(ByVal hStmt As Long, ByVal i As Long, ByVal n As Currency) As Long
stub_sqlite3_bind_zeroblob64 = sqlite3_bind_zeroblob64(hStmt, i, n)
End Function

Public Function stub_sqlite3_blob_bytes(ByVal pBlob As Long) As Long
stub_sqlite3_blob_bytes = sqlite3_blob_bytes(pBlob)
End Function

Public Function stub_sqlite3_blob_close(ByVal pBlob As Long) As Long
stub_sqlite3_blob_close = sqlite3_blob_close(pBlob)
End Function

Public Function stub_sqlite3_blob_open(ByVal hDB As Long, ByVal pzDB As Long, ByVal pzTable As Long, ByVal pzColumn As Long, ByVal iRow As Currency, ByVal Flags As Long, ByRef pBlob As Long) As Long
stub_sqlite3_blob_open = sqlite3_blob_open(hDB, pzDB, pzTable, pzColumn, iRow, Flags, pBlob)
End Function

Public Function stub_sqlite3_blob_read(ByVal pBlob As Long, ByVal pz As Long, ByVal n As Long, ByVal iOffset As Long) As Long
stub_sqlite3_blob_read = sqlite3_blob_read(pBlob, pz, n, iOffset)
End Function

Public Function stub_sqlite3_blob_reopen(ByVal pBlob As Long, ByVal iRow As Currency) As Long
stub_sqlite3_blob_reopen = sqlite3_blob_reopen(pBlob, iRow)
End Function

Public Function stub_sqlite3_blob_write(ByVal pBlob As Long, ByVal pz As Long, ByVal n As Long, ByVal iOffset As Long) As Long
stub_sqlite3_blob_write = sqlite3_blob_write(pBlob, pz, n, iOffset)
End Function

Public Function stub_sqlite3_busy_handler(ByVal hDB As Long, ByVal xBusy As Long, ByVal pArg As Long) As Long
stub_sqlite3_busy_handler = sqlite3_busy_handler(hDB, xBusy, pArg)
End Function

Public Function stub_sqlite3_busy_timeout(ByVal hDB As Long, ByVal dwMilliseconds As Long) As Long
stub_sqlite3_busy_timeout = sqlite3_busy_timeout(hDB, dwMilliseconds)
End Function

Public Function stub_sqlite3_cancel_auto_extension(ByVal xEntryPoint As Long) As Long
stub_sqlite3_cancel_auto_extension = sqlite3_cancel_auto_extension(xEntryPoint)
End Function

Public Function stub_sqlite3_changes(ByVal hDB As Long) As Long
stub_sqlite3_changes = sqlite3_changes(hDB)
End Function

Public Function stub_sqlite3_clear_bindings(ByVal hStmt As Long) As Long
stub_sqlite3_clear_bindings = sqlite3_clear_bindings(hStmt)
End Function

Public Function stub_sqlite3_close(ByVal hDB As Long) As Long
stub_sqlite3_close = sqlite3_close(hDB)
End Function

Public Function stub_sqlite3_close_v2(ByVal hDB As Long) As Long
stub_sqlite3_close_v2 = sqlite3_close_v2(hDB)
End Function

Public Function stub_sqlite3_collation_needed(ByVal hDB As Long, ByVal pCollNeededArg As Long, ByVal xCollNeeded As Long) As Long
stub_sqlite3_collation_needed = sqlite3_collation_needed(hDB, pCollNeededArg, xCollNeeded)
End Function

Public Function stub_sqlite3_collation_needed16(ByVal hDB As Long, ByVal pCollNeededArg As Long, ByVal xCollNeeded16 As Long) As Long
stub_sqlite3_collation_needed16 = sqlite3_collation_needed16(hDB, pCollNeededArg, xCollNeeded16)
End Function

Public Function stub_sqlite3_column_blob(ByVal hStmt As Long, ByVal i As Long) As Long
stub_sqlite3_column_blob = sqlite3_column_blob(hStmt, i)
End Function

Public Function stub_sqlite3_column_bytes(ByVal hStmt As Long, ByVal i As Long) As Long
stub_sqlite3_column_bytes = sqlite3_column_bytes(hStmt, i)
End Function

Public Function stub_sqlite3_column_bytes16(ByVal hStmt As Long, ByVal i As Long) As Long
stub_sqlite3_column_bytes16 = sqlite3_column_bytes16(hStmt, i)
End Function

Public Function stub_sqlite3_column_count(ByVal hStmt As Long) As Long
stub_sqlite3_column_count = sqlite3_column_count(hStmt)
End Function

Public Function stub_sqlite3_column_database_name(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_database_name = sqlite3_column_database_name(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_database_name16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_database_name16 = sqlite3_column_database_name16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_decltype(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_decltype = sqlite3_column_decltype(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_decltype16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_decltype16 = sqlite3_column_decltype16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_double(ByVal hStmt As Long, ByVal iCol As Long) As Double
stub_sqlite3_column_double = sqlite3_column_double(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_int(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_int = sqlite3_column_int(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_int64(ByVal hStmt As Long, ByVal iCol As Long) As Currency
stub_sqlite3_column_int64 = sqlite3_column_int64(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_name(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_name = sqlite3_column_name(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_name16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_name16 = sqlite3_column_name16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_origin_name(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_origin_name = sqlite3_column_origin_name(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_origin_name16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_origin_name16 = sqlite3_column_origin_name16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_table_name(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_table_name = sqlite3_column_table_name(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_table_name16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_table_name16 = sqlite3_column_table_name16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_text(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_text = sqlite3_column_text(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_text16(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_text16 = sqlite3_column_text16(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_type(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_type = sqlite3_column_type(hStmt, iCol)
End Function

Public Function stub_sqlite3_column_value(ByVal hStmt As Long, ByVal iCol As Long) As Long
stub_sqlite3_column_value = sqlite3_column_value(hStmt, iCol)
End Function

Public Function stub_sqlite3_commit_hook(ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
stub_sqlite3_commit_hook = sqlite3_commit_hook(hDB, lpfnCallback, pArg)
End Function

Public Function stub_sqlite3_compileoption_get(ByVal n As Long) As Long
stub_sqlite3_compileoption_get = sqlite3_compileoption_get(n)
End Function

Public Function stub_sqlite3_compileoption_used(ByVal pzOptName As Long) As Long
stub_sqlite3_compileoption_used = sqlite3_compileoption_used(pzOptName)
End Function

Public Function stub_sqlite3_complete(ByVal pzSQL As Long) As Long
stub_sqlite3_complete = sqlite3_complete(pzSQL)
End Function

Public Function stub_sqlite3_complete16(ByVal pzSQL As Long) As Long
stub_sqlite3_complete16 = sqlite3_complete16(pzSQL)
End Function

Public Function stub_sqlite3_context_db_handle(ByVal pCtx As Long) As Long
stub_sqlite3_context_db_handle = sqlite3_context_db_handle(pCtx)
End Function

Public Function stub_sqlite3_create_collation(ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long) As Long
stub_sqlite3_create_collation = sqlite3_create_collation(hDB, pzName, eTextRep, pArg, lpfnCompare)
End Function

Public Function stub_sqlite3_create_collation_v2(ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_create_collation_v2 = sqlite3_create_collation_v2(hDB, pzName, eTextRep, pArg, lpfnCompare, lpfnDestroy)
End Function

Public Function stub_sqlite3_create_collation16(ByVal hDB As Long, ByVal pzName As Long, ByVal eTextRep As Long, ByVal pArg As Long, ByVal lpfnCompare As Long) As Long
stub_sqlite3_create_collation16 = sqlite3_create_collation16(hDB, pzName, eTextRep, pArg, lpfnCompare)
End Function

Public Function stub_sqlite3_create_function(ByVal hDB As Long, ByVal pzFunc As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long) As Long
stub_sqlite3_create_function = sqlite3_create_function(hDB, pzFunc, nArg, eTextRep, pApp, lpfnFunc, lpfnStep, lpfnFinal)
End Function

Public Function stub_sqlite3_create_function_v2(ByVal hDB As Long, ByVal pzFunc As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_create_function_v2 = sqlite3_create_function_v2(hDB, pzFunc, nArg, eTextRep, pApp, lpfnFunc, lpfnStep, lpfnFinal, lpfnDestroy)
End Function

Public Function stub_sqlite3_create_function16(ByVal hDB As Long, ByVal pzFunctionName As Long, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As Long, ByVal lpfnFunc As Long, ByVal lpfnStep As Long, ByVal lpfnFinal As Long) As Long
stub_sqlite3_create_function16 = sqlite3_create_function16(hDB, pzFunctionName, nArg, eTextRep, pApp, lpfnFunc, lpfnStep, lpfnFinal)
End Function

Public Function stub_sqlite3_create_module(ByVal hDB As Long, ByVal pzName As Long, ByVal pModule As Long, ByVal pAux As Long) As Long
stub_sqlite3_create_module = sqlite3_create_module(hDB, pzName, pModule, pAux)
End Function

Public Function stub_sqlite3_create_module_v2(ByVal hDB As Long, ByVal pzName As Long, ByVal pModule As Long, ByVal pAux As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_create_module_v2 = sqlite3_create_module_v2(hDB, pzName, pModule, pAux, lpfnDestroy)
End Function

Public Function stub_sqlite3_data_count(ByVal hStmt As Long) As Long
stub_sqlite3_data_count = sqlite3_data_count(hStmt)
End Function

Public Function stub_sqlite3_db_cacheflush(ByVal hDB As Long) As Long
stub_sqlite3_db_cacheflush = sqlite3_db_cacheflush(hDB)
End Function

Public Function stub_sqlite3_db_filename(ByVal hDB As Long, ByVal pzDbName As Long) As Long
stub_sqlite3_db_filename = sqlite3_db_filename(hDB, pzDbName)
End Function

Public Function stub_sqlite3_db_handle(ByVal hStmt As Long) As Long
stub_sqlite3_db_handle = sqlite3_db_handle(hStmt)
End Function

Public Function stub_sqlite3_db_mutex(ByVal hDB As Long) As Long
stub_sqlite3_db_mutex = sqlite3_db_mutex(hDB)
End Function

Public Function stub_sqlite3_db_readonly(ByVal hDB As Long, ByVal pzDbName As Long) As Long
stub_sqlite3_db_readonly = sqlite3_db_readonly(hDB, pzDbName)
End Function

Public Function stub_sqlite3_db_release_memory(ByVal hDB As Long) As Long
stub_sqlite3_db_release_memory = sqlite3_db_release_memory(hDB)
End Function

Public Function stub_sqlite3_db_status(ByVal hDB As Long, ByVal StatusOpt As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
stub_sqlite3_db_status = sqlite3_db_status(hDB, StatusOpt, pCurrent, pHighwater, ResetFlag)
End Function

Public Function stub_sqlite3_declare_vtab(ByVal hDB As Long, ByVal pzCreateTable As Long) As Long
stub_sqlite3_declare_vtab = sqlite3_declare_vtab(hDB, pzCreateTable)
End Function

Public Function stub_sqlite3_enable_load_extension(ByVal hDB As Long, ByVal iOnOff As Long) As Long
stub_sqlite3_enable_load_extension = sqlite3_enable_load_extension(hDB, iOnOff)
End Function

Public Function stub_sqlite3_enable_shared_cache(ByVal fEnable As Long) As Long
stub_sqlite3_enable_shared_cache = sqlite3_enable_shared_cache(fEnable)
End Function

Public Function stub_sqlite3_errcode(ByVal hDB As Long) As Long
stub_sqlite3_errcode = sqlite3_errcode(hDB)
End Function

Public Function stub_sqlite3_errmsg(ByVal hDB As Long) As Long
stub_sqlite3_errmsg = sqlite3_errmsg(hDB)
End Function

Public Function stub_sqlite3_errmsg16(ByVal hDB As Long) As Long
stub_sqlite3_errmsg16 = sqlite3_errmsg16(hDB)
End Function

Public Function stub_sqlite3_errstr(ByVal ResultCode As Long) As Long
stub_sqlite3_errstr = sqlite3_errstr(ResultCode)
End Function

Public Function stub_sqlite3_exec(ByVal hDB As Long, ByVal pzSQL As Long, ByVal lpfnCallback As Long, ByVal pArg As Long, ByVal pzErrMsg As Long) As Long
stub_sqlite3_exec = sqlite3_exec(hDB, pzSQL, lpfnCallback, pArg, pzErrMsg)
End Function

Public Function stub_sqlite3_expanded_sql(ByVal hStmt As Long) As Long
stub_sqlite3_expanded_sql = sqlite3_expanded_sql(hStmt)
End Function

Public Function stub_sqlite3_expired(ByVal hStmt As Long) As Long
stub_sqlite3_expired = sqlite3_expired(hStmt)
End Function

Public Function stub_sqlite3_extended_errcode(ByVal hDB As Long) As Long
stub_sqlite3_extended_errcode = sqlite3_extended_errcode(hDB)
End Function

Public Function stub_sqlite3_extended_result_codes(ByVal hDB As Long, ByVal iOnOff As Long) As Long
stub_sqlite3_extended_result_codes = sqlite3_extended_result_codes(hDB, iOnOff)
End Function

Public Function stub_sqlite3_file_control(ByVal hDB As Long, ByVal pzDbName As Long, ByVal Code As Long, ByVal pArg As Long) As Long
stub_sqlite3_file_control = sqlite3_file_control(hDB, pzDbName, Code, pArg)
End Function

Public Function stub_sqlite3_finalize(ByVal hStmt As Long) As Long
stub_sqlite3_finalize = sqlite3_finalize(hStmt)
End Function

Public Function stub_sqlite3_free(ByVal pMem As Long) As Long
stub_sqlite3_free = sqlite3_free(pMem)
End Function

Public Function stub_sqlite3_free_table(ByVal azResult As Long) As Long
stub_sqlite3_free_table = sqlite3_free_table(azResult)
End Function

Public Function stub_sqlite3_get_autocommit(ByVal hDB As Long) As Long
stub_sqlite3_get_autocommit = sqlite3_get_autocommit(hDB)
End Function

Public Function stub_sqlite3_get_auxdata(ByVal pCtx As Long, ByVal iArg As Long) As Long
stub_sqlite3_get_auxdata = sqlite3_get_auxdata(pCtx, iArg)
End Function

Public Function stub_sqlite3_get_table(ByVal hDB As Long, ByVal pzSQL As Long, ByVal pazResult As Long, ByVal pnRow As Long, ByVal pnColumn As Long, ByVal pzErrMsg As Long) As Long
stub_sqlite3_get_table = sqlite3_get_table(hDB, pzSQL, pazResult, pnRow, pnColumn, pzErrMsg)
End Function

Public Function stub_sqlite3_global_recover() As Long
stub_sqlite3_global_recover = sqlite3_global_recover()
End Function

Public Function stub_sqlite3_initialize() As Long
stub_sqlite3_initialize = sqlite3_initialize()
End Function

Public Function stub_sqlite3_interrupt(ByVal hDB As Long) As Long
stub_sqlite3_interrupt = sqlite3_interrupt(hDB)
End Function

Public Function stub_sqlite3_keyword_check(ByVal pzName As Long, ByVal pnName As Long) As Long
stub_sqlite3_keyword_check = sqlite3_keyword_check(pzName, pnName)
End Function

Public Function stub_sqlite3_keyword_count() As Long
stub_sqlite3_keyword_count = sqlite3_keyword_count()
End Function

Public Function stub_sqlite3_keyword_name(ByVal i As Long, ByVal pzName As Long, ByVal pnName As Long) As Long
stub_sqlite3_keyword_name = sqlite3_keyword_name(i, pzName, pnName)
End Function

Public Function stub_sqlite3_last_insert_rowid(ByVal hDB As Long) As Currency
stub_sqlite3_last_insert_rowid = sqlite3_last_insert_rowid(hDB)
End Function

Public Function stub_sqlite3_libversion() As Long
stub_sqlite3_libversion = sqlite3_libversion()
End Function

Public Function stub_sqlite3_libversion_number() As Long
stub_sqlite3_libversion_number = sqlite3_libversion_number()
End Function

Public Function stub_sqlite3_limit(ByVal hDB As Long, ByVal LimitID As Long, ByVal NewLimit As Long) As Long
stub_sqlite3_limit = sqlite3_limit(hDB, LimitID, NewLimit)
End Function

Public Function stub_sqlite3_load_extension(ByVal hDB As Long, ByVal pzFile As Long, ByVal pzProc As Long, ByVal pzErrMsg As Long) As Long
stub_sqlite3_load_extension = sqlite3_load_extension(hDB, pzFile, pzProc, pzErrMsg)
End Function

Public Function stub_sqlite3_malloc(ByVal n As Long) As Long
stub_sqlite3_malloc = sqlite3_malloc(n)
End Function

Public Function stub_sqlite3_malloc64(ByVal n As Currency) As Long
stub_sqlite3_malloc64 = sqlite3_malloc64(n)
End Function

Public Function stub_sqlite3_memory_alarm(ByVal lpfnCallback As Long, ByVal pArg As Long, ByVal iThreshold As Currency) As Long
stub_sqlite3_memory_alarm = sqlite3_memory_alarm(lpfnCallback, pArg, iThreshold)
End Function

Public Function stub_sqlite3_memory_highwater(ByVal ResetFlag As Long) As Currency
stub_sqlite3_memory_highwater = sqlite3_memory_highwater(ResetFlag)
End Function

Public Function stub_sqlite3_memory_used() As Currency
stub_sqlite3_memory_used = sqlite3_memory_used()
End Function

Public Function stub_sqlite3_msize(ByVal pMem As Long) As Currency
stub_sqlite3_msize = sqlite3_msize(pMem)
End Function

Public Function stub_sqlite3_mutex_alloc(ByVal pMtx As Long) As Long
stub_sqlite3_mutex_alloc = sqlite3_mutex_alloc(pMtx)
End Function

Public Function stub_sqlite3_mutex_enter(ByVal pMtx As Long) As Long
stub_sqlite3_mutex_enter = sqlite3_mutex_enter(pMtx)
End Function

Public Function stub_sqlite3_mutex_free(ByVal pMtx As Long) As Long
stub_sqlite3_mutex_free = sqlite3_mutex_free(pMtx)
End Function

Public Function stub_sqlite3_mutex_leave(ByVal pMtx As Long) As Long
stub_sqlite3_mutex_leave = sqlite3_mutex_leave(pMtx)
End Function

Public Function stub_sqlite3_mutex_try(ByVal pMtx As Long) As Long
stub_sqlite3_mutex_try = sqlite3_mutex_try(pMtx)
End Function

Public Function stub_sqlite3_next_stmt(ByVal hDB As Long, ByVal hStmt As Long) As Long
stub_sqlite3_next_stmt = sqlite3_next_stmt(hDB, hStmt)
End Function

Public Function stub_sqlite3_open(ByVal pzFilename As Long, ByRef hDB As Long) As Long
stub_sqlite3_open = sqlite3_open(pzFilename, hDB)
End Function

Public Function stub_sqlite3_open_v2(ByVal pzFilename As Long, ByRef hDB As Long, ByVal Flags As Long, ByVal pzVfs As Long) As Long
stub_sqlite3_open_v2 = sqlite3_open_v2(pzFilename, hDB, Flags, pzVfs)
End Function

Public Function stub_sqlite3_open16(ByVal pzFilename As Long, ByRef hDB As Long) As Long
stub_sqlite3_open16 = sqlite3_open16(pzFilename, hDB)
End Function

Public Function stub_sqlite3_os_end() As Long
stub_sqlite3_os_end = sqlite3_os_end()
End Function

Public Function stub_sqlite3_os_init() As Long
stub_sqlite3_os_init = sqlite3_os_init()
End Function

Public Function stub_sqlite3_overload_function(ByVal hDB As Long, ByVal pzName As Long, ByVal nArg As Long) As Long
stub_sqlite3_overload_function = sqlite3_overload_function(hDB, pzName, nArg)
End Function

Public Function stub_sqlite3_prepare(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare = sqlite3_prepare(hDB, pzSQL, nByte, hStmt, pzTail)
End Function

Public Function stub_sqlite3_prepare_v2(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare_v2 = sqlite3_prepare_v2(hDB, pzSQL, nByte, hStmt, pzTail)
End Function

Public Function stub_sqlite3_prepare_v3(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByVal PrepFlags As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare_v3 = sqlite3_prepare_v3(hDB, pzSQL, nByte, PrepFlags, hStmt, pzTail)
End Function

Public Function stub_sqlite3_prepare16(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare16 = sqlite3_prepare16(hDB, pzSQL, nByte, hStmt, pzTail)
End Function

Public Function stub_sqlite3_prepare16_v2(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare16_v2 = sqlite3_prepare16_v2(hDB, pzSQL, nByte, hStmt, pzTail)
End Function

Public Function stub_sqlite3_prepare16_v3(ByVal hDB As Long, ByVal pzSQL As Long, ByVal nByte As Long, ByVal PrepFlags As Long, ByRef hStmt As Long, ByVal pzTail As Long) As Long
stub_sqlite3_prepare16_v3 = sqlite3_prepare16_v3(hDB, pzSQL, nByte, PrepFlags, hStmt, pzTail)
End Function

Public Function stub_sqlite3_profile(ByVal hDB As Long, ByVal lpfnProfile As Long, ByVal pArg As Long) As Long
stub_sqlite3_profile = sqlite3_profile(hDB, lpfnProfile, pArg)
End Function

Public Function stub_sqlite3_progress_handler(ByVal hDB As Long, ByVal nOps As Long, ByVal lpfnProgress As Long, ByVal pArg As Long) As Long
stub_sqlite3_progress_handler = sqlite3_progress_handler(hDB, nOps, lpfnProgress, pArg)
End Function

Public Function stub_sqlite3_randomness(ByVal n As Long, ByVal pBuf As Long) As Long
stub_sqlite3_randomness = sqlite3_randomness(n, pBuf)
End Function

Public Function stub_sqlite3_realloc(ByVal pOld As Long, ByVal n As Long) As Long
stub_sqlite3_realloc = sqlite3_realloc(pOld, n)
End Function

Public Function stub_sqlite3_realloc64(ByVal pOld As Long, ByVal n As Currency) As Long
stub_sqlite3_realloc64 = sqlite3_realloc64(pOld, n)
End Function

Public Function stub_sqlite3_release_memory(ByVal n As Long) As Long
stub_sqlite3_release_memory = sqlite3_release_memory(n)
End Function

Public Function stub_sqlite3_reset(ByVal hStmt As Long) As Long
stub_sqlite3_reset = sqlite3_reset(hStmt)
End Function

Public Function stub_sqlite3_reset_auto_extension() As Long
stub_sqlite3_reset_auto_extension = sqlite3_reset_auto_extension()
End Function

Public Function stub_sqlite3_result_blob(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_blob = sqlite3_result_blob(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_blob64(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Currency, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_blob64 = sqlite3_result_blob64(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_double(ByVal pCtx As Long, ByVal rValue As Double) As Long
stub_sqlite3_result_double = sqlite3_result_double(pCtx, rValue)
End Function

Public Function stub_sqlite3_result_error(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long) As Long
stub_sqlite3_result_error = sqlite3_result_error(pCtx, pz, n)
End Function

Public Function stub_sqlite3_result_error_code(ByVal pCtx As Long, ByVal ErrCode As Long) As Long
stub_sqlite3_result_error_code = sqlite3_result_error_code(pCtx, ErrCode)
End Function

Public Function stub_sqlite3_result_error_nomem(ByVal pCtx As Long) As Long
stub_sqlite3_result_error_nomem = sqlite3_result_error_nomem(pCtx)
End Function

Public Function stub_sqlite3_result_error_toobig(ByVal pCtx As Long) As Long
stub_sqlite3_result_error_toobig = sqlite3_result_error_toobig(pCtx)
End Function

Public Function stub_sqlite3_result_error16(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long) As Long
stub_sqlite3_result_error16 = sqlite3_result_error16(pCtx, pz, n)
End Function

Public Function stub_sqlite3_result_int(ByVal pCtx As Long, ByVal iValue As Long) As Long
stub_sqlite3_result_int = sqlite3_result_int(pCtx, iValue)
End Function

Public Function stub_sqlite3_result_int64(ByVal pCtx As Long, ByVal iValue As Currency) As Long
stub_sqlite3_result_int64 = sqlite3_result_int64(pCtx, iValue)
End Function

Public Function stub_sqlite3_result_null(ByVal pCtx As Long) As Long
stub_sqlite3_result_null = sqlite3_result_null(pCtx)
End Function

Public Function stub_sqlite3_result_subtype(ByVal pCtx As Long, ByVal eSubtype As Long) As Long
stub_sqlite3_result_subtype = sqlite3_result_subtype(pCtx, eSubtype)
End Function

Public Function stub_sqlite3_result_text(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_text = sqlite3_result_text(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_text16(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_text16 = sqlite3_result_text16(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_text16be(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_text16be = sqlite3_result_text16be(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_text16le(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_text16le = sqlite3_result_text16le(pCtx, pz, n, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_text64(ByVal pCtx As Long, ByVal pz As Long, ByVal n As Currency, ByVal lpfnDestroy As Long, ByVal Encoding As Byte) As Long
stub_sqlite3_result_text64 = sqlite3_result_text64(pCtx, pz, n, lpfnDestroy, Encoding)
End Function

Public Function stub_sqlite3_result_value(ByVal pCtx As Long, ByVal pValue As Long) As Long
stub_sqlite3_result_value = sqlite3_result_value(pCtx, pValue)
End Function

Public Function stub_sqlite3_result_pointer(ByVal pCtx As Long, ByVal pPtr As Long, ByVal pzPType As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_result_pointer = sqlite3_result_pointer(pCtx, pPtr, pzPType, lpfnDestroy)
End Function

Public Function stub_sqlite3_result_zeroblob(ByVal pCtx As Long, ByVal n As Long) As Long
stub_sqlite3_result_zeroblob = sqlite3_result_zeroblob(pCtx, n)
End Function

Public Function stub_sqlite3_result_zeroblob64(ByVal pCtx As Long, ByVal n As Currency) As Long
stub_sqlite3_result_zeroblob64 = sqlite3_result_zeroblob64(pCtx, n)
End Function

Public Function stub_sqlite3_rollback_hook(ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
stub_sqlite3_rollback_hook = sqlite3_rollback_hook(hDB, lpfnCallback, pArg)
End Function

Public Function stub_sqlite3_rtree_geometry_callback(ByVal hDB As Long, ByVal pzGeom As Long, ByVal lpfnGeom As Long, ByVal pContext As Long) As Long
stub_sqlite3_rtree_geometry_callback = sqlite3_rtree_geometry_callback(hDB, pzGeom, lpfnGeom, pContext)
End Function

Public Function stub_sqlite3_rtree_query_callback(ByVal hDB As Long, ByVal pzQueryFunc As Long, ByVal lpfnQueryFunc As Long, ByVal pContext As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_rtree_query_callback = sqlite3_rtree_query_callback(hDB, pzQueryFunc, lpfnQueryFunc, pContext, lpfnDestroy)
End Function

Public Function stub_sqlite3_set_authorizer(ByVal hDB As Long, ByVal xAuth As Long, ByVal pArg As Long) As Long
stub_sqlite3_set_authorizer = sqlite3_set_authorizer(hDB, xAuth, pArg)
End Function

Public Function stub_sqlite3_set_auxdata(ByVal pCtx As Long, ByVal iArg As Long, ByVal pAux As Long, ByVal lpfnDestroy As Long) As Long
stub_sqlite3_set_auxdata = sqlite3_set_auxdata(pCtx, iArg, pAux, lpfnDestroy)
End Function

Public Function stub_sqlite3_set_last_insert_rowid(ByVal hDB As Long, ByVal iRow As Currency) As Long
stub_sqlite3_set_last_insert_rowid = sqlite3_set_last_insert_rowid(hDB, iRow)
End Function

Public Function stub_sqlite3_shutdown() As Long
stub_sqlite3_shutdown = sqlite3_shutdown()
End Function

Public Function stub_sqlite3_sleep(ByVal dwMilliseconds As Long) As Long
stub_sqlite3_sleep = sqlite3_sleep(dwMilliseconds)
End Function

Public Function stub_sqlite3_soft_heap_limit(ByVal n As Long) As Long
stub_sqlite3_soft_heap_limit = sqlite3_soft_heap_limit(n)
End Function

Public Function stub_sqlite3_soft_heap_limit64(ByVal n As Currency) As Currency
stub_sqlite3_soft_heap_limit64 = sqlite3_soft_heap_limit64(n)
End Function

Public Function stub_sqlite3_sourceid() As Long
stub_sqlite3_sourceid = sqlite3_sourceid()
End Function

Public Function stub_sqlite3_sql(ByVal hStmt As Long) As Long
stub_sqlite3_sql = sqlite3_sql(hStmt)
End Function

Public Function stub_sqlite3_status(ByVal Code As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
stub_sqlite3_status = sqlite3_status(Code, pCurrent, pHighwater, ResetFlag)
End Function

Public Function stub_sqlite3_status64(ByVal Code As Long, ByVal pCurrent As Long, ByVal pHighwater As Long, ByVal ResetFlag As Long) As Long
stub_sqlite3_status64 = sqlite3_status64(Code, pCurrent, pHighwater, ResetFlag)
End Function

Public Function stub_sqlite3_step(ByVal hStmt As Long) As Long
stub_sqlite3_step = sqlite3_step(hStmt)
End Function

Public Function stub_sqlite3_stmt_busy(ByVal hStmt As Long) As Long
stub_sqlite3_stmt_busy = sqlite3_stmt_busy(hStmt)
End Function

Public Function stub_sqlite3_stmt_readonly(ByVal hStmt As Long) As Long
stub_sqlite3_stmt_readonly = sqlite3_stmt_readonly(hStmt)
End Function

Public Function stub_sqlite3_stmt_status(ByVal hStmt As Long, ByVal Code As Long, ByVal ResetFlag As Long) As Long
stub_sqlite3_stmt_status = sqlite3_stmt_status(hStmt, Code, ResetFlag)
End Function

Public Function stub_sqlite3_strglob(ByVal pzGlobPattern As Long, ByVal pzString As Long) As Long
stub_sqlite3_strglob = sqlite3_strglob(pzGlobPattern, pzString)
End Function

Public Function stub_sqlite3_stricmp(ByVal pzLeft As Long, ByVal pzRight As Long) As Long
stub_sqlite3_stricmp = sqlite3_stricmp(pzLeft, pzRight)
End Function

Public Function stub_sqlite3_strlike(ByVal pzPattern As Long, ByVal pzStr As Long, ByVal cEsc As Long) As Long
stub_sqlite3_strlike = sqlite3_strlike(pzPattern, pzStr, cEsc)
End Function

Public Function stub_sqlite3_strnicmp(ByVal pzLeft As Long, ByVal pzRight As Long, ByVal n As Long) As Long
stub_sqlite3_strnicmp = sqlite3_strnicmp(pzLeft, pzRight, n)
End Function

Public Function stub_sqlite3_system_errno() As Long
stub_sqlite3_system_errno = sqlite3_system_errno()
End Function

Public Function stub_sqlite3_table_column_metadata(ByVal hDB As Long, ByVal pzDbName As Long, ByVal pzTableName As Long, ByVal pzColumnName As Long, ByVal pzDataType As Long, ByVal pzCollSeq As Long, ByVal pNotNull As Long, ByVal pPrimaryKey As Long, ByVal pAutoinc As Long) As Long
stub_sqlite3_table_column_metadata = sqlite3_table_column_metadata(hDB, pzDbName, pzTableName, pzColumnName, pzDataType, pzCollSeq, pNotNull, pPrimaryKey, pAutoinc)
End Function

Public Function stub_sqlite3_thread_cleanup() As Long
stub_sqlite3_thread_cleanup = sqlite3_thread_cleanup()
End Function

Public Function stub_sqlite3_threadsafe() As Long
stub_sqlite3_threadsafe = sqlite3_threadsafe()
End Function

Public Function stub_sqlite3_total_changes(ByVal hDB As Long) As Long
stub_sqlite3_total_changes = sqlite3_total_changes(hDB)
End Function

Public Function stub_sqlite3_trace(ByVal hDB As Long, ByVal lpfnTrace As Long, ByVal pArg As Long) As Long
stub_sqlite3_trace = sqlite3_trace(hDB, lpfnTrace, pArg)
End Function

Public Function stub_sqlite3_trace_v2(ByVal hDB As Long, ByVal uMask As Long, ByVal lpfnCallback As Long, ByVal pCtx As Long) As Long
stub_sqlite3_trace_v2 = sqlite3_trace_v2(hDB, uMask, lpfnCallback, pCtx)
End Function

Public Function stub_sqlite3_transfer_bindings(ByVal hStmtFrom As Long, ByVal hStmtTo As Long) As Long
stub_sqlite3_transfer_bindings = sqlite3_transfer_bindings(hStmtFrom, hStmtTo)
End Function

Public Function stub_sqlite3_update_hook(ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
stub_sqlite3_update_hook = sqlite3_update_hook(hDB, lpfnCallback, pArg)
End Function

Public Function stub_sqlite3_uri_boolean(ByVal pzFilename As Long, ByVal pzParam As Long, ByVal bDefault As Long) As Long
stub_sqlite3_uri_boolean = sqlite3_uri_boolean(pzFilename, pzParam, bDefault)
End Function

Public Function stub_sqlite3_uri_int64(ByVal pzFilename As Long, ByVal pzParam As Long, ByVal bDefault As Currency) As Currency
stub_sqlite3_uri_int64 = sqlite3_uri_int64(pzFilename, pzParam, bDefault)
End Function

Public Function stub_sqlite3_uri_parameter(ByVal pzFilename As Long, ByVal pzParam As Long) As Long
stub_sqlite3_uri_parameter = sqlite3_uri_parameter(pzFilename, pzParam)
End Function

Public Function stub_sqlite3_user_data(ByVal pCtx As Long) As Long
stub_sqlite3_user_data = sqlite3_user_data(pCtx)
End Function

Public Function stub_sqlite3_value_blob(ByVal pValue As Long) As Long
stub_sqlite3_value_blob = sqlite3_value_blob(pValue)
End Function

Public Function stub_sqlite3_value_bytes(ByVal pValue As Long) As Long
stub_sqlite3_value_bytes = sqlite3_value_bytes(pValue)
End Function

Public Function stub_sqlite3_value_bytes16(ByVal pValue As Long) As Long
stub_sqlite3_value_bytes16 = sqlite3_value_bytes16(pValue)
End Function

Public Function stub_sqlite3_value_double(ByVal pValue As Long) As Double
stub_sqlite3_value_double = sqlite3_value_double(pValue)
End Function

Public Function stub_sqlite3_value_dup(ByVal pOrig As Long) As Long
stub_sqlite3_value_dup = sqlite3_value_dup(pOrig)
End Function

Public Function stub_sqlite3_value_free(ByVal pOld As Long) As Long
stub_sqlite3_value_free = sqlite3_value_free(pOld)
End Function

Public Function stub_sqlite3_value_int(ByVal pValue As Long) As Long
stub_sqlite3_value_int = sqlite3_value_int(pValue)
End Function

Public Function stub_sqlite3_value_int64(ByVal pValue As Long) As Currency
stub_sqlite3_value_int64 = sqlite3_value_int64(pValue)
End Function

Public Function stub_sqlite3_value_pointer(ByVal pValue As Long, ByVal pzPType As Long) As Long
stub_sqlite3_value_pointer = sqlite3_value_pointer(pValue, pzPType)
End Function

Public Function stub_sqlite3_value_numeric_type(ByVal pValue As Long) As Long
stub_sqlite3_value_numeric_type = sqlite3_value_numeric_type(pValue)
End Function

Public Function stub_sqlite3_value_nochange(ByVal pValue As Long) As Long
stub_sqlite3_value_nochange = sqlite3_value_nochange(pValue)
End Function

Public Function stub_sqlite3_value_subtype(ByVal pValue As Long) As Long
stub_sqlite3_value_subtype = sqlite3_value_subtype(pValue)
End Function

Public Function stub_sqlite3_value_text(ByVal pValue As Long) As Long
stub_sqlite3_value_text = sqlite3_value_text(pValue)
End Function

Public Function stub_sqlite3_value_text16(ByVal pValue As Long) As Long
stub_sqlite3_value_text16 = sqlite3_value_text16(pValue)
End Function

Public Function stub_sqlite3_value_text16be(ByVal pValue As Long) As Long
stub_sqlite3_value_text16be = sqlite3_value_text16be(pValue)
End Function

Public Function stub_sqlite3_value_text16le(ByVal pValue As Long) As Long
stub_sqlite3_value_text16le = sqlite3_value_text16le(pValue)
End Function

Public Function stub_sqlite3_value_type(ByVal pValue As Long) As Long
stub_sqlite3_value_type = sqlite3_value_type(pValue)
End Function

Public Function stub_sqlite3_vfs_find(ByVal pzVfs As Long) As Long
stub_sqlite3_vfs_find = sqlite3_vfs_find(pzVfs)
End Function

Public Function stub_sqlite3_vfs_register(ByVal pVfs As Long, ByVal MakeDefault As Long) As Long
stub_sqlite3_vfs_register = sqlite3_vfs_register(pVfs, MakeDefault)
End Function

Public Function stub_sqlite3_vfs_unregister(ByVal pVfs As Long) As Long
stub_sqlite3_vfs_unregister = sqlite3_vfs_unregister(pVfs)
End Function

Public Function stub_sqlite3_vmprintf(ByVal pzFormat As Long, ByVal va_list As Long) As Long
stub_sqlite3_vmprintf = sqlite3_vmprintf(pzFormat, va_list)
End Function

Public Function stub_sqlite3_vsnprintf(ByVal n As Long, ByVal pzBuffer As Long, ByVal pzFormat As Long, ByVal va_list As Long) As Long
stub_sqlite3_vsnprintf = sqlite3_vsnprintf(n, pzBuffer, pzFormat, va_list)
End Function

Public Function stub_sqlite3_vtab_collation(ByVal pIdxInfo As Long, ByVal iCons As Long) As Long
stub_sqlite3_vtab_collation = sqlite3_vtab_collation(pIdxInfo, iCons)
End Function

Public Function stub_sqlite3_vtab_nochange(ByVal pCtx As Long) As Long
stub_sqlite3_vtab_nochange = sqlite3_vtab_nochange(pCtx)
End Function

Public Function stub_sqlite3_vtab_on_conflict(ByVal hDB As Long) As Long
stub_sqlite3_vtab_on_conflict = sqlite3_vtab_on_conflict(hDB)
End Function

Public Function stub_sqlite3_wal_autocheckpoint(ByVal hDB As Long, ByVal nFrame As Long) As Long
stub_sqlite3_wal_autocheckpoint = sqlite3_wal_autocheckpoint(hDB, nFrame)
End Function

Public Function stub_sqlite3_wal_checkpoint(ByVal hDB As Long, ByVal pzDB As Long) As Long
stub_sqlite3_wal_checkpoint = sqlite3_wal_checkpoint(hDB, pzDB)
End Function

Public Function stub_sqlite3_wal_checkpoint_v2(ByVal hDB As Long, ByVal pzDB As Long, ByVal eMode As Long, ByVal pnLog As Long, ByVal pnCkpt As Long) As Long
stub_sqlite3_wal_checkpoint_v2 = sqlite3_wal_checkpoint_v2(hDB, pzDB, eMode, pnLog, pnCkpt)
End Function

Public Function stub_sqlite3_wal_hook(ByVal hDB As Long, ByVal lpfnCallback As Long, ByVal pArg As Long) As Long
stub_sqlite3_wal_hook = sqlite3_wal_hook(hDB, lpfnCallback, pArg)
End Function

Public Function stub_sqlite3_win32_is_nt() As Long
stub_sqlite3_win32_is_nt = sqlite3_win32_is_nt()
End Function

Public Function stub_sqlite3_win32_mbcs_to_utf8(ByVal pzFilename As Long) As Long
stub_sqlite3_win32_mbcs_to_utf8 = sqlite3_win32_mbcs_to_utf8(pzFilename)
End Function

Public Function stub_sqlite3_win32_set_directory(ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
stub_sqlite3_win32_set_directory = sqlite3_win32_set_directory(DirectoryType, pzValue)
End Function

Public Function stub_sqlite3_win32_set_directory8(ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
stub_sqlite3_win32_set_directory8 = sqlite3_win32_set_directory8(DirectoryType, pzValue)
End Function

Public Function stub_sqlite3_win32_set_directory16(ByVal DirectoryType As Long, ByVal pzValue As Long) As Long
stub_sqlite3_win32_set_directory16 = sqlite3_win32_set_directory16(DirectoryType, pzValue)
End Function

Public Function stub_sqlite3_win32_sleep(ByVal dwMilliseconds As Long) As Long
stub_sqlite3_win32_sleep = sqlite3_win32_sleep(dwMilliseconds)
End Function

Public Function stub_sqlite3_win32_utf8_to_mbcs(ByVal pzFilename As Long) As Long
stub_sqlite3_win32_utf8_to_mbcs = sqlite3_win32_utf8_to_mbcs(pzFilename)
End Function

Public Function stub_sqlite3_win32_write_debug(ByVal pzBuffer As Long, ByVal nBuffer As Long) As Long
stub_sqlite3_win32_write_debug = sqlite3_win32_write_debug(pzBuffer, nBuffer)
End Function
