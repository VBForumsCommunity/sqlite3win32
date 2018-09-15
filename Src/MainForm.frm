VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "sqlite3win32 demo"
   ClientHeight    =   1065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "MainForm"
   ScaleHeight     =   1065
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create Test DB"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MsgBox Version()"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private hLib As Long, hDB As Long

Private Sub Command1_Click()
MsgBox SQLiteUTF8PtrToStr(stub_sqlite3_libversion())
End Sub

Private Sub Command2_Click()
Call OpenDB(AppPath() & "Test.db", False)
Call Execute("DROP TABLE IF EXISTS test_table")
Call Execute("CREATE TABLE test_table(ID INT)")
Call Execute("INSERT INTO test_table (ID) VALUES (1)")
Call CloseDB
End Sub

Private Sub Form_Load()
hLib = LoadLibrary(StrPtr("sqlite3win32.dll"))
If hLib = 0 Then LoadLibrary (StrPtr(lib_dir_sqlite3win32()))
If hLib <> 0 Then Call SQLiteAddRef
End Sub

Private Sub Form_Unload(Cancel As Integer)
If hLib <> 0 Then
    FreeLibrary hLib
    hLib = 0
    Call SQLiteRelease
End If
End Sub

Private Sub Execute(ByVal Query As String)
If hDB = 0 Then Err.Raise Number:=5, Description:="DB must be opened before it can be used"
Dim QueryUTF8() As Byte, Result As Long
QueryUTF8() = UTF16_To_UTF8(Query & vbNullChar)
Result = stub_sqlite3_exec(hDB, VarPtr(QueryUTF8(0)), 0, 0, 0)
If Result <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(hDB), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(hDB))
End Sub

Private Sub OpenDB(ByVal FileName As String, Optional ByVal ReadOnly As Boolean)
If hDB <> 0 Then
    stub_sqlite3_close_v2 hDB
    hDB = 0
End If
Dim FileNameUTF8() As Byte, Flags As Long, Result As Long
FileNameUTF8() = UTF16_To_UTF8(FileName & vbNullChar)
If ReadOnly = False Then
    Flags = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
Else
    Flags = SQLITE_OPEN_READONLY
End If
Result = stub_sqlite3_open_v2(VarPtr(FileNameUTF8(0)), hDB, Flags, 0)
If Result <> SQLITE_OK Then
    Dim ErrVal As Long, ErrMsg As String
    If hDB <> 0 Then
        ErrVal = stub_sqlite3_errcode(hDB)
        ErrMsg = SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(hDB))
        stub_sqlite3_close_v2 hDB
        hDB = 0
    End If
    Err.Raise Number:=vbObjectError + ErrVal, Description:=ErrMsg
End If
End Sub

Private Sub CloseDB()
If hDB <> 0 Then
    stub_sqlite3_close_v2 hDB
    hDB = 0
End If
End Sub
