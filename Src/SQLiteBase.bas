Attribute VB_Name = "SQLiteBase"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 As Long = 65001
Private Const JULIANDAY_OFFSET As Double = 2415018.5
Private SQLiteRefCount As Long

Public Sub SQLiteAddRef()
' It is recommended that applications always invoke sqlite3_initialize() directly prior to using any other functions.
' Future releases of SQLite may require this. In other words, the behavior exhibited when SQLite is compiled with SQLITE_OMIT_AUTOINIT might become the default behavior in some future release of SQLite.
If SQLiteRefCount = 0 Then stub_sqlite3_initialize
SQLiteRefCount = SQLiteRefCount + 1
End Sub

Public Sub SQLiteRelease()
SQLiteRefCount = SQLiteRefCount - 1
If SQLiteRefCount = 0 Then stub_sqlite3_shutdown
End Sub

Public Function SQLiteUTF8PtrToStr(ByVal Ptr As Long) As String
If Ptr <> 0 Then
    Dim Size As Long, Length As Long
    Size = lstrlenA(Ptr)
    Length = MultiByteToWideChar(CP_UTF8, 0, Ptr, Size, 0, 0)
    If Length > 0 Then
        SQLiteUTF8PtrToStr = Space$(Length)
        MultiByteToWideChar CP_UTF8, 0, Ptr, Size, StrPtr(SQLiteUTF8PtrToStr), Length
    End If
End If
End Function

Public Function SQLiteUTF16PtrToStr(ByVal Ptr As Long) As String
If Ptr <> 0 Then
    Dim Length As Long
    Length = lstrlen(Ptr)
    If Length > 0 Then
        SQLiteUTF16PtrToStr = Space$(Length)
        CopyMemory ByVal StrPtr(SQLiteUTF16PtrToStr), ByVal Ptr, Length * 2
    End If
End If
End Function

Public Function CDateToJulianDay(ByVal DateValue As Date) As Double
CDateToJulianDay = CDbl(DateValue) + JULIANDAY_OFFSET
End Function

Public Function CJulianDayToDate(ByVal JulianDay As Double) As Date
CJulianDayToDate = CDate(JulianDay - JULIANDAY_OFFSET)
End Function
