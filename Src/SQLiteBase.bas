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

Public Function SQLiteBlobToByteArray(ByVal Ptr As Long, ByVal Size As Long) As Variant
If Ptr <> 0 And Size > 0 Then
    Dim B() As Byte
    ReDim B(0 To (Size - 1)) As Byte
    CopyMemory B(0), ByVal Ptr, Size
    SQLiteBlobToByteArray = B()
Else
    SQLiteBlobToByteArray = Null
End If
End Function

Public Function SQLiteUTF8PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
If Ptr <> 0 Then
    If Size = -1 Then Size = lstrlenA(Ptr)
    If Size > 0 Then
        Dim Length As Long
        Length = MultiByteToWideChar(CP_UTF8, 0, Ptr, Size, 0, 0)
        If Length > 0 Then
            SQLiteUTF8PtrToStr = Space$(Length)
            MultiByteToWideChar CP_UTF8, 0, Ptr, Size, StrPtr(SQLiteUTF8PtrToStr), Length
        End If
    End If
End If
End Function

Public Function SQLiteUTF16PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
If Ptr <> 0 Then
    If Size = -1 Then Size = lstrlen(Ptr)
    If Size > 0 Then
        SQLiteUTF16PtrToStr = Space$(Size)
        CopyMemory ByVal StrPtr(SQLiteUTF16PtrToStr), ByVal Ptr, Size * 2
    End If
End If
End Function

Public Function CDateToJulianDay(ByVal DateValue As Date) As Double
If CDbl(DateValue) >= 0 Then
    CDateToJulianDay = CDbl(DateValue) + JULIANDAY_OFFSET
Else
    Dim Temp As Double
    Temp = -Int(-CDbl(DateValue))
    CDateToJulianDay = Temp - (CDbl(DateValue) - Temp) + JULIANDAY_OFFSET
End If
End Function

Public Function CJulianDayToDate(ByVal JulianDay As Double) As Date
Const MIN_DATE As Double = -657434# + JULIANDAY_OFFSET ' 01/01/0100
Const MAX_DATE As Double = 2958465# + JULIANDAY_OFFSET ' 12/31/9999
If JulianDay < MIN_DATE Or JulianDay > MAX_DATE Then Exit Function
If JulianDay >= JULIANDAY_OFFSET Then
    CJulianDayToDate = CDate(JulianDay - JULIANDAY_OFFSET)
Else
    Dim DateDbl As Double, Temp As Double
    DateDbl = JulianDay - JULIANDAY_OFFSET
    Temp = Int(DateDbl)
    CJulianDayToDate = CDate(Temp + (Temp - DateDbl))
End If
End Function
