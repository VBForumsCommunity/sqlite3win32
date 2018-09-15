Attribute VB_Name = "Common"
Option Explicit
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const MAX_PATH As Long = 260, MAX_PATH_W As Long = 32767

Public Function AppPath() As String
If InIDE() = False Then
    Dim Buffer As String, RetVal As Long
    Buffer = String(MAX_PATH, vbNullChar)
    RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
    If RetVal = MAX_PATH Then ' Path > MAX_PATH
        Buffer = String(MAX_PATH_W, vbNullChar)
        RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
    End If
    If RetVal > 0 Then
        Buffer = Left$(Buffer, RetVal)
        AppPath = Left$(Buffer, InStrRev(Buffer, "\"))
    Else
        AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
    End If
Else
    AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")
End If
End Function

Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
If B = True Then Debug.Assert Not InIDE(InIDE) Else B = True
End Function

Public Function UTF16_To_UTF8(ByRef Source As String) As Byte()
Const CP_UTF8 As Long = 65001
Dim Length As Long, Pointer As Long, Size As Long
Length = Len(Source)
Pointer = StrPtr(Source)
Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
If Size > 0 Then
    Dim Buffer() As Byte
    ReDim Buffer(0 To Size - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), Size, 0, 0
    UTF16_To_UTF8 = Buffer()
End If
End Function

Public Function UTF8_To_UTF16(ByRef Source() As Byte) As String
If (0 / 1) + (Not Not Source()) = 0 Then Exit Function
Const CP_UTF8 As Long = 65001
Dim Size As Long, Pointer As Long, Length As Long
Size = UBound(Source) - LBound(Source) + 1
Pointer = VarPtr(Source(LBound(Source)))
Length = MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, 0, 0)
If Length > 0 Then
    UTF8_To_UTF16 = Space$(Length)
    MultiByteToWideChar CP_UTF8, 0, Pointer, Size, StrPtr(UTF8_To_UTF16), Length
End If
End Function
