Attribute VB_Name = "mdlReadWriteINI"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long
Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strSection, strKey, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = vbNullString
    End If
End Function

Function WriteIniFile(strSection As String, strKey As String, strValue As String, strFileName As String)
On Error Resume Next
Call WritePrivateProfileString(strSection, strKey, strValue, strFileName)
End Function
