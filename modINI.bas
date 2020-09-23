Attribute VB_Name = "modINI"
Option Explicit

#If Win16 Then

Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer

Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
    ' NOTE: The lpKeyName argument for GetPr
    '     ofileString, WriteProfileString,
    'GetPrivateProfileString, and WritePriva
    '     teProfileString can be either
    'a string or NULL. This is why the argum
    '     ent is defined as "As Any".
    ' For example, to pass a string specifyB
    '     yVal "wallpaper"
    ' To pass NULL specifyByVal 0&
    'You can also pass NULL for the lpString
    '     argument for WriteProfileString
    'and WritePrivateProfileString
    ' Below it has been changed to a string
    '     due to the ability to use vbNullString


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function
