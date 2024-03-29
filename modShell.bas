Attribute VB_Name = "modShell"
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib _
"kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds _
As Long) As Long

Declare Function CreateProcessA Lib "kernel32" _
(ByVal lpApplicationName As Long, ByVal lpCommandLine As _
String, ByVal lpProcessAttributes As Long, ByVal _
lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
ByVal lpCurrentDirectory As Long, lpStartupInfo As _
STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) _
As Long

Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Sub ExecCmd(cmdline$)

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    
    'Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
    'Start the shelled application:
    ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
    'Wait for the shelled application to finish:
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)
    
End Sub

