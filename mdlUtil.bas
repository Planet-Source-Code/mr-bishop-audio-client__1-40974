Attribute VB_Name = "mdlUtil"
Option Explicit

'**********************************************************************************
'The mciSendString function sends a command string to an MCI device.
'The device that the command is sent to is specified in the command string.
'**********************************************************************************
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'**********************************************************************************
'The mciSendCommand function sends a command message to the specified MCI device.
'**********************************************************************************
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

'**********************************************************************************
'The mciGetErrorString function retrieves a string that describes the specified MCI error code.
'**********************************************************************************
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long


Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
  ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Declare Function Process32First Lib "kernel32" ( _
  ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Declare Function Process32Next Lib "kernel32" ( _
  ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long
Declare Function EnumProcessModules Lib "PSAPI" ( _
    ByVal hProcess As Long, lphModule As Long, _
    ByVal cb As Long, lpcbNeeded As Long) As Long
Declare Function GetModuleInformation Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long

Declare Function CreateProcessBynum Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function WaitForInputIdle Lib "User32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Const hNull = &O0
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const TH32CS_SNAPMODULE = &H8&
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_VM_READ = &H10
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_NORMAL = 1

Public Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const WAIT_TIMEOUT = &H102&
Public Const CREATE_NO_WINDOW = &H8000000


Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process's threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type

Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long        ' This module
    th32ProcessID As Long       ' owning process
    GlblcntUsage As Long        ' Global usage count on the module
    ProccntUsage As Long        ' Module usage count in th32ProcessID's context
    modBaseAddr As Long         ' Base address of module in th32ProcessID's context
    modBaseSize As Long         ' Size in bytes of module starting at modBaseAddr
    hModule As Long             ' The hModule of this module in th32ProcessID's context
    szModule As String * 256    ' MAX_MODULE_NAME32 + 1
    szExePath As String * 260   ' MAX_PATH
End Type

Type STARTUPINFO
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

Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

'API for getting all drives on your PC
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Api to determine if you have a CD, floppy or Hard drive
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public doWizard As Boolean

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long
    Dim sShortPathName As String
    Dim iLen As Integer
    
    'Set up buffer area for API function cal
    '     l return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function
