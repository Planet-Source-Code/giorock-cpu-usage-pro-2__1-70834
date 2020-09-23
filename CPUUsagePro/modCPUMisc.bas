Attribute VB_Name = "modCPUMisc"
Option Explicit

'Operating system  number dwMajorVersion    dwMinorVersion                              Other

'Windows 7          6.1         6                  1            OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'Windows Server     6.1         6                  1            OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'   2008 R2
'Windows Server     6.0         6                  0            OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
'    2008
'Windows Vista      6.0         6                  0            OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
'Windows Server     5.2         5                  2            GetSystemMetrics(SM_SERVERR2) != 0
'   2003 R2
'Windows Server     5.2         5                  2            GetSystemMetrics(SM_SERVERR2) == 0
'    2003
'Windows XP         5.1         5                  1            Not applicable
'Windows 2000       5.0         5                  0            Not applicable


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = &H1

'return True is the OS is WindowsNT3.5(1), NT4.0, 2000 or XP
Public Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT, return True
    IsWinNT = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function WinNTVersion() As Long
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT and OS > XP, return version number
    WinNTVersion = CLng(OSInfo.dwMajorVersion & OSInfo.dwMinorVersion)
End Function
