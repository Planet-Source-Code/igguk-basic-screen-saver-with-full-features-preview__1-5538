Attribute VB_Name = "Saver"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function PwdChangePassword& Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname$, ByVal hwnd&, ByVal uiReserved1&, ByVal uiReserved2&)
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpStruct As OSVERSIONINFO)
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatform As Long
   szCSDVersion As String * 128
End Type

Public Const APP_NAME = "Igguk - Screen Saver"
Public Const HWND_TOP = 0
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_CHILD = &H40000000
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Global Const rmConfigure = "/C"
Global Const rmScreenSaver = "/S"
Global Const rmPreview = "/P"
Global Const rmPassword = "/A"
Global RunMode As String * 2

Public Preview_Hwnd As Long
Public OsVers As OSVERSIONINFO
Public Version As String

Public Sub Main()
Dim Preview_Rect As RECT
Dim Window_Style As Long
Dim A As Long

OsVers.dwVersionInfoSize = 148&
GetVersionEx OsVers
Select Case OsVers.dwPlatform
    Case 1
        Version = "9x"
    Case 2
        Version = "NT"
    Case Else
        Version = "Unknown"
End Select

RunMode = Left(UCase(Trim(Command)) + "  ", 2)
If RunMode = "/P" Or RunMode = "/A" Then
    Preview_Hwnd = CLng(Right(Command, Len(Command) - 3))
End If

With frmMain
    Select Case RunMode
        Case "/C", "  "
            'Configure
            RunMode = rmConfigure
            frmSetup.Show
        Case "/P"
            'Create a preview screen
            GetClientRect Preview_Hwnd, Preview_Rect
            Load frmMain
            Window_Style = GetWindowLong(.hwnd, GWL_STYLE)
            Window_Style = (Window_Style Or WS_CHILD)
            SetWindowLong .hwnd, GWL_STYLE, Window_Style
            SetParent .hwnd, Preview_Hwnd
            SetWindowLong .hwnd, GWL_HWNDPARENT, Preview_Hwnd
            SetWindowPos .hwnd, HWND_TOP, 0&, 0&, Preview_Rect.Right, Preview_Rect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Case "/A"
            'Change the screensaver password
            On Error GoTo Error
            A = PwdChangePassword("SCRSAVE", Preview_Hwnd, 0, 0)
            On Error GoTo 0
        Case "/S"
            'Run the screensaver
            If App.PrevInstance And FindWindow(vbNullString, APP_NAME) Then End
            .Show
    End Select
End With
Exit Sub

Error:       MsgBox "Password could not be changed", vbOKOnly
End Sub
Public Function UsePassword() As Boolean
'Check wether a password has been used or not
Dim lHandle As Long
Dim lResult As Long
Dim lValue As Long
UsePassword = False
If Version <> "9x" Then Exit Function
lResult = RegOpenKeyEx(&H80000001, "Control Panel\Desktop", 0, 1, lHandle)
If lResult = 0 Then
    lResult = RegQueryValueEx(lHandle, "ScreenSaveUsePassword", 0, 4, lValue, 32)
    If lResult = 0 Then
        UsePassword = lValue
        lResult = RegCloseKey(lHandle)
    End If
End If
End Function
Public Sub LockOn(frmId As Form)

'If a password is beeing used in Windows 9x prevent the use of CTRL+ALT+DEL
'and so on. Make the screensaver screen the only available screen until
'a correct password has been entered
If UsePassword Then
    Call SystemParametersInfo(97, 1, 0, 0)
    SetWindowPos frmId.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End If

'Hide the mouse
ShowCursor False

End Sub
Public Sub LockOff(frmId As Form)
'Show the mouse
ShowCursor True

'Make the other screens available again
If UsePassword Then
    Call SystemParametersInfo(97, 0, 0, 0)
    SetWindowPos frmId.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End If
End Sub

