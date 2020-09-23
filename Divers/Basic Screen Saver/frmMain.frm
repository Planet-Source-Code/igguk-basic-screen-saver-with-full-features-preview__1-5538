VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   1695
   ClientTop       =   2025
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3975
      Top             =   75
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' First, all the Win32 video declares and whatnot.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean
Private Sub Animate()
    'Type in here your code
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If RunMode = rmScreenSaver Then
        Unload Me
        End
    End If
    
End Sub
Private Sub Form_Load()
Dim I As Integer
Dim hGLRC As Long

    If Mid$(Command, 1, 2) <> "/p" Then
        LockOn Me
    End If
    
    'Type your initialization code here
    
    
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If RunMode = rmScreenSaver Then
        Unload Me
        End
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Put the focus away from the screensaver
    If RunMode = rmScreenSaver Then
        LockOff Me
    End If

End Sub
Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static Count As Integer
    Count = Count + 1 ' Give enough time for program to run
    
    If Count > 5 Then
        If RunMode = rmScreenSaver Then
            Unload Me
            End
        End If
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = False

    'If Windows is shut down close this application too
    If UnloadMode = vbAppWindows Then
        Exit Sub
    End If
    
    'if a password is beeing used ask for it and check its validity
    If RunMode = rmScreenSaver And UsePassword Then
        ShowCursor True
        If (VerifyScreenSavePwd(Me.hwnd)) = False Then
            Cancel = True
        End If
        ShowCursor False
    End If

End Sub

Private Sub tmrAnimate_Timer()
    Animate
End Sub
