VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "Setup Window"
   ClientHeight    =   1380
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   3690
   Begin VB.CommandButton IDP_ABOUT 
      Caption         =   "&About"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "You can use this windows to allow users to customize your screen saver"
      Height          =   600
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   3300
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IDP_ABOUT_Click()

    frmAbout.Show
    
End Sub
