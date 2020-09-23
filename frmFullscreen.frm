VERSION 5.00
Begin VB.Form frmFullscreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CamEVU Fullscreen View"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmFullscreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3960
      Picture         =   "frmFullscreen.frx":000C
      ToolTipText     =   " Refreshing Image ... "
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Picture1 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmFullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture2_Click()

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = "27" Or KeyAscii = "32" Then
   Unload Me
End If

End Sub

Private Sub Form_Load()

Picture1.Picture = frmMain.Image2.Picture

End Sub


Private Sub Form_Resize()

If Not Picture1.Width = Me.Width Then Picture1.Width = Me.Width
If Not Picture1.Height = Me.Height Then Picture1.Height = Me.Height
If Not Image1.Top = Picture1.Height - Image1.Height - 120 Then Image1.Top = Picture1.Height - Image1.Height - 120
If Not Image1.Left = Picture1.Width - Image1.Width - 120 Then Image1.Left = Picture1.Width - Image1.Width - 120

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Picture1_Click()

   Unload Me

End Sub




