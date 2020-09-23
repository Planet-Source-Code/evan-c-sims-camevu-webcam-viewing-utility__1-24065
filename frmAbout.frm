VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CamEVU"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   3600
      ScaleHeight     =   1575
      ScaleWidth      =   3495
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Credits"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source Baby, Yeah!"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   960
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   15
      Left            =   960
      Top             =   2295
      Width           =   3495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Erroneous Data, Inc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2000-2001"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cam Easy Viewing Utility"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CamEVU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":08D6
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Picture1.Visible = Not Picture1.Visible

End Sub


Private Sub Form_Load()

If UserPref.KeepOnTop = True Then
   SetTopmost Me, True
Else
   SetTopmost Me, False
End If

Picture1.Left = 960
Picture1.Top = 600

AddOfficeBorder (Command1.hWnd)
AddOfficeBorder (Command2.hWnd)

Label5.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'Label6.Caption = "Programming by Evan Sims" & vbCrLf & vbCrLf & "Alpha Testing by Brian Herbert, Jim Weinhart, John York and Kevin Sanders."

Text1.Text = "Programming by Evan Sims." & vbCrLf & vbCrLf & _
"Special thanks to EarthCam, Square Eight and Dawn Patrol for their inspiration and support." & vbCrLf & vbCrLf & _
"Extra special thanks to Brian Herbert, Jim Weinhart, and Kevin Sanders for testing the many beta releases and providing me with much needed feedback. You guys rule."



End Sub



Private Sub Label4_Click()

Call Shell("Explorer http://www.erroneousdata.com", vbNormalFocus)

End Sub

