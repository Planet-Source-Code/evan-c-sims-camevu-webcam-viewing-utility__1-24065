VERSION 5.00
Begin VB.Form frmAddModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Cam"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
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
   Icon            =   "frmAddModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Refresh Rate"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAddModify.frx":000C
      Left            =   1200
      List            =   "frmAddModify.frx":0022
      TabIndex        =   4
      Text            =   "60"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   15
      Left            =   960
      Top             =   2410
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   960
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   1980
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cam Location:"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAddModify.frx":003E
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cam Title:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmAddModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ModifyMode As Boolean

Public camTitle As String
Public camLocal As String
Public camRate As String

Dim m_cIni As New cInifile
Private Sub Check1_Click()

If Check1.Value = 1 Then
   Combo1.Enabled = True
   Label4.ForeColor = &H80000012
Else
   Combo1.Enabled = False
   Label4.ForeColor = &H80000011
End If

End Sub


Private Sub Combo1_GotFocus()

Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1)

End Sub


Private Sub Command1_Click()

If Combo1.Text > 3600 Or Combo1.Text < 1 Then
   MsgBox "Refresh Rate must be an integer between 1 and 3600."
   Exit Sub
End If

With m_cIni
.Path = App.Path & "\favorites.dat"

If ModifyMode = True Then

        .Section = Text1.Text
        .Key = "Address"
        .Default = "OK"
        
        If .Value = "OK" Then

            If camTitle = Text1.Text Then

               .Section = camTitle
               .Key = "Address"
               .Value = Text2.Text
               
               .Key = "UpdateInterval"
               If Check1.Value = 1 Then
                  .Value = Combo1.Text
               Else
                  .Value = "x"
               End If
               
               'MsgBox "Cam Modified!"

            Else
            
               .Section = camTitle
               .DeleteSection
               
               .Section = Text1.Text
               .Key = "Address"
               .Value = Text2.Text
               
               .Key = "UpdateInterval"
               If Check1.Value = 1 Then
                  .Value = Combo1.Text
               Else
                  .Value = "x"
               End If
            
            End If

        Else
            If Not camTitle = Text1.Text Then
                Dim msgReply
                msgReply = MsgBox("A cam by that name already exists." & vbCrLf & "Would you like to replace it?", vbYesNo)
            
                If msgReply = vbYes Then
                   .Section = camTitle
                   .Key = "Address"
                   .Value = Text2.Text
                   
                   .Key = "UpdateInterval"
                   If Check1.Value = 1 Then
                      .Value = Combo1.Text
                   Else
                      .Value = "x"
                   End If
                   
                   'MsgBox "Cam Replaced!"
                Else
                   Exit Sub
                   'MsgBox "Aborted!"
                End If
            Else
                .Section = camTitle
                .Key = "Address"
                .Value = Text2.Text
                
                .Key = "UpdateInterval"
                If Check1.Value = 1 Then
                   .Value = Combo1.Text
                Else
                   .Value = "x"
                End If
            End If
        End If

Else
    .Section = Text1.Text
    .Key = "Address"
    .Value = Text2.Text

    .Key = "UpdateInterval"
    If Check1.Value = 1 Then
       .Value = Combo1.Text
    Else
       .Value = "x"
    End If
    
    'MsgBox "Cam Added!"
End If

End With

Unload Me
frmMain.ReloadCamList
frmMain.ReloadAddressCombo

End Sub


Private Sub Command2_Click()

Unload Me

End Sub


Private Sub Form_Load()

If UserPref.KeepOnTop = True Then
   SetTopmost Me, True
Else
   SetTopmost Me, False
End If

AddOfficeBorder (Command1.hWnd)
AddOfficeBorder (Command2.hWnd)

If ModifyMode = True Then
   Me.Caption = "Modify Cam"
   Command1.Caption = "&Apply"
   Command2.Caption = "&Close"
   
   Text1.Text = camTitle
   Text2.Text = camLocal
   
   If Not camRate = "x" Then
      Check1.Value = 1
      Combo1.Text = camRate
      Check1_Click
   End If
Else
   Me.Caption = "Add Cam"
   Command1.Caption = "&Add"
   Command2.Caption = "&Cancel"
   
   Text1.Text = "Untitled Cam"
   Text2.Text = "http://"
   
   Check1.Value = 1
   Combo1.Text = "60"
   Check1_Click
End If

End Sub


Private Sub Text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1)

End Sub


Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2)

End Sub


