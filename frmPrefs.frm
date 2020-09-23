VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CamEVU Preferences"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
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
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   2040
      TabIndex        =   1
      Top             =   -240
      Width           =   5055
      Begin VB.Frame framePrefs 
         Height          =   3855
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox Check14 
            Caption         =   "Low resource mode"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   22
            ToolTipText     =   "Reduces the system resources CamEVU uses by disabling certain features."
            Top             =   3360
            Width           =   1695
         End
         Begin VB.CheckBox Check12 
            Caption         =   "When a download error is encountered, disable refresh for that cam."
            Enabled         =   0   'False
            Height          =   555
            Left            =   360
            TabIndex        =   20
            ToolTipText     =   "When a download error is encountered, that cam will not automatically refresh unless you refresh it manually."
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Keep Windows On Top"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            ToolTipText     =   "When enabled CamEVU dialogs will remain in front of all other programs."
            Top             =   2640
            Width           =   1935
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Lock lists during downloads"
            Height          =   195
            Left            =   360
            TabIndex        =   18
            ToolTipText     =   "Recommended. When enabled, you will not be able to change cams via either the address bar or the cam list during a download."
            Top             =   3000
            Width           =   2295
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Import Cam DB..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   17
            Top             =   3240
            Width           =   1695
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Force a 60-second refresh on cams without a rate set."
            Height          =   315
            Left            =   360
            TabIndex        =   16
            ToolTipText     =   "Toggle this to refresh cams that do not have a refresh rate already set."
            Top             =   2040
            Width           =   2655
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Do not display download errors."
            Height          =   195
            Left            =   360
            TabIndex        =   15
            ToolTipText     =   "Turns off all download/winsock-related error messages. NOT RECOMMENDED."
            Top             =   960
            Width           =   2655
         End
         Begin VB.CheckBox Check7 
            Caption         =   "After successful download, delete the temporary cam file."
            Height          =   435
            Left            =   360
            TabIndex        =   14
            ToolTipText     =   "Deletes the temeporary files CamEVU creates when it downloads a cam graphic."
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame framePrefs 
         Height          =   3855
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox Check6 
            Caption         =   "Use Hovers in Lists"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            ToolTipText     =   "Toggles hover effects for list views."
            Top             =   480
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin VB.Frame framePrefs 
         Height          =   3855
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   4575
         Begin VB.CheckBox Check13 
            Caption         =   "Only redownload cams if they've been changed."
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   2880
            Width           =   3855
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Show In System Tray"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            ToolTipText     =   "When checked, CamEVU will sit in the system tray instead of the task bar."
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Show Statusbar"
            Height          =   195
            Left            =   360
            TabIndex        =   7
            ToolTipText     =   $"frmPrefs.frx":000C
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show Address Bar"
            Height          =   195
            Left            =   360
            TabIndex        =   6
            ToolTipText     =   "Toogles the Address bar in the main window. The address bar can be used to type web address into for easy access to the Internet."
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Look For New Versions"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   5
            Top             =   840
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Load Window Positions"
            Height          =   195
            Left            =   360
            TabIndex        =   4
            ToolTipText     =   "When checked, CamEVU will rmember the size and position it's window was set at and return to that state the next time it's run."
            Top             =   480
            Value           =   1  'Checked
            Width           =   1935
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General"
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
         Left            =   360
         TabIndex        =   2
         Top             =   510
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         Top             =   480
         Width           =   4575
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GoLicense As Boolean
Dim IsDirty As Boolean

Private Sub Check1_Click()

IsDirty = True

End Sub


Private Sub Check10_Click()

IsDirty = True

End Sub

Private Sub Check11_Click()

IsDirty = True

End Sub


Private Sub Check2_Click()

IsDirty = True

End Sub


Private Sub Check3_Click()

IsDirty = True

End Sub


Private Sub Check4_Click()

IsDirty = True

End Sub


Private Sub Check5_Click()

IsDirty = True

End Sub


Private Sub Check6_Click()

IsDirty = True

End Sub


Private Sub Check7_Click()

IsDirty = True

End Sub


Private Sub Check8_Click()

IsDirty = True

End Sub


Private Sub Check9_Click()

IsDirty = True

End Sub


Private Sub Command1_Click()

Call ApplyPreferences
IsDirty = False

End Sub


Private Sub Command2_Click()

If IsDirty = True Then
   Dim msgReply
   msgReply = MsgBox("Do you wish to save the changes you have made to your preferences?", vbQuestion Or vbYesNoCancel)
   
   If msgReply = vbYes Then
      Call ApplyPreferences
   ElseIf msgReply = vbCancel Then
      Exit Sub
   End If
End If

Unload Me

End Sub

Private Sub Command3_Click()

MsgBox "Warning: Importing a CamEVU database will overwrite your current list of cams. However, a backup of your current list is made before it is overwritten to the file ""favorites.bak"". You can reimport that file to retrieve your old list.", vbCritical

End Sub

Private Sub Form_Load()

If UserPref.KeepOnTop = True Then
   SetTopmost Me, True
Else
   SetTopmost Me, False
End If

AddOfficeBorder (Command1.hWnd)
AddOfficeBorder (Command2.hWnd)
AddOfficeBorder (Command3.hWnd)

TreeView1.Nodes.Add , tvwChild, "General", "General"
TreeView1.Nodes(1).Selected = True
TreeView1.Nodes.Add , tvwChild, "Appearance", "Appearance"
TreeView1.Nodes.Add , tvwChild, "Advanced", "Advanced"

If UserPref.UseHovers = True Then TreeView1.HotTracking = True

LoadPrefsForDialog ' Set Toggles Appropriately
IsDirty = False

End Sub

Private Sub Timer1_Timer()

If IsDirty = True Then
   Command1.Enabled = True
   Command2.Caption = "&Cancel"
Else
   Command1.Enabled = False
   Command2.Caption = "&Close"
End If

End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

Label1.Caption = Node.Text

For i = 1 To framePrefs.Count
    If framePrefs(i).Index = Node.Index Then
       framePrefs(i).Visible = True
       framePrefs(i).ZOrder 0
    Else
       framePrefs(i).Visible = False
    End If
Next

End Sub


