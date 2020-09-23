VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "CamEVU"
   ClientHeight    =   5445
   ClientLeft      =   4425
   ClientTop       =   3555
   ClientWidth     =   7575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7575
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1349
      BandCount       =   2
      _CBWidth        =   7575
      _CBHeight       =   765
      _Version        =   "6.7.8862"
      Child1          =   "Toolbar1"
      MinWidth1       =   300
      MinHeight1      =   330
      Width1          =   795
      NewRow1         =   0   'False
      Child2          =   "Picture5"
      MinWidth2       =   495
      MinHeight2      =   345
      Width2          =   450
      NewRow2         =   -1  'True
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   165
         ScaleHeight     =   345
         ScaleWidth      =   7320
         TabIndex        =   15
         Top             =   390
         Width           =   7320
         Begin MSComctlLib.ImageCombo ImageCombo1 
            Height          =   330
            Left            =   720
            TabIndex        =   16
            Top             =   0
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "Homepage"
            ImageList       =   "ImageList1"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   0
            TabIndex        =   17
            Top             =   60
            Width           =   615
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   14
         Top             =   30
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ToolImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add ..."
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modify Cam"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Remove Cam"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy Image"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save Image As ..."
               ImageIndex      =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print Image"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Online Support"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picBrgnd 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   435
      ScaleWidth      =   150
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   2350
      ScaleHeight     =   3855
      ScaleWidth      =   105
      TabIndex        =   11
      Top             =   600
      Width           =   100
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   5880
      TabIndex        =   7
      Top             =   5235
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":959A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B41E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":104A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1680
      Top             =   2760
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1680
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3720
      ScaleHeight     =   855
      ScaleWidth      =   2655
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Release"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmMain.frx":10A42
         MousePointer    =   99  'Custom
         ToolTipText     =   " Click here to visit the CamEVU Website "
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EVU"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1320
         TabIndex        =   5
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
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
         Left            =   1350
         TabIndex        =   4
         Top             =   600
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   255
         Left            =   1275
         Top             =   570
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      Height          =   4335
      Left            =   2430
      ScaleHeight     =   4275
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   720
      Width           =   5175
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   600
         ScaleHeight     =   240
         ScaleWidth      =   2535
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Image Image3 
            Height          =   240
            Left            =   0
            Picture         =   "frmMain.frx":1130C
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Downloading Webcam ..."
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
            TabIndex        =   9
            Top             =   23
            Width           =   2040
         End
      End
      Begin VB.Image Image2 
         Height          =   3600
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   4560
      End
   End
   Begin MSComctlLib.ImageList TreeListImage 
      Left            =   1680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11365
            Key             =   "imgUser"
            Object.Tag             =   "imgUser"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11901
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12439
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":129D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15189
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E35
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ToolImageList 
      Left            =   1680
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B5E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B745
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B8A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BA05
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BB69
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BCCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BE29
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C145
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5190
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7726
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "12:32 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   7858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "TreeListImage"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add Cam"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrefs 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuFileX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "&Hide"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add Cam"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditX4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "&Refresh Cam"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditX3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "&Remove Cam"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "&Modify Cam"
      End
      Begin VB.Menu mnuEditX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy Image"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNormal 
         Caption         =   "&Normal Size"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewHalf 
         Caption         =   "&Half Size"
      End
      Begin VB.Menu mnuViewDouble 
         Caption         =   "&Double Size"
      End
      Begin VB.Menu mnuViewX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFull 
         Caption         =   "&Fullscreen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewStretch 
         Caption         =   "&Stretch"
      End
      Begin VB.Menu mnuViewX2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSlideshow 
         Caption         =   "&Slideshow"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpFAQ 
         Caption         =   "&Homepage"
      End
      Begin VB.Menu mnuHelpForums 
         Caption         =   "&Discussion Forums"
      End
      Begin VB.Menu mnuHelpX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUpdates 
         Caption         =   "Check for &Updates"
      End
      Begin VB.Menu mnuHelpX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About CamEVU ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewNode As String
Dim m_cIni As New cInifile

Dim TimerVal                  ' New timer value
Dim OrigTimerVal              ' Original time value
Dim CurrSelected              ' Selected cam node data
Dim CurrSelectedText          ' Text/name of selected cam
Dim IsMinimized As Boolean    ' Minimized state of form
Public StopProcess As Boolean ' If user is not an adult

Dim ACDownloadSuccess As Boolean
Dim DownloadSuccess As Boolean
Dim ACDownloadError As Boolean
Dim DownloadError As Boolean
Dim NoChangeInCam As Boolean
Dim CancelSearch As Boolean
Dim ACDone As Boolean

Dim ViewMode As Integer

Dim WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Dim WithEvents m_cSplit As cSplitter
Attribute m_cSplit.VB_VarHelpID = -1

Dim CamLastModified As String

Dim OrigImgWidth As Integer
Dim OrigImgHeight As Integer

Public Cam_Name As String
Public Cam_OwnerName As String
Public Cam_OwnerEmail As String
Public Cam_OwnerWebsite As String

Public LockToolbarStates As Boolean
Function FirstInAlphabeticalOrder(strOne As String, strTwo As String) As Long

   Dim intChar As Integer, intLen As Integer
   Dim strChar1 As String, strChar2 As String
   
      'Check to see which string has more length
      'assign intLen% the length of that string.
      If Len(strOne$) > Len(strTwo$) Then
         intLen% = Len(strOne$)
      ElseIf Len(strTwo$) > Len(strOne$) Then
         intLen% = Len(strTwo$)
      Else
         intLen% = Len(strOne$)
      End If
        
   
      For intChar% = 1 To intLen%
        strChar1$ = UCase$(Mid$(strOne$, intChar%, 1))
        strChar2$ = UCase$(Mid$(strTwo$, intChar%, 1))
           
            'if no more character's are left on a string
            'then that string automatically takes precedence.
            'So exit the function.
            If Len(strChar1$) = 0 Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Len(strChar2$) = 0 Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            End If
            
            'if character ascii value is between the ascii
            'value of 'A' and 'Z', and the other character's
            'ascii value is not. Precednce goes to the first
            'string. If that and vice-versa is false. Check
            'which ascii value is lower than the other ascii value.
            'If one if it is lower that string takes precedence. If
            'their equal, continue to the next character.
            
            If Asc(strChar1$) >= Asc("A") And Asc(strChar1$) <= Asc("Z") And Asc(strChar2$) <= Asc("A") And Asc(strChar2$) >= Asc("Z") Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Asc(strChar2$) >= Asc("A") And Asc(strChar2$) <= Asc("Z") And Asc(strChar1$) <= Asc("A") And Asc(strChar1$) >= Asc("Z") Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            ElseIf Asc(strChar1$) < Asc(strChar2$) Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Asc(strChar2$) < Asc(strChar1$) Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            End If
               
      Next intChar%

End Function

Public Function GetURLPath(strURL As String, Optional GetRidOf As String) As String

    On Error Resume Next

    strURL = Replace(strURL, GetRidOf, "")
    
    strURL = Replace(strURL, " ", "[SPACE]")
    strURL = Replace(strURL, "/", " / ")
    
    strURL = MidWord(strURL, 1, Words(strURL) - 1)
    
    strURL = Replace(strURL, " / ", "/")
    strURL = Replace(strURL, " /", "/")
    strURL = Replace(strURL, "[SPACE]", " ")
    
    GetURLPath = strURL

End Function
Public Function GetFilenameFromURL(strURL As String, Optional GetRidOf As String) As String

    On Error Resume Next

    strURL = Replace(strURL, GetRidOf, "")
    
    strURL = Replace(strURL, " ", "[SPACE]")
    strURL = Replace(strURL, "/", " / ")
    
    strURL = Word(strURL, Words(strURL))

    strURL = Replace(strURL, " / ", "/")
    strURL = Replace(strURL, " /", "/")
    strURL = Replace(strURL, "[SPACE]", " ")

    GetFilenameFromURL = strURL

End Function
Public Function LoadCamFromNode(Optional ByVal Node As MSComctlLib.Node)

On Error Resume Next
If Node = Empty Then Node = TreeView1.SelectedItem

'MsgBox Node
ACDownloadError = False
DownloadError = False
StopProcess = False
NoChangeInCam = False

If InStr(Node.Key, "WEBNODE") > 0 Then
   Timer2.Enabled = False
   Image2.Visible = False
   
   'Call ResetActiveCamObjects
   
   CurrSelected = Node.Text
   
   Inet1.Cancel
   
   NewNode = Node.Key
   NewNode = Replace(NewNode, "WEBNODE ", "")
   CurrSelectedText = Node.Text
   
   Label5.Caption = "Downloading Webcam ..."
   'Label6.Caption = "Downloading Webcam ..."
   
   If ViewMode = 4 Then frmFullscreen.Image1.Visible = True
   
   StatusBar1.Panels(1).Text = "Busy"
   'StatusBar1.Panels(2).Text = "Looking for ActiveCam specifications for: " & Node.Text ' & " (" & Node.Key & ")"
   
   Picture3.Visible = False
   Picture4.Visible = True
   
   Image2.Stretch = False
   
   'Call DownloadActiveCam(GetURLPath(Node.Key, "WEBNODE") & "activecams.evu", App.Path & "\tempac.dat")

   'If ACDownloadError = False Then
      'Call ParseActiveCam2(App.Path & "\tempac.dat", GetFilenameFromURL(Node.Key, "WEBNODE"), GetURLPath(Node.Key, "WEBNODE"))

      'If UserPref.DeleteTemps = True Then
         'On Error Resume Next
         'Kill (App.Path & "\tempac.dat")
      'End If

      'If StopProcess = True Then
      '   Picture4.Visible = False
      '   Image2.Picture = LoadPicture(App.Path & "\resource1.dat")

      '   If ViewMode = 4 Then frmFullscreen.Picture1.Picture = Image2.Picture
      '   If ViewMode = 4 Then frmFullscreen.Image1.Visible = False

      '   OrigImgWidth = Image2.Width
      '   OrigImgHeight = Image2.Height
      '   Image2.Visible = True
      '   StatusBar1.Panels(1).Text = "Ready"
      '   StatusBar1.Panels(2).Text = ""
      '   TimerVal = ""
      '   OrigTimerVal = ""
      '   Timer2.Enabled = False
      '   Exit Function
      'End If
   'End If
   
   StatusBar1.Panels(2).Text = "Downloading Cam: " & Node.Text ' & " (" & Node.Key & ")"
   
   Call DownloadFile(NewNode, App.Path & "\temp.dat")

    Picture4.Visible = False
    Image2.Picture = LoadPicture(App.Path & "\temp.dat")

    If ViewMode = 4 Then frmFullscreen.Picture1.Picture = Image2.Picture
    If ViewMode = 4 Then frmFullscreen.Image1.Visible = False
    
    OrigImgWidth = Image2.Width
    OrigImgHeight = Image2.Height

    Image2.Visible = True

    StatusBar1.Panels(1).Text = "Ready"
    StatusBar1.Panels(2).Text = ""

    TimerVal = ""
    OrigTimerVal = ""

    With m_cIni
        .Path = App.Path & "\favorites.dat"
        .Section = Node.Text
        .Key = "UpdateInterval"
        .Default = "60"
        
        Timer2.Enabled = False
        
        If LCase(.Value) = "x" Then
           If UserPref.ForceNull60Sec = True Then
              Timer2.Enabled = True
              TimerVal = "60"
              OrigTimerVal = "60"
           End If
        Else
           Timer2.Enabled = True
           
           StatusBar1.Panels(2).Text = "Refresh Scheduled in " & .Value & " Seconds"

           TimerVal = .Value
           OrigTimerVal = .Value
        End If
    End With
    
    If NoChangeInCam = False Then
        If DownloadSuccess = True And UserPref.DeleteTemps = True Then
            On Error Resume Next
            Kill (App.Path & "\temp.dat")
        End If
        
        If DownloadError = True Then
            Timer2.Enabled = False
            StatusBar1.Panels(2).Text = ""
        End If
    Else
        StatusBar1.Panels(2).Text = Node.Text & " has not changed"
    End If

Else
   Timer2.Enabled = False

   Image2.Visible = False
   Picture3.Visible = True
End If

'Form_Resize

End Function
Public Function ParseAddressChange()

On Error Resume Next

   If Left(ImageCombo1.Text, 7) = "http://" Then
      Dim FoundPagesTab As Boolean
      Dim ComboAlreadyExists As Boolean
      Dim CurrComboText As String
      
      CurrComboText = ImageCombo1.Text
      
      For i = 1 To ImageCombo1.ComboItems.Count
          If ImageCombo1.ComboItems(i).Text = "Webpages" Then
             FoundPagesTab = True
             Exit For
          End If
      Next
      
      If FoundPagesTab = False Then ImageCombo1.ComboItems.Add , , "Webpages", 6, 6
      
      For i = 1 To ImageCombo1.ComboItems.Count
          If ImageCombo1.ComboItems(i).Text = ImageCombo1.Text Then
             ComboAlreadyExists = True
          End If
      Next

      If ComboAlreadyExists = False Then
         ImageCombo1.ComboItems.Add , ImageCombo1.Text, ImageCombo1.Text, 7, 7, 2
      End If
      ImageCombo1.ComboItems.item(ImageCombo1.ComboItems.Count).Selected = True

      If Right(CurrComboText, 4) = ".jpg" Or Right(CurrComboText, 4) = ".gif" Or Right(CurrComboText, 4) = ".bmp" Then
         ' Download & Display
            Timer2.Enabled = False ' No refresh for web-graphics.
            
            Label5.Caption = "Downloading Graphic ..."
            'Label6.Caption = "Downloading Graphic ..."

            StatusBar1.Panels(1).Text = "Busy"
            StatusBar1.Panels(2).Text = "Looking for ActiveCam specifications for: " & Node.Text ' & " (" & Node.Key & ")"
            If ViewMode = 4 Then frmFullscreen.Image1.Visible = True

            Image2.Visible = False
            Image2.Stretch = False
            
            Picture3.Visible = False
            Picture4.Visible = True
            Call DownloadFile(CurrComboText, App.Path & "\temp.dat")
            Picture4.Visible = False
            If ViewMode = 4 Then frmFullscreen.Picture1.Picture = Image2.Picture
            If ViewMode = 4 Then frmFullscreen.Image1.Visible = False
            OrigImgWidth = Image2.Width
            OrigImgHeight = Image2.Height
            Image2.Picture = LoadPicture(App.Path & "\temp.dat")
            Image2.Visible = True
            
            TreeView1.Nodes(1).Selected = True
            
            If DownloadSuccess = True And UserPref.DeleteTemps = True Then
                Kill (App.Path & "\temp.dat")
                Timer2.Enabled = False
            End If
      Else
         
         Call Shell("Explorer " & ImageCombo1.Text, vbNormalFocus)
         TreeView1.Nodes(1).Selected = True
      End If

   Else
   
         If Not CurrSelected = ImageCombo1.Text Then
            For i = 1 To TreeView1.Nodes.Count
                If InStr(TreeView1.Nodes(i).Key, "WEBNODE") > 0 Then
                   If TreeView1.Nodes(i).Text = ImageCombo1.Text Then
                      CamLastModified = ""
                      TreeView1.Nodes(i).Selected = True
                      StatusBar1.Panels(1).Text = "Ready"
                      StatusBar1.Panels(2).Text = ""
                      CurrSelected = TreeView1.Nodes(i).Text
                      LoadCamFromNode TreeView1.Nodes(i)
                      Exit For
                   End If
                End If
            Next
         End If

   End If

End Function

Public Function PopupEditMnu()

On Error Resume Next

If CurrSelected = TreeView1.SelectedItem.Text Then
   mnuEditView.Caption = "&Refresh Cam"
Else
   mnuEditView.Caption = "&View Cam"
End If

        mnuEditRefresh.Visible = False
        mnuEditX1.Visible = False
        mnuEditX2.Visible = False
        mnuEditCopy.Visible = False
        
        mnuEditView.Visible = True
        mnuEditX3.Visible = True
        
        PopupMenu mnuEdit
        
        mnuEditX3.Visible = False
        mnuEditView.Visible = False

        mnuEditRefresh.Visible = True
        mnuEditX1.Visible = True
        mnuEditX2.Visible = True
        mnuEditCopy.Visible = True

End Function
Public Function PopupEditMnu2()

On Error Resume Next

If CurrSelected = TreeView1.SelectedItem.Text Then
   mnuEditView.Caption = "&Refresh Cam"
Else
   mnuEditView.Caption = "&View Cam"
End If

        mnuEditRefresh.Visible = False
        mnuEditX1.Visible = False
        mnuEditX2.Visible = False
        mnuEditCopy.Visible = False
        
        mnuEditView.Visible = True
        mnuEditX3.Visible = True
        mnuEditX4.Visible = True
        
        mnuEditAdd.Visible = True
        
        PopupMenu mnuEdit, , , , mnuEditAdd

        mnuEditAdd.Visible = False
        
        mnuEditX4.Visible = False
        mnuEditX3.Visible = False
        mnuEditView.Visible = False

        mnuEditRefresh.Visible = True
        mnuEditX1.Visible = True
        mnuEditX2.Visible = True
        mnuEditCopy.Visible = True

End Function
Public Function ReloadAddressCombo()

On Error Resume Next

ImageCombo1.ComboItems.Clear

    StatusBar1.Panels(1).Text = "Loading"
    StatusBar1.Panels(2).Text = "Populating Address List ..."

ImageCombo1.ComboItems.Add , , "Homepage", 1, 1
ImageCombo1.ComboItems.Add , , "Webcams", 2, 2

For i = 1 To TreeView1.Nodes.Count
   If InStr(TreeView1.Nodes(i).Key, "WEBNODE") > 0 Then
        ImageCombo1.ComboItems.Add , , TreeView1.Nodes(i).Text, 8, 8, 2
   End If
Next

ImageCombo1.ComboItems.item(1).Selected = True

    StatusBar1.Panels(1).Text = "Ready"
    StatusBar1.Panels(2).Text = ""

End Function

Public Function ReloadCamList()

On Error Resume Next

' ================================
' Parse Cam List
' ================================

    StatusBar1.Panels(1).Text = "Loading"
    StatusBar1.Panels(2).Text = "Populating Cam List ..."

    TreeView1.Nodes.Clear

    TreeView1.Nodes.Add , tvwChild, "AddCam", "Add a Cam", 6, 6

    TreeView1.Nodes.Add , tvwChild, "Cams", "Webcams", 5, 5
    TreeView1.Nodes(1).Selected = True
    TreeView1.Nodes.item("Cams").Expanded = True

    ' Add Cam List Here
    Dim sSections() As String
    Dim iSectionCount As Long

    With m_cIni
        .Path = App.Path & "\favorites.dat"
        
        .EnumerateAllSections sSections(), iSectionCount
        For iSection = 1 To iSectionCount
            'lstIni.AddItem "[" & sSections(iSection) & "]"
            .Section = sSections(iSection)
            .Key = "Address"
            .Default = ""
            
            TreeView1.Nodes.Add "Cams", tvwChild, "WEBNODE " & .Value, Trim(sSections(iSection)), 1, 2
        Next iSection
    End With
    
    'TreeView1.Nodes.Add , tvwChild, "OrgCam", "Organize", 7, 7

    TreeView1.Nodes.Add , tvwChild, "Dirs", "Web Directories", 8, 8
    TreeView1.Nodes("Cams").Sorted = True
    TreeView1.Nodes("Dirs").Sorted = True
    TreeView1.Nodes("Dirs").Expanded = True
    
    ' Add Dir List Here

    With m_cIni
        .Path = App.Path & "\directories.dat"
        
        'If .Success = False Then
        
        .EnumerateAllSections sSections(), iSectionCount
        For iSection = 1 To iSectionCount
            'lstIni.AddItem "[" & sSections(iSection) & "]"
            .Section = sSections(iSection)
            .Key = "Address"
            .Default = ""
            
            TreeView1.Nodes.Add "Dirs", tvwChild, "DIRNODE " & .Value, Trim(sSections(iSection)), 9, 9
        Next iSection
    End With

    StatusBar1.Panels(1).Text = "Ready"
    StatusBar1.Panels(2).Text = ""

End Function

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)

With TreeView1
    .Top = NewHeight
    .Height = Me.Height - NewHeight - StatusBar1.Height - 660
    picSplit.Height = .Height
    Picture2.Top = .Top
    Picture2.Height = .Height
End With

End Sub

Private Sub Form_Load()

On Error Resume Next


If WordCount(Command, "/updatenow") > 0 Then
   frmUpdate.Show 1, Me
   Unload Me
   End
End If



ViewMode = 1

LoadPreferences

    Set m_cSplit = New cSplitter
    m_cSplit.Initialise picSplit, Me

If UserPref.ShowSystemTray = True Then
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeToolTip "CamEVU"
    gSysTray.ChangeIcon Me.Icon
    gSysTray.DefaultDblClk = True
End If

If UserPref.UseHovers = True Then TreeView1.HotTracking = True

If UserPref.ShowAddyBar = False Then Toolbar2.Visible = False: Picture1.Visible = False
If UserPref.ShowStatusBar = False Then StatusBar1.Visible = False

' ================================
' Display CamEVU version correctly
' ================================

If App.Revision = "0" Then
    Label4.Caption = "Version " & App.Major & "." & App.Minor
Else
    Label4.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End If

Shape1.Left = Label4.Left - 82.5
Shape1.Width = Label4.Width + 165

' ================================
' Grab positioning options
' ================================

If UserPref.LoadWinPos = True Then
    If UserPref.XPos > "" And UserPref.YPos > "" Then
       frmMain.Left = UserPref.XPos
       frmMain.Top = UserPref.YPos
    End If

    frmMain.Height = UserPref.HSet
    frmMain.Width = UserPref.WSet
End If

' ================================
' Reload Webcam List
' ================================

ReloadCamList
ReloadAddressCombo

If UserPref.KeepOnTop = True Then
   SetTopmost Me, True
Else
   SetTopmost Me, False
End If

' ================================
' Edit Menu
' ================================

mnuEditX5.Visible = False

' ================================

Me.Show

End Sub
Private Function StripDelimitedItem(startStrg As String, delimiter As String) As String

  'take a string separated by nulls,
  'split off 1 item, and shorten the string
  'so the next item is ready for removal.
   Dim pos As Long
   Dim item As String
   
   pos = InStr(1, startStrg, delimiter)
   
   If pos Then

      StripDelimitedItem = Mid$(startStrg, 1, pos)
      startStrg = Mid$(startStrg, pos + 1, Len(startStrg))
    
   End If

End Function


Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean
'strDestination = App.Path & "\temp.jpg"
' Funtion DownloadFile: Download a file via HTTP
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

   LockToolbarStates = True
   For i = 1 To Toolbar1.Buttons.Count
       DoEvents
       Toolbar1.Buttons(i).Enabled = False
   Next i
   Toolbar1.Buttons(2).Enabled = True

If UserPref.LockLists = True Then
   TreeView1.Enabled = False
   ImageCombo1.Enabled = False
End If

Me.MousePointer = 11

Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const ROLLBACK As Long = 4096   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim lastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host

'On Local Error GoTo InternetErrorHandler

DownloadError = False

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
'''SourceLabel = Empty
'''TimeLabel = Empty
'''ToLabel = Empty
'''RateLabel = Empty

StartDownload:

If blnResume Then
    '''StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    '''StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    If blnResume Then
        ' GET file, sending the magic resume input header...
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    Else
        ' Standard file GET
        .Execute , "GET"
    End If
End With

' While initiating connection, yield CPU to Windows
While Inet1.StillExecuting
    DoEvents
    ' If user pressed Cancel button on StatusForm
    ' then fail, cancel, and exit this download
    If CancelSearch Then GoTo ExitDownload
Wend

''StatusLabel = "Saving:"
''SourceLabel = FitText(SourceLabel, strHost & " from " & Inet1.RemoteHost)
''ToLabel = FitText(ToLabel, strDestination)

If UserPref.OnlyDownloadIfUpdated = True Then
    If CamLastModified = Inet1.GetHeader("Last-modified") Then
       Inet1.Cancel
       'MsgBox "Image did not change."
       NoChangeInCam = True
       GoTo ExitDownload
    Else
       CamLastModified = Inet1.GetHeader("Last-modified")
    End If
End If

' Get first header ("HTTP/X.X XXX ...")
strHeader = Inet1.GetHeader

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            'If MsgBox("The server is unable to resume this download." & _
                      vbCr & vbCr & _
                      "Do you want to continue anyway?", _
                      vbExclamation + vbYesNo, _
                      "Unable to Resume Download") = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                'Else
                    ' No - cancel
                    'CancelSearch = True
                    'GoTo ExitDownload
               ' End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        If UserPref.HideErrors = False Then MsgBox "Nothing to download!", _
               vbInformation, _
               "No Content"
        DownloadError = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        If UserPref.HideErrors = False Then MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        DownloadError = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        If UserPref.HideErrors = False Then MsgBox "Cam download failed!" & vbCrLf & "The image, " & _
               """" & Inet1.URL & """" & _
               " was not found.", _
               vbCritical, _
               "Error Downloading File"
        DownloadError = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        If UserPref.HideErrors = False Then MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        DownloadError = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        If UserPref.HideErrors = False Then MsgBox "The server returned the following response:" & vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error Downloading File"
        DownloadError = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    lastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        If UserPref.HideErrors = False Then MsgBox "There is not enough free space on disk for this file." _
               & vbCr & vbCr & "Please free up some disk space and try again.", _
               vbCritical, _
               "Insufficient Disk Space"
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
With ProgressBar1
    .Visible = True
    .Min = 0
    .Value = 0
    .Max = lngFileLength
End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


'On Local Error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile()        ' Set intFile to an unused file.
' Open a file to write to.
Open strDestination For Binary Access Write As #intFile
' If resuming, then seek byte position in downloaded file
' where we last left off...
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    'sglRate = lngBytesReceived / (Timer - lastTime)
    'sglTime = (lngFileLength - lngBytesReceived) / sglRate
    'TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    'RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar1.Value = lngBytesReceived
    'Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " of " & strFile & " Completed"
Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
Close #intFile

ExitDownload:

   LockToolbarStates = False
   For i = 1 To Toolbar1.Buttons.Count
       DoEvents
       Toolbar1.Buttons(i).Enabled = True
   Next i
   Toolbar1.Buttons(2).Enabled = False

If UserPref.LockLists = True Then
   TreeView1.Enabled = True
   ImageCombo1.Enabled = True
End If
Me.MousePointer = 0
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    'StatusLabel = "Download completed!"
    DownloadSuccess = True
    LockToolbarStates = False
    ProgressBar1.Visible = False
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            'If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to try to resume the download." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to abort the download.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            'End If
        End If
    End If
    ' No or unresumable failure:
    DownloadSuccess = False
    LockToolbarStates = False
    ProgressBar1.Value = 0
End If

CleanUp:
' Close AVI
'''Animation1.Close

If UserPref.LockLists = True Then TreeView1.Enabled = True
If UserPref.LockLists = True Then ImageCombo1.Enabled = True
Me.MousePointer = 0

' Make sure that the Internet connection is closed...
Inet1.Cancel
LockToolbarStates = False
' Delete any partially downloaded file
'If CancelSearch And Dir(strDestination) > Empty Then Kill strDestination
' ...and exit this function
'''Form2.Show
'''Form1.CDIal.FileName = "C:\windows\command\act.jpg"
'''If Form1.CDIal.FileName <> "" Then Set Form1.mcscrollpic1.Picture = LoadPicture(Form1.CDIal.FileName)
'''Unload frmDownload

Exit Function

InternetErrorHandler:
    ' This is a catch-all that hasn't been fired once in the
    ' almost 2 yrs this code has existed, so...
    CancelSearch = True
    If UserPref.HideErrors = False Then MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    Resume Next
    
FileErrorHandler:
    If Err.Number <> 9 Then
        ' Err# 9 occurs when UBound(bData,1) < 0
        If UserPref.HideErrors = False Then MsgBox "Cannot write file to disk." & _
               vbCr & vbCr & _
               "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, _
               "Error Downloading File"
        CancelSearch = True
    End If
    Err.Clear
    GoTo CleanUp
    
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    m_cSplit.MouseMove x

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    m_cSplit.MouseUp x

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error Resume Next
    If UserPref.ShowSystemTray = True Then gSysTray.RemoveFromSysTray
    'Call ResetActiveCamObjects ' Cleans Up Memory, Maybe?

End Sub

Private Sub Form_Resize()

On Error Resume Next

If UserPref.ShowSystemTray = True Then
    If Me.WindowState = vbMinimized Then
        If IsMinimized = False Then
           gSysTray.MinToSysTray
           IsMinimized = True
        End If
    Else
        IsMinimized = False
    End If
End If

' ==========================================

With CoolBar1.Bands(2)
    If UserPref.ShowAddyBar = False Then
       If Not .Visible = False Then .Visible = False
    Else
       If Not .Visible = True Then .Visible = True
    End If
End With

' ==========================================

'ImageCombo1.Width = Picture5.Width - 700

With TreeView1
    .Top = CoolBar1.Height
    .Height = Me.Height - CoolBar1.Height - StatusBar1.Height - 660
    picSplit.Height = .Height
End With

Picture2.Top = TreeView1.Top
Picture2.Height = TreeView1.Height
Picture2.Width = Me.Width - TreeView1.Width - 190

If Picture3.Visible = True Then
   With Picture3
        .Left = Picture2.Left + Picture2.Width / 2 - .Width / 2
        .Top = Picture2.Top + Picture2.Height / 2 - .Height / 2
   End With
Else
   
   With Image2
   If ViewMode = 1 Then
      If Not .Width = OrigImgWidth Then .Width = OrigImgWidth
      If Not .Height = OrigImgHeight Then .Height = OrigImgHeight
      If Not .Left = Picture2.Width / 2 - .Width / 2 Then .Left = Picture2.Width / 2 - .Width / 2
      If Not .Top = Picture2.Height / 2 - .Height / 2 Then .Top = Picture2.Height / 2 - .Height / 2
      If Not .Stretch = False Then .Stretch = False
   ElseIf ViewMode = 2 Then
      If Not .Width = OrigImgWidth / 2 Then .Width = OrigImgWidth / 2
      If Not .Height = OrigImgHeight / 2 Then .Height = OrigImgHeight / 2
      If Not .Left = Picture2.Width / 2 - .Width / 2 Then .Left = Picture2.Width / 2 - .Width / 2
      If Not .Top = Picture2.Height / 2 - .Height / 2 Then .Top = Picture2.Height / 2 - .Height / 2
      If Not .Stretch = True Then .Stretch = True
   ElseIf ViewMode = 3 Then
      If Not .Width = OrigImgWidth * 2 Then .Width = OrigImgWidth * 2
      If Not .Height = OrigImgHeight * 2 Then .Height = OrigImgHeight * 2
      If Not .Left = Picture2.Width / 2 - .Width / 2 Then .Left = Picture2.Width / 2 - .Width / 2
      If Not .Top = Picture2.Height / 2 - .Height / 2 Then .Top = Picture2.Height / 2 - .Height / 2
      If Not .Stretch = True Then .Stretch = True
   ElseIf ViewMode = 5 Then
      If Not .Width = Picture2.Width Then .Width = Picture2.Width
      If Not .Height = Picture2.Height Then .Height = Picture2.Height
      If Not .Left = 0 Then .Left = 0
      If Not .Top = 0 Then .Top = 0
      If Not .Stretch = True Then .Stretch = True
   End If
   End With

   With Picture4
        If Not .Left = Picture2.Width / 2 - .Width / 2 Then .Left = Picture2.Width / 2 - .Width / 2
        If Not .Top = Picture2.Height / 2 - .Height / 2 Then .Top = Picture2.Height / 2 - .Height / 2
   End With
End If

With ProgressBar1
    If UserPref.ShowStatusBar = True Then
       If Not .Top = StatusBar1.Top + 50 Then .Top = StatusBar1.Top + 50
       If Not .Left = Me.Width - .Width - 400 Then .Left = Me.Width - .Width - 400
    Else
       If Not .Top = Me.Height - 890 Then .Top = Me.Height - 890
       If Not .Left = 0 Then .Left = 0
       If Not .Width = Me.ScaleWidth Then .Width = Me.ScaleWidth
    End If
End With

End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

' Save Window Positions
Call SavePosition(frmMain)

'If UserPref.ShowSystemTray = True Then Call RemoveFromTray

Inet1.Cancel
Timer1.Enabled = False
Timer2.Enabled = False

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Shell("Explorer http://www.erroneousdata.com/camevu.shtml", vbNormalFocus)

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Label2.ForeColor = &HFF0000
'Label3.ForeColor = &HFF0000

'StatusBar1.Panels.Item(3).Text = "Open Address: http://www.erroneousdata.com/camevu/"
'StatusBar1.Refresh

End Sub


Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next
If Button = vbRightButton Then
    If TreeView1.SelectedItem.Text <> "" Then
        PopupEditMnu2
    End If
End If

End Sub


Private Sub ImageCombo1_Change()

If Left(ImageCombo1.Text, 7) = "http://" Then
   ImageCombo1.ForeColor = &HFF0000
Else
   ImageCombo1.ForeColor = &H80000008
End If

End Sub

Private Sub ImageCombo1_Click()

ImageCombo1_Change
ParseAddressChange

End Sub


Private Sub ImageCombo1_GotFocus()

ImageCombo1.SelStart = 0
ImageCombo1.SelLength = Len(ImageCombo1.Text)

End Sub

Private Sub ImageCombo1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    ParseAddressChange
End If

End Sub

Private Sub m_cSplit_SplitComplete()
    'Form_Resize
    
    On Error Resume Next
    
    If picSplit.Left <= 1175 Then picSplit.Left = 1175
    If picSplit.Left >= 3525 Then picSplit.Left = 3525
    
    TreeView1.Width = picSplit.Left
    Picture2.Left = picSplit.Left + picSplit.Width
    
    Form_Resize
End Sub


Private Sub mnuEditAdd_Click()

mnuFileAdd_Click

End Sub

Private Sub mnuEditCopy_Click()

Clipboard.SetData Image2.Picture

End Sub

Private Sub mnuEditModify_Click()

frmAddModify.camTitle = TreeView1.SelectedItem.Text
frmAddModify.camLocal = Replace(TreeView1.SelectedItem.Key, "WEBNODE ", "")

    With m_cIni
        .Path = App.Path & "\favorites.dat"
        .Section = TreeView1.SelectedItem.Text
        .Key = "UpdateInterval"
        .Default = "60"
        frmAddModify.camRate = .Value
    End With

frmAddModify.ModifyMode = True
frmAddModify.Show 1

Call LoadCamFromNode

End Sub

Private Sub mnuEditRefresh_Click()

Dim SelectedItemTXT As String
Dim SelectedItemKEY As String
SelectedItemTXT = TreeView1.SelectedItem.Text
SelectedItemKEY = TreeView1.SelectedItem.Key

ReloadCamList

For i = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(i).Text = SelectedItemTXT And TreeView1.Nodes(i).Key = SelectedItemKEY Then
       TreeView1.Nodes(i).Selected = True
       LoadCamFromNode
       Exit For
    End If
Next

End Sub

Private Sub mnuEditRemove_Click()

Dim msgReply

msgReply = MsgBox("Are you certain you wish to remove the cam, """ & TreeView1.SelectedItem.Text & """?", vbYesNo)

If msgReply = vbYes Then
    With m_cIni
         .Section = TreeView1.SelectedItem.Text
         .DeleteSection
    End With
    
    ReloadCamList
End If

End Sub


Private Sub mnuEditView_Click()

'If CurrSelected = Node.Text Then
'   mnuEditRefresh_Click
'Else
   LoadCamFromNode
'End If

End Sub

Private Sub mnuFileAdd_Click()

frmAddModify.ModifyMode = False
frmAddModify.Show 1

End Sub

Private Sub mnuFileExit_Click()

Unload Me

End Sub

Private Sub mnuFileHide_Click()

Me.WindowState = 1

End Sub


Private Sub mnuFilePrefs_Click()

frmPrefs.Show 1

End Sub

Private Sub mnuFilePrint_Click()

'Printer.NewPage
'Dim CdlgEx1 As New CdlgEx
Dim MyCenteredText As String
'CdlgEx1.ShowPrinter

MyCenteredText = TreeView1.SelectedItem.Text
CamEVUInfoTXT = "Powered by CamEVU"
CamEVUInfoTXT2 = "http://www.erroneousdata.com"

'Printer.Page
Printer.PaintPicture Image2.Picture, Printer.ScaleWidth / 2 - Image2.Width / 2, Printer.ScaleHeight / 2 - Image2.Height / 2

Printer.FontName = "Tahoma"
Printer.FontSize = 12
Printer.FontBold = True
Printer.ForeColor = RGB(0, 0, 0)

Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth(MyCenteredText) / 2
Printer.CurrentY = Printer.ScaleHeight / 2 + Image2.Height + 75

Printer.Print MyCenteredText

Printer.FontName = "Tahoma"
Printer.FontSize = 8
Printer.FontBold = False
Printer.FontUnderline = True
Printer.ForeColor = RGB(0, 0, 0)

Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(CamEVUInfoTXT)
Printer.CurrentY = Printer.ScaleHeight / 2 + Image2.Height + 75

Printer.EndDoc

End Sub

Private Sub mnuFileSaveAs_Click()

  'used in call setup
   Dim sFilters As String
   
  'used after call
   Dim buff As String
   Dim sLongname As String
   Dim sShortname As String

  'create a string of filters for the dialog
   sFilters = "All Files" & vbNullChar & "*.*" & vbNullChar & _
              "24-bit Bitmap Format" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
   
   '"24-bit Bitmap" & vbNullChar & "*.bmp" & vbNullChar & _
              "JPEG File Interchange Format" & vbNullChar & "*.bas" & vbNullChar & _

   With OFN
   
      .nStructSize = Len(OFN)
      .hWndOwner = Me.hWnd
      .sFilter = sFilters
      .nFilterIndex = 2
      .sFile = "Untitled.bmp" & Space$(1024) & vbNullChar & vbNullChar
      .nMaxFile = Len(.sFile)
      .sDefFileExt = "bmp" & vbNullChar & vbNullChar
      .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .nMaxTitle = Len(OFN.sFileTitle)
      .sInitialDir = "c:\" & vbNullChar & vbNullChar
      .sDialogTitle = "Save Webcam Image As..."
      .flags = OFS_FILE_SAVE_FLAGS

   End With
   
  'call the API
   If GetSaveFileName(OFN) Then
    
     'see illustration for descriptions
      Call SavePicture(Image2.Picture, OFN.sFile)
  
  End If

End Sub

Private Sub mnuHelpAbout_Click()

frmAbout.Show 1, Me

End Sub


Private Sub mnuHelpFAQ_Click()

On Error Resume Next

Call Shell("Explorer http://www.erroneousdata.com/camevu.shtml", vbNormalFocus)

End Sub

Private Sub mnuHelpForums_Click()

On Error Resume Next

Call Shell("Explorer http://www.chatbear.com/?444", vbNormalFocus)

End Sub


Private Sub mnuHelpUpdates_Click()

On Error Resume Next

frmUpdate.Show 1, Me

End Sub

Private Sub mnuViewDouble_Click()

   ViewMode = 3

End Sub

Private Sub mnuViewFull_Click()

   Dim OldMode As Integer

   OldMode = ViewMode
   ViewMode = 4
   Timer1.Enabled = False
   frmFullscreen.Show 1 ' Show Full Screen
   
   'Do While frmFullscreen.Visible = True
   '   DoEvents
   'Loop
   
   Timer1.Enabled = True
   ViewMode = OldMode

End Sub

Private Sub mnuViewHalf_Click()

   ViewMode = 2

End Sub

Private Sub mnuViewNormal_Click()

   ViewMode = 1

End Sub

Private Sub mnuViewStretch_Click()

   ViewMode = 5

End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

m_cSplit.MouseDown x

End Sub


Private Sub Picture5_Resize()

ImageCombo1.Width = Picture5.Width - 700

End Sub


Private Sub Timer1_Timer()

If IsMinimized = True Then Exit Sub

On Error Resume Next
Form_Resize

If LockToolbarStates = False Then
    If InStr(TreeView1.SelectedItem.Key, "WEBNODE") > 0 Then
       Toolbar1.Buttons(4).Enabled = True
       Toolbar1.Buttons(5).Enabled = True
       Toolbar1.Buttons(7).Enabled = True
       Toolbar1.Buttons(8).Enabled = True
       Toolbar1.Buttons(10).Enabled = True
       mnuFileSaveAs.Enabled = True
       mnuFilePrint.Enabled = True
       mnuEditRemove.Enabled = True
       mnuEditModify.Enabled = True
       mnuEditCopy.Enabled = True
       
       mnuView.Enabled = True
    Else
       Toolbar1.Buttons(4).Enabled = False
       Toolbar1.Buttons(5).Enabled = False
       Toolbar1.Buttons(7).Enabled = False
       Toolbar1.Buttons(8).Enabled = False
       Toolbar1.Buttons(10).Enabled = False
       mnuFileSaveAs.Enabled = False
       mnuFilePrint.Enabled = False
       mnuEditRemove.Enabled = False
       mnuEditModify.Enabled = False
       mnuEditCopy.Enabled = False
       
       mnuView.Enabled = False
    End If
End If

mnuViewNormal.Checked = False
mnuViewHalf.Checked = False
mnuViewDouble.Checked = False
mnuViewFull.Checked = False
mnuViewStretch.Checked = False
mnuViewCube.Checked = False

If ViewMode = 1 Then
   mnuViewNormal.Checked = True
ElseIf ViewMode = 2 Then
   mnuViewHalf.Checked = True
ElseIf ViewMode = 3 Then
   mnuViewDouble.Checked = True
ElseIf ViewMode = 4 Then
   mnuViewFull.Checked = True
ElseIf ViewMode = 5 Then
   mnuViewStretch.Checked = True
ElseIf ViewMode = 6 Then
   mnuViewCube.Checked = True
End If

'If Me.WindowState = 2 Then
'   Call RemoveFromTray
'   Me.Show
'Else
'   Call AddToTray(Me.Icon, Me.Caption, Me)
'   Me.Hide
'End If

End Sub


Private Sub Timer2_Timer()

If TimerVal = "0" Then
    StatusBar1.Panels(2).Text = "Refreshing ..."
    
    If Image2.Visible = True Then
       'Image2.Visible = False
    
       Label5.Caption = "Refreshing Webcam ..."
       'Label6.Caption = "Refreshing Webcam ..."
       
       Timer1.Enabled = False
       
       Picture4.Left = Image2.Left + Image2.Width / 2 - Picture4.Width / 2
       Picture4.Top = Image2.Top + Image2.Height + 355
       
       If ViewMode = 4 Then frmFullscreen.Image1.Visible = True
    
       Picture4.Visible = True
       Call DownloadFile(NewNode, App.Path & "\temp.dat")
       Picture4.Visible = False
       Image2.Picture = LoadPicture(App.Path & "\temp.dat")
       
       If ViewMode = 4 Then frmFullscreen.Picture1.Picture = Image2.Picture
       If ViewMode = 4 Then frmFullscreen.Image1.Visible = False
       'Image2.Visible = True
       
       Timer1.Enabled = True
       
        If DownloadSuccess = True And UserPref.DeleteTemps = True Then
            Kill (App.Path & "\temp.dat")
        End If
    End If

    TimerVal = OrigTimerVal
    StatusBar1.Panels(2).Text = "Refresh Scheduled in " & TimerVal & " Seconds"
Else
    TimerVal = TimerVal - 1
    StatusBar1.Panels(2).Text = "Refresh Scheduled in " & TimerVal & " Seconds"
End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

    Case 1 ' New
    mnuFileAdd_Click
    
    Case 2 ' Stop
    CancelSearch = True
    
    Case 4 ' Modify
    mnuEditModify_Click
    
    Case 5 ' Remove
    mnuEditRemove_Click
    
    Case 7 ' Copy
    mnuEditCopy_Click
    
    Case 8 ' Save
    mnuFileSaveAs_Click
    
    Case 10 ' Print
    mnuFilePrint_Click
    
    Case 12 ' Help
    mnuHelpFAQ_Click

End Select

End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next
If Button = vbRightButton Then
    If TreeView1.SelectedItem.Text <> "" Then
        PopupEditMnu
    End If
End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

If InStr(Node.Key, "WEBNODE") > 0 Then
   If Not CurrSelected = Node.Text Then
        CamLastModified = ""
        
        StatusBar1.Panels(1).Text = "Ready"
        StatusBar1.Panels(2).Text = ""
        
        For i = 1 To ImageCombo1.ComboItems.Count
           If ImageCombo1.ComboItems(i).Text = Node.Text Then
             ImageCombo1.ComboItems(i).Selected = True
           End If
        Next
        
        Call LoadCamFromNode(Node)
   End If
ElseIf Node.Text = "Add a Cam" Then
   'ResetActiveCamObjects
   
   StatusBar1.Panels(1).Text = "Ready"
   StatusBar1.Panels(2).Text = ""
   
   Timer2.Enabled = False
   Image2.Visible = False
   Picture3.Visible = True

   frmAddModify.ModifyMode = False
   frmAddModify.Show 1
   TreeView1.Nodes(1).Selected = True
   
   CurrSelected = "AddCam_____"
ElseIf InStr(Node.Key, "DIRNODE") > 0 Then
   If Not CurrSelected = Node.Text Then
        Call Shell("Explorer " & Replace(Node.Key, "DIRNODE ", ""), vbNormalFocus)
   End If
Else
   StatusBar1.Panels(1).Text = "Ready"
   StatusBar1.Panels(2).Text = ""
   
   Timer2.Enabled = False
   Image2.Visible = False
   Picture3.Visible = True
   
   CurrSelected = "WebCamLST_____"
End If

End Sub


