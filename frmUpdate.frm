VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check for Updates"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   4815
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   3960
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checking for available updates ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmUpdate.frx":000C
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CancelSearch As Boolean
Public NoShowOKMsg As Boolean
Private Sub Command1_Click()

CancelSearch = True

End Sub

Private Sub Form_Load()

AddOfficeBorder (Command1.hWnd)
Me.Visible = True

' ============================================
Dim bData() As Byte             ' Data var
Dim strHeader As Long ' 1024
Dim FoundUpdate As Boolean

CancelSearch = False
NoShowOKMsg = False
FoundUpdate = False

' Download update file
With Inet1
    .URL = "http://www.erroneousdata.com/camevu.updat"
    .UserName = ""
    .Password = ""
    .Execute , "GET"
End With

' While initiating connection, yield CPU to Windows
While Inet1.StillExecuting
    DoEvents
    ' If user pressed Cancel button on StatusForm
    ' then fail, cancel, and exit this download
    If CancelSearch Then GoTo ExitDownload
Wend

strTempOrig = Inet1.GetChunk(1024)

strTemp = Replace(strTempOrig, ".", " . ")
If App.Major < Word(strTemp, 1) Then
   Label2.Caption = "Version " & strTempOrig & " now available"
   strmsg = "A new major upgrade is available for CamEVU. It is" & vbCrLf & "highly recommened you download this update." & vbCrLf & vbCrLf & "Would you like to do so now?"
   FoundUpdate = True
ElseIf App.Minor < Word(strTemp, 3) Then
   Label2.Caption = "Version " & strTempOrig & " now available"
   strmsg = "A new update is available for CamEVU. It is" & vbCrLf & "not absolutely necessary you download this patch," & vbCrLf & "but it may resolve issues with your software and" & vbCrLf & "help you enjoy your experience more." & vbCrLf & vbCrLf & "Would you like to do so now?"
   FoundUpdate = True
ElseIf App.Revision < Word(strTemp, 5) Then
   Label2.Caption = "Version " & strTempOrig & " now available"
   strmsg = "A new revision patch is available for CamEVU. It is" & vbCrLf & "not absolutely necessary you download this patch," & vbCrLf & "but it may resolve issues with your software and" & vbCrLf & "help you enjoy your experience more." & vbCrLf & vbCrLf & "Would you like to do so now?"
   FoundUpdate = True
Else
   FoundUpdate = False
End If

If FoundUpdate = False Then
   'If NoShowOKMsg = True Then
      MsgBox "You are currently running the latest version of CamEVU." & vbCrLf & "No updates are necessary at this time.", vbInformation, "Update Notice"
   'End If
Else
   strMsg2 = MsgBox(strmsg, vbQuestion Or vbYesNo, "Update Available")
   
   If strMsg2 = vbYes Then
      Call Shell("Explorer ftp://ftp.erroneousdata.com/software/camevu.exe", vbNormalFocus)
      Unload Me
      Unload frmMain
      End
   End If
End If

GoTo ExitDownload





ExitDownload:
   Inet1.Cancel
   Unload Me

End Sub

