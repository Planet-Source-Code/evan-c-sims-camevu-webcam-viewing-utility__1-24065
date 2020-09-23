Attribute VB_Name = "mdlTopMost"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" _
     (ByVal hWnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, _
     ByVal cx As Long, ByVal cy As Long, _
     ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Sub SetTopmost(frm As Form, bTopmost As Boolean)
     Dim i As Long
     If bTopmost = True Then
          i = SetWindowPos(frm.hWnd, HWND_TOPMOST, _
               0, 0, 0, 0, SWP_WNDFLAGS)
     Else
          i = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, _
               0, 0, 0, 0, SWP_WNDFLAGS)
     End If
End Sub

