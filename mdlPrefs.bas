Attribute VB_Name = "mdlPrefs"
Public UserPref As UserPrefs
Public UserLicense As UserLic

Type UserPrefs
     LoadWinPos As Boolean
     AutoUpdate As Boolean
     ShowAddyBar As Boolean
     ShowStatusBar As Boolean
     ShowSystemTray As Boolean
     
     UseHovers As Boolean
     
     DeleteTemps As Boolean
     HideErrors As Boolean
     ForceNull60Sec As Boolean
     
     LockLists As Boolean
     KeepOnTop As Boolean
     
     OnlyDownloadIfUpdated As Boolean ' Tech Beta 3
     
     XPos As String ' X-position of form
     YPos As String ' Y-position of form
     HSet As String ' Height of form
     WSet As String ' Width of form
End Type

Type UserLic
     LicUser As String
     LicCode As String
End Type
Public Function ApplyPreferences()

    Call SavePreferences
    Call LoadPreferences

End Function

Public Function ConvertBooleanToInteger(Boolean2Convert As Boolean) As Integer

   If Boolean2Convert = True Then
      ConvertBooleanToInteger = 1
   Else
      ConvertBooleanToInteger = 0
   End If

End Function
Public Function ConvertIntegerToBoolean(Integer2Convert As Integer) As Boolean

   If Integer2Convert = "1" Then
      ConvertIntegerToBoolean = True
   Else
      ConvertIntegerToBoolean = False
   End If

End Function
Public Function LoadPrefsForDialog()

   With frmPrefs
    
       .Check1.Value = ConvertBooleanToInteger(UserPref.LoadWinPos)
       .Check2.Value = ConvertBooleanToInteger(UserPref.AutoUpdate)
       .Check3.Value = ConvertBooleanToInteger(UserPref.ShowAddyBar)
       .Check4.Value = ConvertBooleanToInteger(UserPref.ShowStatusBar)
       .Check5.Value = ConvertBooleanToInteger(UserPref.ShowSystemTray)
       
       .Check6.Value = ConvertBooleanToInteger(UserPref.UseHovers)
       
       .Check7.Value = ConvertBooleanToInteger(UserPref.DeleteTemps)
       .Check8.Value = ConvertBooleanToInteger(UserPref.HideErrors)
       .Check9.Value = ConvertBooleanToInteger(UserPref.ForceNull60Sec)

       .Check10.Value = ConvertBooleanToInteger(UserPref.LockLists)
       .Check11.Value = ConvertBooleanToInteger(UserPref.KeepOnTop)
       
       .Check13.Value = ConvertBooleanToInteger(UserPref.OnlyDownloadIfUpdated)

   End With

End Function
Public Function LoadPreferences()

    ' Load User Preferences

     UserPref.LoadWinPos = GetSetting(App.Title, "Prefs", "LoadWinPos", "True")
     UserPref.AutoUpdate = GetSetting(App.Title, "Prefs", "AutoUpdate", "True")
     UserPref.ShowAddyBar = GetSetting(App.Title, "Prefs", "ShowAddyBar", "True")
     UserPref.ShowStatusBar = GetSetting(App.Title, "Prefs", "ShowStatusBar", "True")
     UserPref.ShowSystemTray = GetSetting(App.Title, "Prefs", "ShowSystemTray", "True")
     
     UserPref.UseHovers = GetSetting(App.Title, "Prefs", "UseHovers", "True")
     
     UserPref.DeleteTemps = GetSetting(App.Title, "Prefs", "DeleteTemps", "False")
     UserPref.HideErrors = GetSetting(App.Title, "Prefs", "HideErrors", "True")
     UserPref.ForceNull60Sec = GetSetting(App.Title, "Prefs", "ForceNull60Sec", "False")

     UserPref.LockLists = GetSetting(App.Title, "Prefs", "LockLists", "True")
     UserPref.KeepOnTop = GetSetting(App.Title, "Prefs", "KeepOnTop", "False")
     
     UserPref.OnlyDownloadIfUpdated = GetSetting(App.Title, "Prefs", "CheckModified", "True")

     UserPref.XPos = GetSetting(App.Title, "Position", "X", "")
     UserPref.YPos = GetSetting(App.Title, "Position", "Y", "")
     UserPref.HSet = GetSetting(App.Title, "Position", "H", "5850")
     UserPref.WSet = GetSetting(App.Title, "Position", "W", "7695")

End Function
Public Function SavePosition(FormPos As Form)

If FormPos.WindowState = 2 Then
     Call SaveSetting(App.Title, "Prefs", "X", FormPos.Left)
     Call SaveSetting(App.Title, "Prefs", "Y", FormPos.Top)
     Call SaveSetting(App.Title, "Prefs", "H", FormPos.Height)
     Call SaveSetting(App.Title, "Prefs", "W", FormPos.Width)
End If

End Function

Public Function SavePreferences()

    ' Save User Preferences

    With frmPrefs
     Call SaveSetting(App.Title, "Prefs", "LoadWinPos", ConvertIntegerToBoolean(.Check1))
     Call SaveSetting(App.Title, "Prefs", "AutoUpdate", ConvertIntegerToBoolean(.Check2))
     Call SaveSetting(App.Title, "Prefs", "ShowAddyBar", ConvertIntegerToBoolean(.Check3))
     Call SaveSetting(App.Title, "Prefs", "ShowStatusBar", ConvertIntegerToBoolean(.Check4))
     Call SaveSetting(App.Title, "Prefs", "ShowSystemTray", ConvertIntegerToBoolean(.Check5))
     
     Call SaveSetting(App.Title, "Prefs", "UseHovers", ConvertIntegerToBoolean(.Check6))
     
     Call SaveSetting(App.Title, "Prefs", "DeleteTemps", ConvertIntegerToBoolean(.Check7))
     Call SaveSetting(App.Title, "Prefs", "HideErrors", ConvertIntegerToBoolean(.Check8))
     Call SaveSetting(App.Title, "Prefs", "ForceNull60Sec", ConvertIntegerToBoolean(.Check9))

     Call SaveSetting(App.Title, "Prefs", "LockLists", ConvertIntegerToBoolean(.Check10))
     Call SaveSetting(App.Title, "Prefs", "KeepOnTop", ConvertIntegerToBoolean(.Check11))
     
     Call SaveSetting(App.Title, "Prefs", "CheckModified", ConvertIntegerToBoolean(.Check13))
    End With

End Function
