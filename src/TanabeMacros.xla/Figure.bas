Attribute VB_Name = "Figure"

Private Const app As String = "excel"



Sub SaveTrimingInfo()
  Dim s As shape
  Const trim As String = "Triming"
  If TypeName(Selection) = "Picture" Then
    Set s = Selection.ShapeRange(1)
    With s.PictureFormat
      SaveSetting app, trim, "right", CStr(.CropRight)
      SaveSetting app, trim, "top", CStr(.CropTop)
      SaveSetting app, trim, "bottom", CStr(.CropBottom)
      SaveSetting app, trim, "left", CStr(.CropLeft)
    End With
  End If
End Sub


Sub LoadTrimingInfo()
  Const trim As String = "Triming"
  Dim s As shape
  If TypeName(Selection) = "Picture" Then
    For Each s In Selection.ShapeRange
      With s.PictureFormat
        .CropRight = CSng(GetSetting(app, trim, "right", "0.0"))
        .CropBottom = CSng(GetSetting(app, trim, "bottom", "0.0"))
        .CropTop = CSng(GetSetting(app, trim, "top", "0.0"))
        .CropLeft = CSng(GetSetting(app, trim, "left", "0.0"))
      End With
    Next s
  End If
End Sub

