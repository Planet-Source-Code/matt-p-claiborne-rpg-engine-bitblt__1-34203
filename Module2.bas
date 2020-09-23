Attribute VB_Name = "Module2"


Public Function GenerateDC(ByVal FileName As String, BitmapProperties As BITMAP) As Long
    Dim DC As Long
    Dim hBitmap As Long
    DC = CreateCompatibleDC(0)



    hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)


    If hBitmap = 0 Then
        DeleteDC DC
        GenerateDC = 0
        MsgBox "Error Loading File!"
        Exit Function
    End If
    GetObjectAPI hBitmap, Len(BitmapProperties), BitmapProperties


    SelectObject DC, hBitmap
        GenerateDC = DC
        DeleteObject hBitmap
End Function

Public Function NewDC(hdcScreen As Long, HorRes As Long, VerRes As Long) As Long
  Dim hdcCompatible As Long
  Dim hbmScreen As Long
  hdcCompatible = CreateCompatibleDC(hdcScreen)
  hbmScreen = CreateCompatibleBitmap(hdcScreen, HorRes, VerRes)
  If SelectObject(hdcCompatible, hbmScreen) = vbNull Then
    NewDC = 1
  Else
    NewDC = hdcCompatible
  End If
  DeleteDC (hdcScreen)
End Function


