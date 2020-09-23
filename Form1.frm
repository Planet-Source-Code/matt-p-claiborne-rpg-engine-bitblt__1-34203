VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7320
         Top             =   1080
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DELAY_TIME = 20
'set height width as big as your map
Const MAP_WIDTH = 40 * 32
Const MAP_HEIGHT = 40 * 32

Dim Char As MoB
Dim Running As Boolean
Dim CharDC As Long
Dim CharSpriteDC As Long
Dim TempDC As Long
Dim bmpProperties As BITMAP
Dim FPS As Long
Dim LastCheck As Long
Dim AnimCount As Single
Dim Map(MAP_WIDTH / 32, MAP_HEIGHT / 32, 1) As Integer
Dim PicDC As Long
Dim StageDC As Long
Dim BmpOld As Long
Dim TileSet As Long
Dim Retval As Long
Dim Xt As Integer
Dim Yt As Integer
Dim Xz As Integer
Dim Yz As Integer
Dim PosX As Long
Dim PosY As Long
Dim Xr As Long
Dim Yr As Long



Private Sub Form_Load()

Char.x = 320
Char.y = 240
Char.OldX = 320
Char.OldY = 240
Char.Direction = 1
Running = True
Me.Visible = True
Char.Movement = 3


MainLoop
End Sub

Function MainLoop()
'load bitmaps into memory DC's
CharDC = GenerateDC(App.Path & "\cecil.bmp", bmpProperties)
CharSpriteDC = GenerateDC(App.Path & "\cecilsprite.bmp", bmpProperties)
TempDC = GenerateDC(App.Path & "\Image2.bmp", bmpProperties)
TileSet = GenerateDC(App.Path & "\tiles1.bmp", bmpProperties)




'make a *buffer* to *store* stuff...
StageDC = NewDC(Picture1.hdc, 640, 480)
'StageDC = CreateCompatibleDC(Picture1.hdc)
'Retval = SelectObject(StageDC, Picture1)

'Randomize GetTickCount
'For Xt = 0 To MAP_WIDTH / 32
'For Yt = 0 To MAP_HEIGHT / 32

'Map(Xt, Yt, 0) = (Int(Rnd * 1.05) + 1) * 32

'Next
'Next
Dim T As String

Open App.Path & "/map.txt" For Input As 1
For Yt = 0 To MAP_HEIGHT / 32 - 1
Input #1, T
For Xt = 0 To (MAP_WIDTH / 32) - 1
Map(Xt, Yt, 0) = Mid(T, Xt + 1, 1) * 32
If Map(Xt, Yt, 0) = 0 Then
Map(Xt, Yt, 1) = 1
Else
Map(Xt, Yt, 1) = 0
End If
Next
Next






Do While Running = True

'limit frames to 50 fps
If Not GetTickCount - LastCheck >= DELAY_TIME Then
GoTo Nada
End If
LastCheck = GetTickCount


MoveChar
DrawStage



'copy buffer to screen

BitBlt Picture1.hdc, 0, 0, Picture1.Width, Picture1.Height, StageDC, 0, 0, SRCCOPY


FPS = FPS + 1
Nada:



DoEvents
Loop

DeleteDC (CharDC)
DeleteDC (TempDC)
DeleteDC (StageDC)
DeleteDC (TileSet)
DeleteDC (CharSpriteDC)

End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
DeleteDC (CharDC)
DeleteDC (TempDC)
DeleteDC (StageDC)
DeleteDC (CharSpriteDC)
End
End Sub

Private Sub Timer1_Timer()
Form1.Caption = FPS & " - FPS"
FPS = 0
End Sub


Private Sub MoveChar()
Dim Xd As Long
Dim Yd As Long

Char.Movement = 4
Char.Moved = False

If GetAsyncKeyState(vbKeyA) Then
Char.Movement = 15
End If

If GetAsyncKeyState(vbKeyS) Then
Char.Movement = 1
End If

If GetAsyncKeyState(GM_ESCAPE) Then
Running = False
End If



If GetAsyncKeyState(GM_LEFT) Then
Char.OldX = Char.x
Char.OldY = Char.y
Char.x = Char.x - Char.Movement
Char.Direction = 3
Char.Moved = True
GoTo Moved
End If

If GetAsyncKeyState(GM_UP) Then
Char.OldY = Char.y
Char.OldX = Char.x
Char.y = Char.y - Char.Movement
Char.Direction = 0
Char.Moved = True
GoTo Moved
End If

If GetAsyncKeyState(GM_RIGHT) Then
Char.OldX = Char.x
Char.OldY = Char.y
Char.x = Char.x + Char.Movement
Char.Direction = 1
Char.Moved = True
GoTo Moved
End If

If GetAsyncKeyState(GM_DOWN) Then
Char.OldY = Char.y
Char.OldX = Char.x
Char.y = Char.y + Char.Movement
Char.Direction = 2
Char.Moved = True
GoTo Moved
End If


Moved:

If Char.x < 0 Then
Char.x = 0
ElseIf Char.x + 32 > MAP_WIDTH Then
Char.x = MAP_WIDTH - 32
End If

If Char.y < 0 Then
Char.y = 0
ElseIf Char.y + 32 > MAP_HEIGHT Then
Char.y = MAP_HEIGHT - 32
End If


Xd = Int(Char.x / 32)
Yd = Int(Char.y / 32)

If Int(Char.x / 32) = Char.x / 32 Then
Xz = 0
Else: Xz = 1
End If

If Int(Char.y / 32) = Char.y / 32 Then
Yz = 0
Else: Yz = 1
End If

For Xr = 0 To Xz
For Yr = 0 To Yz

If Map(Xd + Xr, Yd + Yr, 1) = 1 Then

    If Char.Direction = 1 Then
        Char.x = (Xd + Xr - 1) * 32
        Char.y = Char.OldY
    End If
    
    If Char.Direction = 3 Then
        Char.x = (Xd + Xr + 1) * 32
        Char.y = Char.OldY
    End If
    
    If Char.Direction = 2 Then
        Char.x = Char.OldX
        Char.y = (Yd + Yr - 1) * 32
    End If
    
    If Char.Direction = 0 Then
        Char.x = Char.OldX
        Char.y = (Yd + Yr + 1) * 32
    End If

End If

Next
Next


If Char.Moved = True Then
AnimCount = AnimCount + Char.Movement / 20
If AnimCount > 2 Then AnimCount = 0
Else
AnimCount = 0
End If


If Char.x < 320 Then
PosX = -(320 - Char.x)
ElseIf Char.x > MAP_WIDTH - 320 Then
PosX = (Char.x - MAP_WIDTH) + 320
Else
PosX = 0
End If



If Char.y < 240 Then
PosY = -(240 - Char.y)
ElseIf Char.y > MAP_HEIGHT - 240 Then
PosY = (Char.y - MAP_HEIGHT) + 240
Else
PosY = 0
End If






End Sub

Private Sub DrawStage()


Xr = (Char.x - 320)
Yr = (Char.y - 240)

If Xr < 0 Then Xr = 0
If Yr < 0 Then Yr = 0

If Xr + 640 > MAP_WIDTH Then Xr = MAP_WIDTH - 640
If Yr + 480 > MAP_HEIGHT Then Yr = MAP_HEIGHT - 480

Xz = (Xr) Mod 32
Yz = (Yr) Mod 32

For Xt = 0 To 640 Step 32
For Yt = 0 To 480 Step 32

BitBlt StageDC, Xt - (Xz), Yt - (Yz), 32, 32, TileSet, Map(Int((Xr + Xt) / 32), Int((Yr + Yt) / 32), 0), 0, SRCCOPY

Next
Next

BitBlt StageDC, 320 + PosX, 240 + PosY, 32, 32, CharSpriteDC, CInt(AnimCount) * 32, Char.Direction * 32, SRCAND
BitBlt StageDC, 320 + PosX, 240 + PosY, 32, 32, CharDC, CInt(AnimCount) * 32, Char.Direction * 32, SRCPAINT


End Sub
