VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   2970
   ClientTop       =   1245
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MouseIcon       =   "Rotated fonts.frx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oX, oY
Dim iType
Public sDegree
Dim rdc, rdc2
Dim sX, sY
Dim sStr As String
Dim strs As New Collection
Dim iCur As Integer
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Sub RotaFont()
  On Error GoTo EndPoint
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
  Dim FONTSIZE As Integer
  FONTSIZE = Val(36) 'current virtual font size
  F.lfEscapement = 10 * Val(sDegree) 'rotation angle, in tenths
  FontName = "Arial Black" + Chr$(0) 'null terminated
  F.lfFacename = FontName 'setting font
  Me.FONTSIZE = FONTSIZE
  F.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
  hFont = CreateFontIndirect(F) 'create the virtual font
  hPrevFont = SelectObject(Me.hdc, hFont)
  CurrentX = sX
  CurrentY = sY
'  FONTSIZE = FONTSIZE + 12
'Generating Color
If iType = 1 Then ForeColor = RndColor Else ForeColor = ColorPercentage(rdc, rdc2, sDegree / 360 * 100)
  Print sStr
  CurrentX = sX + 20
  CurrentY = sY + 20
'  FONTSIZE = FONTSIZE - 12
If Not iType = 1 Then ForeColor = ColorPercentage(rdc2, rdc, sDegree / 360 * 100)
  Print sStr
'  Clean up, restore original font
  hFont = SelectObject(Me.hdc, hPrevFont)
  DeleteObject hFont
  Exit Sub
EndPoint:
  Exit Sub
End Sub
Private Sub Command1_Click()
  Me.Cls
  FontStuff
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End
End Sub
Private Sub Form_Load()
Show
Me.MousePointer = 99 'hide mouse cursor
sDegree = Int((Rnd * 361))
rdc = RndColor
rdc2 = RndColor
Randomize
sX = (Rnd * (Me.Width - TextWidth(sStr))) + (TextWidth(sStr) / 2)
sY = (Rnd * (Me.Height - TextWidth(sStr))) + (TextWidth(sStr) / 2)
iType = 1 + Int(Rnd * 2)
Dim MyStr As String
MyStr = GetSetting("TechnosoftScreen", "MarcoRotate", "SaverText")
If Trim(MyStr) = "" Then MyStr = "Marco Samy Nasif"
GetAllAB MyStr, " ", " ", strs
If iCur = strs.Count Then iCur = 0
iCur = iCur + 1
sStr = strs.Item(iCur)
DoEvents
Do
If sDegree = 360 Then
sDegree = 0
If iCur = strs.Count Then
iCur = 0
'bluring the screen, is just an effect
Dim wX, wY, TmpStd As StdPicture
Set TmpStd = Image
AutoRedraw = False
For wX = 1 To Me.Width Step Width / 100
For wY = 1 To Me.Height Step Height / 100
'remeber the things drawan on any object is saved in the Image
'so we will use the Image
PaintPicture TmpStd, wX, wY, (Width / 100), (Height / 100), wX - 150, wY - 150, (Width / 100) - 300, (Height / 100) - 300
Next: Next
AutoRedraw = True
'AutoCls
End If
iCur = iCur + 1
sStr = strs.Item(iCur)
Randomize
sX = (Rnd * (Me.Width - TextWidth(sStr))) + (TextWidth(sStr) / 2)
Randomize
sY = (Rnd * (Me.Height - TextWidth(sStr))) + (TextWidth(sStr) / 2)
rdc = RndColor
rdc2 = RndColor
Randomize
iType = 1 + Int(Rnd * 2)
End If
sDegree = sDegree + 1
''Cls
RotaFont
DoEvents
Loop
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If (oX = 0) And (oY = 0) Then oX = x: oY = Y
If (Abs(x - oX) > 500) Or (Abs(Y - oY) > 500) Then End 'when it moves long, we end
'If Started = True Then End
End Sub
