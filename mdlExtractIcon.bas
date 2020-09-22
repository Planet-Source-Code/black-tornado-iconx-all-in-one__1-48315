Attribute VB_Name = "mdlExtractIcon"
Option Explicit
'Declare the API Calls
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
'Declare the Variables used
Public IconIndex As Integer
Public IconSize As Integer
Public IconFile As String
Public LargeIco() As Long
Public SmallIco() As Long
Public IconsCount As Integer
Public Const DI_NORMAL = 3
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
'Public Const DI_NORMAL = DI_MASK Or DI_IMAGE

Public Function ExtractIconFile(IconIndex As String, IconFile As String, OutPic As PictureBox, IconSize As Integer)
'If Len(IconFile) < 3 Then MsgBox GetStrLang("NoIconSelected"), vbCritical: Exit Sub
ReDim LargeIco(IconIndex)
ReDim SmallIco(IconIndex)
Call ExtractIconEx(IconFile, IconIndex, LargeIco(IconIndex), SmallIco(IconIndex), 1)
With OutPic
    .Picture = LoadPicture("")
    .AutoRedraw = True
     Call DrawIconEx(.hDC, 0, 0, LargeIco(IconIndex), IconSize, IconSize, 0, 0, DI_NORMAL)
    .Refresh
End With
End Function
Public Function CountIcons(FileName As String) As Integer
CountIcons = ExtractIconEx(FileName, -1, 0, 0, 0)
End Function

