Attribute VB_Name = "mdlTransperantFormAndPictureMask"
' FormSkinner
' This program allows you to skin a form by a picture with masking it
' Please vote for me @ a1vbcode.com
Option Explicit
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Function MakeRegion(picSkin As PictureBox, Optional ByVal TransCol = "THE_FIRST_PIXEL") As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region.
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable enter the colors as HEX like &H000000FF&
    ' or RGB(100,100,100) or vbBlack, vbWhite etc...
    If TransCol = "THE_FIRST_PIXEL" Then
    TransparentColor = GetPixel(picSkin.hDC, 0, 0)
    GoTo ProcessNow:
    End If
    TransparentColor = TransCol
ProcessNow:
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function
Function SkinMe(sknFrm As Form, TempPic As PictureBox, Optional TransperantColor = "THE_FIRST_PIXEL")
    Dim WindowRegion As Long
    TempPic.ScaleMode = vbPixels
    TempPic.AutoRedraw = True
    TempPic.AutoSize = True
    TempPic.BorderStyle = vbBSNone
    sknFrm.BorderStyle = vbBSNone
    Set TempPic.Picture = sknFrm.Picture  ' LoadPicture("E:\Black Tornado\VB Projects\FormSkinner" & "\skin.bmp")   ' loads the skin
    sknFrm.Width = TempPic.Width * 15
    sknFrm.Height = TempPic.Height * 15
    WindowRegion = MakeRegion(TempPic, TransperantColor)
    SetWindowRgn sknFrm.hwnd, WindowRegion, True
    sknFrm.Refresh
    ' sknFrm.Picture = TempPic.Picture
sknFrm.Refresh
End Function

Function SkinPic(sknFrm As PictureBox, TempPic As PictureBox, Optional TransperantColor = "THE_FIRST_PIXEL")
    Dim WindowRegion As Long
    TempPic.ScaleMode = vbPixels
    TempPic.AutoRedraw = True
    TempPic.AutoSize = True
    TempPic.BorderStyle = vbBSNone
    sknFrm.BorderStyle = vbBSNone
    Set TempPic.Picture = sknFrm.Picture  ' LoadPicture("E:\Black Tornado\VB Projects\FormSkinner" & "\skin.bmp")   ' loads the skin
    sknFrm.Width = TempPic.Width * 15
    sknFrm.Height = TempPic.Height * 15
    WindowRegion = MakeRegion(TempPic, TransperantColor)
    SetWindowRgn sknFrm.hwnd, WindowRegion, True
    sknFrm.Refresh
    ' sknFrm.Picture = TempPic.Picture
sknFrm.Refresh
End Function


