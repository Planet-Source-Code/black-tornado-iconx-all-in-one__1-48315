VERSION 5.00
Begin VB.UserControl IconX 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   InvisibleAtRuntime=   -1  'True
   Picture         =   "IconX Control.ctx":0000
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   71
   ToolboxBitmap   =   "IconX Control.ctx":0ECA
   Begin VB.PictureBox picTemp 
      BackColor       =   &H00404080&
      Height          =   2055
      Left            =   2880
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "IconX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------------------
'         IconX ActiveX control by Black Tornado
' --------------------------------------------------------
' This code was written on 2/9/2003 by Black Tornado
' This control gives you access to many functions easily
' Wich will help you build an easy Icon application
' --------------------------------------------------------
' You are FREE to use this code, please show my name if you
' want because this code took me many days to wrote and test
' If you found any bug or you want to discuss something with
' me, then feel free to contact me by e-mail:
'          My e-mail:     btsoft@burntmail.com
'         My website: http://www.blacktornado.cjb.net
' --------------------------------------------------------
'         Black Tornado Software is always the best
'            OUR LOGO: NOTHING IS IMPOSSIBLE
' ---------------------------------------------------------
'        Please vote for me if you like this code
'                 Thank you very much
' ---------------------------------------------------------
'                 WAIT MY FUTURE PROGRAMS
' ---------------------------------------------------------
' Note1: I used SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
' in many functions, this command will refresh the system by nofitying
' windows that system resources has been changed.
' ---------------------------------------------------------
' Note2: When trying to save an icon or skinning a form, the transperant
' color should be a normal hex color, not vbButtonFace or other system
' colors, if you set the color to system colors then the function will
' not work good. So that I recommended that you set the transperant color
' to a normal color.
' ---------------------------------------------------------

Option Explicit
Private TempVal As String                 ' Temp Value, very useful
Private TempVal2 As String                 ' Temp Value, very useful
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETTINGCHANGE = WM_WININICHANGE
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type RBGColor
    R As Byte
    G As Byte
    B As Byte
End Type
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
  
Public Function About()
frmAbout.Show vbModal
End Function
Public Function SavePicToIcon(PicturehDC As Long, TransperantColor As Long, TargetFileName As String, SaveBits As IconSaveTypes) As Boolean
ExportPicToIcon PicturehDC, TransperantColor, TargetFileName, SaveBits
End Function

Public Function ExtractIcon(IconFileName As String, IconIndex As String, OutputPicture As PictureBox, hIconSize As Integer)
Call ExtractIconFile(IconIndex, IconFileName, OutputPicture, hIconSize)
End Function

Public Function CountIconsInFile(ByVal IconFileName As String) As Long
CountIconsInFile = CountIcons(IconFileName)
End Function
Public Function ChangeFolderIcon(TargetFolder As String, PictFile As String, FolderComments As String, CopyIcon As Boolean, PictIndex As String, IconSaveType As IconSaveTypes, Optional ByVal ErrorMsg = "This command only supports fixed drives") As Boolean
' The write of ini should be like this:
' [.ShellClassInfo]
' PictFile  = >>   Your Icon File here     <<
' PictIndex = >>   Your Icon Index here    <<
' InfoTip   = >> Your folder comments here <<
' The following code should do that
TempVal = Replace(TargetFolder, Chr(34), "")   ' Some times when the folder is too long
                                             ' Command$ returns "COMMAND", so that we
                                             ' must remove these " " to get the pure
' Now we must check if the folder is in a CD-ROM or FIXED DRIVE or REMOVABLE DRIVE
' to do that we must use the WinAPI function, GetDriveType
' It will return 3 for fixed drive, 2 for removale drive, 5 for CD-ROM
If GetDriveType(Left$(TempVal, 3)) <> 3 Then
MsgBox ErrorMsg, vbCritical, "IconX - ActiveX Control by Black Tornado"
ChangeFolderIcon = False
Exit Function
End If
' Now every thing is OK so let us begin work
' Now write the sections into the INI file
' INI file name for a folder is Desktop.ini
On Error Resume Next
' Attribute the old Desktop.ini file to ARCHIVE
SetAttr TargetFolder & "\Desktop.ini", vbArchive
Kill TargetFolder & "\Desktop.ini"
' If the icon will be copied then the name of the icon file will change to IconX.ico
' And that is done after that current picture/icon is exported to a seperate file
If CopyIcon = True Then
If Len(PictFile) <= 1 Then GoTo 60
SetAttr TargetFolder & "\IconX.ico", vbArchive     ' Remove the old
Kill TargetFolder & "\IconX.ico"                   ' icons from the dir
WriteIniFile ".ShellClassInfo", "IconFile", "IconX.ico", TargetFolder & "\Desktop.ini"
WriteIniFile ".ShellClassInfo", "IconIndex", "0", TargetFolder & "\Desktop.ini"
WriteIniFile ".ShellClassInfo", "InfoTip", FolderComments, TargetFolder & "\Desktop.ini"
If UCase(Right(PictFile, 4)) = ".ICO" Then FileCopy PictFile, TargetFolder & "\IconX.ico": GoTo ProcessNow
UserControl.picTemp.Picture = LoadPicture("")
Call ExtractIconFile(Int(PictIndex), PictFile, UserControl.picTemp, 32)
Call SavePicToIcon(picTemp.hDC, picTemp.BackColor, TargetFolder & "\IconX.ico", IconSaveType)
SetAttr TargetFolder & "\IconX.ico", vbHidden
GoTo ProcessNow
End If
If Len(PictFile) <= 1 Then GoTo 60
WriteIniFile ".ShellClassInfo", "IconFile", PictFile, TargetFolder & "\Desktop.ini"
WriteIniFile ".ShellClassInfo", "IconIndex", PictIndex, TargetFolder & "\Desktop.ini"
60 WriteIniFile ".ShellClassInfo", "InfoTip", FolderComments, TargetFolder & "\Desktop.ini"
  ' Now after finishing we must set the folder attributes to READ-ONLY, this is
  ' So Important if we want the folder to display the icon
  ' We will also set the attribute of the file Desktop.ini to hidden+system
ProcessNow:
SetAttr TargetFolder & "\Desktop.ini", vbHidden + vbSystem
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(TargetFolder)
    f.Attributes = 1
    Call RebuildIconCache
End Function

Public Function ChangeDriveIcon(DriveLetter As String, Icon As String, RunProgram As String, CopyIcon As Boolean, IconSaveType As IconSaveTypes, Optional ByVal PictIndex = 0, Optional ByVal ErrorMsg = "This command only supports fixed drives") As Boolean
' The name of the inf file is Autorun.inf
' The write of inf should be like this:
' [AutoRun]
' Icon      = >>   Your Icon File here  , IconIndex   <<
' Open      = >>      Your autorun program here       <<
' The following code should do that
' Now we must check if the drive is a CD-ROM Drive or FIXED DRIVE or REMOVABLE DRIVE
' to do that we must use the WinAPI function, GetDriveType
' It will return 3 for fixed drive, 2 for removale drive, 5 for CD-ROM
' I discovered that when we set the OPEN to nothing, explorer will ask you about the file
' So, If we don't want an autorun program, we must not write the key 'OPEN'
TempVal = Left$(DriveLetter, 2)
If GetDriveType(TempVal & "\") <> 3 Then
MsgBox ErrorMsg, vbCritical, "IconX - ActiveX Control by Black Tornado"
ChangeDriveIcon = False
Exit Function
End If
' Now every thing is OK so let us begin work
' Now write the sections into the INF file
' INF file name for a drive is Autorun.inf
On Error Resume Next
' Attribute the old Desktop.ini file to ARCHIVE
SetAttr TempVal & "\Autorun.inf", vbArchive
Kill TempVal & "\Autorun.inf"
' If the icon will be copied then the name of the icon file will change to IconX.ico
' And that is done after that current picture/icon is exported to a seperate file
If Len(Icon) <= 1 And Len(RunProgram) <= 1 Then RemoveDriveIcon DriveLetter: GoTo DoIt
If CopyIcon = True Then
If UCase(Right(Icon, 4)) = ".ICO" Then FileCopy Icon, DriveLetter & "IconX.ico": GoTo DoIt
If Len(RunProgram) <= 1 Then GoTo 10
SetAttr TempVal & "\IconX.ico", vbArchive
Kill TempVal & "\IconX.ico" ' Kill the existing file
WriteIniFile "AutoRun", "Open", RunProgram, TempVal & "\Autorun.inf"
10 WriteIniFile "AutoRun", "Icon", "IconX.ico", TempVal & "\Autorun.inf"
picTemp.Picture = LoadPicture("")              ' Clear the picture of the control
Call ExtractIcon(Icon, PictIndex & " ", UserControl.picTemp, 32) ' We entered PictIndex & " " to make the PC think that
                                                                 ' the current value is string
Call SavePicToIcon(picTemp.hDC, picTemp.BackColor, DriveLetter & "\IconX.ico", IconSaveType)
SetAttr TempVal & "\IconX.ico", vbHidden
GoTo DoIt
End If
If Len(RunProgram) <= 1 Then GoTo 20 'Return to old QuickBasic style
WriteIniFile "AutoRun", "Open", RunProgram, TempVal & "\Autorun.inf"
20 WriteIniFile "AutoRun", "Icon", Icon & "," & IconIndex, TempVal & "\Autorun.inf"
DoIt:
' We will set the attribute of the file Autorun.inf to hidden+system
SetAttr DriveLetter & "\Autorun.inf", vbHidden + vbSystem
' Set the function to TRUE because operation is OK
ChangeDriveIcon = True
  ' Now the final step is to rebuild the icon cache
  ' Wich is also called REFRESH EXPLORER ICONS
Call RebuildIconCache
End Function

Public Function RebuildIconCache()
TempVal = GetString(HKEY_CURRENT_USER, "Control Panel\desktop\WindowMetrics", "Shell icon size")
SaveString HKEY_CURRENT_USER, "Control Panel\desktop\WindowMetrics", "Shell icon size", TempVal + 1
SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
SaveString HKEY_CURRENT_USER, "Control Panel\desktop\WindowMetrics", "Shell icon size", TempVal
SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Function

Public Function RemoveFolderIcon(FolderToRemove As String) As Boolean
' Now we will remove the " " from the folder
TempVal = Replace(FolderToRemove, Chr(34), "")
If GetDriveType(Left$(Replace(FolderToRemove, Chr(34), ""), 3)) <> 3 Then _
                RemoveFolderIcon = False: Exit Function
On Error Resume Next       ' Resume operation even if the folder has no settings
SetAttr TempVal & "\Desktop.ini", vbArchive
SetAttr TempVal & "\IconX.ico", vbArchive
Kill TempVal & "\Desktop.ini"
Kill TempVal & "\IconX.ico"
RemoveFolderIcon = True    ' Folder settings removed
Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(FolderToRemove)
    f.Attributes = 0
' Now the final step is to rebuild the icon cache
' Wich is also called REFRESH EXPLORER ICONS
Call RebuildIconCache
End Function

Public Function RemoveDriveIcon(DriveToRemove As String) As Boolean
' Now we will remove the " " from the drive
TempVal = Left$(Replace(DriveToRemove, Chr(34), ""), 2)
If GetDriveType(Left$(Replace(DriveToRemove, Chr(34), ""), 3)) <> 3 Then _
                RemoveDriveIcon = False: Exit Function
On Error Resume Next       ' Resume operation even if the Drive has no settings
SetAttr TempVal & "\Autorun.inf", vbArchive
SetAttr TempVal & "\IconX.ico", vbArchive
Kill TempVal & "\Autorun.inf"
Kill TempVal & "\IconX.ico"
RemoveDriveIcon = True     ' Drive settings removed
' Now the final step is to rebuild the icon cache
' Wich is also called REFRESH EXPLORER ICONS
Call RebuildIconCache
End Function


Public Function GetFolderIconFile(FolderName As String) As String
GetFolderIconFile = ReadIniFile(FolderName & "\Desktop.ini", ".ShellClassInfo", "IconFile")
End Function

Public Function GetFolderIconIndex(FolderName As String) As String
GetFolderIconIndex = ReadIniFile(FolderName & "\Desktop.ini", ".ShellClassInfo", "IconIndex")
End Function

Public Function GetFolderComment(FolderName As String) As String
GetFolderComment = ReadIniFile(FolderName & "\Desktop.ini", ".ShellClassInfo", "InfoTip")
End Function

Public Function GetDriveIconIndex(DriveLetter As String) As String
Dim WhereIsDot          ' The , location is stored in this item
TempVal2 = ReadIniFile(DriveLetter & "Autorun.inf", "Autorun", "Icon")
If Len(TempVal2) = 0 Then GetDriveIconIndex = "": Exit Function
WhereIsDot = InStr(TempVal2, ",")
TempVal = Mid$(TempVal2, WhereIsDot + 1, Len(TempVal2) - WhereIsDot)
If TempVal = TempVal2 Then GetDriveIconIndex = "": Exit Function
GetDriveIconIndex = TempVal
End Function

Public Function GetDriveIconFileName(DriveLetter As String) As String
Dim WhereIsDot          ' The , location is stored in this item
TempVal2 = ReadIniFile(DriveLetter & "Autorun.inf", "Autorun", "Icon")
If InStr(TempVal2, ",") = 0 Then GetDriveIconFileName = TempVal2: Exit Function
If Len(TempVal2) <= 1 Then GetDriveIconFileName = "": Exit Function
WhereIsDot = InStr(TempVal2, ",")
If WhereIsDot = 0 Then WhereIsDot = Len(TempVal2)
TempVal = Left$(TempVal2, WhereIsDot - 1)
'If TempVal = TempVal2 Then GetDriveIconIndex = "": Exit Function
GetDriveIconFileName = TempVal
End Function

Public Function GetDriveAutoRun(DriveLetter As String) As String
TempVal2 = ReadIniFile(DriveLetter & "Autorun.inf", "Autorun", "Open")
If Len(TempVal2) <= 1 Then GetDriveAutoRun = "": Exit Function
GetDriveAutoRun = TempVal2
End Function

Public Function GetIcon(btFolder As String, OutPic As PictureBox)
'On Error GoTo errr
    Dim hIconHandle     As Long
    Dim TargetPath As String
    TargetPath = btFolder
    
    hIconHandle = ExtractAssociatedIcon(OutPic.hwnd, TargetPath, 2)
    
    'if call is success then icon handle will be obtained
     
    OutPic.Cls
    OutPic.Picture = LoadPicture("")
    If hIconHandle Then
    OutPic.ScaleMode = vbPixels
    DrawIconEx OutPic.hDC, 0, 0, hIconHandle, 32, 32, 0, 0, DI_NORMAL
    DestroyIcon hIconHandle
    Else
        MsgBox "Unable to extract the Icon from the folder or drive.", vbCritical
    End If
End Function

Private Sub UserControl_Paint()
UserControl.Width = 48 * 15
UserControl.Height = 48 * 15
End Sub

Public Function TransForm(FrmToTrans As Form, Optional TransColor = "THE_FIRST_PIXEL")
Call SkinMe(FrmToTrans, UserControl.picTemp, TransColor)
End Function

Public Function ChangeColor(SourcehDC As Long, DesthDc As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long, OldColor As Long, NewColor As Long)
Call ChangePixels(SourcehDC, StartX, StartY, EndX, EndY, OldColor, NewColor, DesthDc)
End Function

Public Function GetADriveType(Drv As String) As String
Select Case GetDriveType(Drv):
Case 0:
GetADriveType = "Unkown Drive"
Case 1:
GetADriveType = "Root directory doesn't exists"
Case 2:
GetADriveType = "Removable Drive"
Case 3:
GetADriveType = "Fixed Drive"
Case 4:
GetADriveType = "Network Drive"
Case 5:
GetADriveType = "CD-ROM Drive"
Case 6:
GetADriveType = "Ram Drive"
End Select
End Function

Public Function WriteIni(IniFile As String, Section As String, Key As String, Value As String)
Call WriteIniFile(Section, Key, Value, IniFile)
End Function

Public Function ReadIni(IniFile As String, Section As String, Key As String) As String
ReadIni = ReadIniFile(IniFile, Section, Key)
End Function

    
Private Function Win2RGB(wincolor As Long) As RBGColor
Dim HexColor As String

' this must be 6 chars long !
HexColor = Right("000000" & Hex(wincolor), 6)

' color format: 0x00bbggrr
With Win2RGB
    .R = CInt("&H" & (Right(HexColor, 2)))
    .B = CInt("&H" & (Left(HexColor, 2)))
    .G = CInt("&H" & (Mid(HexColor, 3, 2)))
End With
End Function

Public Function ConvertHexColorToRGB(HexColor As Long) As String
Dim RedGreenBlue As RBGColor

' pass the BG color of Me to a function and get
' the RGB value
RedGreenBlue = Win2RGB(HexColor)

' return the RGB values to the function
ConvertHexColorToRGB = "Red(" & RedGreenBlue.R & "), Green(" & _
                  RedGreenBlue.G & "), Blue(" & _
                 RedGreenBlue.B & ")"
End Function

Public Function RestoreIconThumbnail()
SaveString HKEY_CLASSES_ROOT, ".ico", "", "icofile"
SaveString HKEY_CLASSES_ROOT, "icofile", "", "Icon"
SaveString HKEY_CLASSES_ROOT, "icofile\DefaultIcon", "", "%1"
SaveString HKEY_CLASSES_ROOT, "icofile\shell", "", "open"
SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
RebuildIconCache
End Function

Public Function EnableFixedDrivesAutorun()
SaveDWORD HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", "00000091"
SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
RebuildIconCache
End Function

Public Function EnableBMPThumbnail()
SaveString HKEY_CLASSES_ROOT, "Paint.Picture\DefaultIcon", "", "%1"
SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
RebuildIconCache
End Function

Public Function MoveForm(FormToMove As Form)
    ReleaseCapture
    SendMessage FormToMove.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function
