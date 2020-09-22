VERSION 5.00
Begin VB.Form frmTestOnly 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   Icon            =   "IconX-Test.frx":0000
   LinkTopic       =   "Iconizer"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "IconX-Test.frx":0ECA
   ScaleHeight     =   8940
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Click this button to exit this program"
      Top             =   7080
      Width           =   1815
   End
   Begin VB.PictureBox picIconX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00DD8911&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   480
      Left            =   9480
      Picture         =   "IconX-Test.frx":16080E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      ToolTipText     =   "A picture of a cube, when you select 'Export picture to icon' this picture will be exported as an example of IconX functions."
      Top             =   5040
      Width           =   480
   End
   Begin VB.CommandButton cmdAboutIconX 
      BackColor       =   &H00CAD8DF&
      Caption         =   "About IconX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Click this button to view some informations about the author and the project"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnabledFDAutorun 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Enable Autorun for Fixed Drives"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click this button to enable 'Autorun' feature for fixed drives."
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdRestoreIconThumbnail 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Restore Icon Thumbnail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sometimes explorer doesn't display the icon file thumbnail, click this button to restore the thumbnails"
      Top             =   6000
      Width           =   3735
   End
   Begin VB.CommandButton cmdExportPictureToIcon1 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Export picture to 1 bit icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click here to export the 'Cube' picture to 1 bit icon file (Black and White)"
      Top             =   7440
      Width           =   3735
   End
   Begin VB.CommandButton cmdExportPictureToIcon4 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Export picture to 4 bit icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Click here to export the 'Cube' picture to 4 bit icon file"
      Top             =   7440
      Width           =   3735
   End
   Begin VB.CommandButton cmdExportPictureToIcon8 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Export picture to 8 bit icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click here to export the 'Cube' picture to 8 bit icon file"
      Top             =   7920
      Width           =   3735
   End
   Begin IconXControl.IconX IconX_Test 
      Left            =   0
      Top             =   0
      _extentx        =   1270
      _extenty        =   1270
   End
   Begin VB.PictureBox picNew 
      BackColor       =   &H0044332B&
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   29
      ToolTipText     =   "New Picture"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.PictureBox picOld 
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   885
      Left            =   8520
      Picture         =   "IconX-Test.frx":1610D8
      ScaleHeight     =   825
      ScaleWidth      =   2370
      TabIndex        =   28
      ToolTipText     =   "Old Picture"
      Top             =   2880
      Width           =   2430
   End
   Begin VB.CommandButton cmdGetDriveType 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get drive type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click this button to get the type of drive 'C:\'"
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetRGBofHEX 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get the RGB color of HEX color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Click this button to get the RGB color of the CUBE's background"
      Top             =   6960
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get drive/folder icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Gets the icon of a drive and a folder, Shows you as explorer shows"
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetDriveIconIndex 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get drive icon index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click this button to get the icon index of drive 'C:\'"
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetDriveAutorun 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get drive autorun"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click this button to get the autorun file of drive 'C:\'"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetFolderIconIndex 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get folder icon index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Click this button to get the icon index of folder 'C:\IconX_Test'"
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetFolderComment 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get folder comment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Click this button to get the comment of the folder 'C:\IconX_Test'"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton cmdWriteIniFile 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Write ini file"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click this button to write a test INI file"
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton cmdReadIniFile 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Read ini file"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click this button to read strings from test INI file"
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton cmdCountIconsInFile 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Count icons in file"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Click here to count the icons in the Shell32.dll file"
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdExportPictureToIcon24 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Export picture to 24 bit icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Click here to export the 'Cube' picture to 24 bit icon file (True Colors)"
      Top             =   7920
      Width           =   3735
   End
   Begin VB.CommandButton cmdRemoveDriveIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Remove drive icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click this button to remove the icon of Drive 'C:\'"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdRemoveFolderIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Remove folder icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click this button to remove the icon of folder 'C:\IconX_Test'"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetFolderIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get the icon of a folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click this button to get the icon file of folder 'C:\IconX_Test'"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdGetDriveIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Get the icon of drive"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click this button to get the icon file of drive 'C:\'"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdExtractIconFromDll 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Extract icon from DLL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click here to extract an icon from DLL"
      Top             =   6960
      Width           =   3735
   End
   Begin VB.CommandButton cmdChangeFolderIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Change folder icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click this button to change the icon of folder 'C:\IconX_Test'"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdChangeDriveIcon 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Change a drive icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click this button to change the icon of Drive 'C:\'"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton cmdChangeColor 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Change color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click here to change the White color in the first picture with the Red color and put the results in the second picture box"
      Top             =   6480
      Width           =   3735
   End
   Begin VB.CommandButton cmdSkinThisForm 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Skin this form to picture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Click this button to skin the form to the picture"
      Top             =   6480
      Width           =   3735
   End
   Begin VB.CommandButton cmdRebuildIconCache 
      BackColor       =   &H00CAD8DF&
      Caption         =   "Rebuild Icon Cache"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click here to rebuild icon cache and refresh explorer icons"
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IconX  -  By Black Tornado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   8320
      Width           =   11175
   End
End
Attribute VB_Name = "frmTestOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAboutIconX_Click()
IconX_Test.About
End Sub

Private Sub cmdChangeColor_Click()
IconX_Test.ChangeColor picOld.hDC, picNew.hDC, 0, 0, picNew.Width / 15, picNew.Height / 15, vbWhite, vbRed
Message "The red color has been replaced with the white color"
End Sub

Private Sub cmdChangeDriveIcon_Click()
IconX_Test.ChangeDriveIcon "C:\", App.Path & "\IconX.ico", "", False, "7"
Message "Drive 'C:\' icon has been changed sucessfully"
End Sub

Private Sub cmdChangeFolderIcon_Click()
Call IconX_Test.ChangeFolderIcon("C:\IconX_Test", App.Path & "\IconX.ico", "Wow! It worked, thanks for IconX", True, "", SaveTrueColors)
Message "The folder 'C:\IconX_Test' icon has been changed with adding some comment"
End Sub

Private Sub cmdCountIconsInFile_Click()
Message "The number of icons in the dynamic link library file (DLL) 'C:\Windows\System\Shell32.dll' is:'" & IconX_Test.CountIconsInFile("C:\Windows\System\Shell32.dll") & "'"
End Sub

Private Sub cmdExportPictureToIcon_Click()
IconX_Test.SavePicToIcon picOld.hDC, picOld.Point(o, o), "c:\ahmedoo.ico", SaveTrueColors
End Sub

Private Sub cmdEnabledFDAutorun_Click()
IconX_Test.EnableFixedDrivesAutorun
Message "Autorun for FIXED DRIVES is now ENABLED"
End Sub

Private Sub cmdExit_Click()
MsgBox "I hope you enjoyed this program, wait the next version..." + vbCrLf & _
       "Please vote for me if you like this program" + vbCrLf & _
       "Happy Programming...", vbInformation, "IconX By Black Tornado"
       ' Now clear all what we did
       On Error Resume Next
       Dim fso, f
       Set fso = CreateObject("Scripting.FileSystemObject")
       Set f = fso.GetFolder("C:\IconX_Test")
       f.Attributes = 0      ' Set folder attributes to Archive
       Kill "C:\Testini.ini"
       Kill "C:\Test Icon.ico"
       SetAttr "C:\IconX_Test\*.*", vbArchive
       Kill "C:\IconX_Test\*.*"
       RmDir "C:\IconX_Test\"
End
End Sub

Private Sub cmdExportPictureToIcon1_Click()
IconX_Test.SavePicToIcon picIconX.hDC, picIconX.Point(0, 0), "C:\Test Icon.ico", Save1Bit
Message "File has been saved to 'C:\Test Icon.ico'"
End Sub

Private Sub cmdExportPictureToIcon24_Click()
IconX_Test.SavePicToIcon picIconX.hDC, picIconX.Point(0, 0), "C:\Test Icon.ico", SaveTrueColors
Message "File has been saved to 'C:\Test Icon.ico'"
End Sub

Private Sub cmdExportPictureToIcon4_Click()
IconX_Test.SavePicToIcon picIconX.hDC, picIconX.Point(0, 0), "C:\Test Icon.ico", Save4Bits
Message "File has been saved to 'C:\Test Icon.ico'"
End Sub

Private Sub cmdExportPictureToIcon8_Click()
IconX_Test.SavePicToIcon picIconX.hDC, picIconX.Point(0, 0), "C:\Test Icon.ico", Save8Bits
Message "File has been saved to 'C:\Test Icon.ico'"
End Sub

Private Sub cmdExtractIconFromDll_Click()
Call IconX_Test.ExtractIcon("C:\Windows\Calc.exe", 0, picNew, 32)
Message "The picture shown is the icon of 'Calculator' program"
picNew.Picture = LoadPicture("")
End Sub

Private Sub cmdGetDriveAutorun_Click()
Message "The autorun program of drive 'C:\'is:'" & IconX_Test.GetDriveAutoRun("c:\") & "'"
End Sub

Private Sub cmdGetDriveIcon_Click()
Message "The icon file of drive 'C:\' is:'" & IconX_Test.GetDriveIconFileName("C:\") & "'"
End Sub

Private Sub cmdGetDriveIconIndex_Click()
Message "The icon index of drive 'C:\' is:'" & IconX_Test.GetDriveIconIndex("C:\") & "'"
End Sub

Private Sub cmdGetDriveType_Click()
Message "The type of drive 'C:\' is:'" & IconX_Test.GetADriveType("C:\") & "'"
End Sub

Private Sub cmdGetFolderComment_Click()
Message "The comment of folder 'C:\IconX_Test' is:'" & IconX_Test.GetFolderComment("C:\IconX_Test") & "'"
End Sub

Private Sub cmdGetFolderIcon_Click()
Message "The icon file of folder 'C:\IconX_Test' is:'" & IconX_Test.GetFolderIconFile("C:\IconX_Test") & "'"
End Sub

Private Sub cmdGetFolderIconIndex_Click()
Message "The icon file index of folder 'C:\IconX_Test' is:'" & IconX_Test.GetFolderIconIndex("C:\IconX_Test") & "'"
End Sub

Private Sub cmdGetIcon_Click()
Call IconX_Test.GetIcon("C:\", picNew)
MsgBox "The picture shown is the icon of drive 'C:\'", vbInformation
Call IconX_Test.GetIcon("C:\My Documents\My Pictures", picNew)
Call IconX_Test.SavePicToIcon(picNew.hDC, &H44332B, "C:\IconX_Test.ico", SaveTrueColors)
MsgBox "The picture shown is the icon of 'My Pictures' folder", vbInformation
picNew.Picture = LoadPicture("")
End Sub

Private Sub cmdGetRGBofHEX_Click()
Message "The RGB Color of cube background is:'" & IconX_Test.ConvertHexColorToRGB(picIconX.BackColor) & "'"
End Sub

Private Sub cmdReadIniFile_Click()
Message "The value of the key 'Author' in the INI section 'IconX' is:'" & IconX_Test.ReadIni("C:\Testini.ini", "IconX", "Author")
End Sub

Private Sub cmdRebuildIconCache_Click()
IconX_Test.RebuildIconCache
Message "Icon Cache has been built"
End Sub

Private Sub cmdRemoveDriveIcon_Click()
IconX_Test.RemoveDriveIcon "c:\"
Message "Drive 'C:\' icon has been removed sucessfully"
End Sub

Sub Message(MessageToDisp As String)
lblMessage.Caption = MessageToDisp
End Sub

Private Sub cmdRemoveFolderIcon_Click()
IconX_Test.RemoveFolderIcon "C:\IconX_Test"
Message "The folder 'C:\IconX_Test' icon has been removed"
End Sub

Private Sub cmdRestoreIconThumbnail_Click()
IconX_Test.RestoreIconThumbnail
Message "Icon Thumbnails are now restored"
End Sub

Private Sub cmdSkinThisForm_Click()
IconX_Test.TransForm Me
End Sub

Private Sub cmdWriteIniFile_Click()
Call IconX_Test.WriteIni("C:\Testini.ini", "IconX", "Author", "Black Tornado")
Message "Ini file 'C:\Testini.ini' created"
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir "C:\IconX_Test"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
IconX_Test.MoveForm Me
End Sub
