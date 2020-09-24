VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "BASSAM~1.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "BaSSaM MuSiC PLaYeR"
   ClientHeight    =   7620
   ClientLeft      =   2100
   ClientTop       =   630
   ClientWidth     =   7035
   ForeColor       =   &H00000000&
   Icon            =   "BaSSaM Media.frx":0000
   LinkTopic       =   "frmmain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BaSSaM Media.frx":08CA
   ScaleHeight     =   7620
   ScaleWidth      =   7035
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicHiddenData 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   -1365
      Picture         =   "BaSSaM Media.frx":7672
      ScaleHeight     =   15
      ScaleWidth      =   69615
      TabIndex        =   26
      Top             =   7845
      Visible         =   0   'False
      Width           =   69611
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdBack 
      Height          =   495
      Left            =   330
      TabIndex        =   15
      ToolTipText     =   "Move Back"
      Top             =   645
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      ButtonColor     =   0
      PictureBack     =   "BaSSaM Media.frx":AD18
      Style           =   1
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   5160
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6120
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5640
      Top             =   5160
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4560
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdLoadList 
      Height          =   360
      Left            =   1160
      TabIndex        =   8
      ToolTipText     =   "Load PlayList"
      Top             =   1310
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   635
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":B17A
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Slider VolumeSlider 
      Height          =   900
      Left            =   6135
      TabIndex        =   0
      Top             =   705
      Width           =   170
      _ExtentX        =   1588
      _ExtentY        =   291
      PictureBack     =   "BaSSaM Media.frx":B55F
      PictureProgress =   "BaSSaM Media.frx":B81C
      Bar             =   "BaSSaM Media.frx":BB95
      BarOver         =   "BaSSaM Media.frx":BE3A
      BarDown         =   "BaSSaM Media.frx":C0E4
      BackColor       =   0
      Value           =   100
   End
   Begin BaSSaM_MuSiC_PLaYeR.Slider TimeSlider 
      Height          =   120
      Left            =   2230
      TabIndex        =   1
      Top             =   1780
      Width           =   2900
      _ExtentX        =   5106
      _ExtentY        =   212
      PictureProgress =   "BaSSaM Media.frx":C389
      Bar             =   "BaSSaM Media.frx":C8FA
      BarOver         =   "BaSSaM Media.frx":CBC4
      BarDown         =   "BaSSaM Media.frx":CE8F
      BackColor       =   0
      Position        =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdSaveList 
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      ToolTipText     =   "Save PlayList"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":D15A
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdOpen 
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      ToolTipText     =   "Add Files"
      Top             =   300
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":D4AC
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdClear 
      Height          =   300
      Left            =   2480
      TabIndex        =   10
      ToolTipText     =   "Clear PlayList"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":D925
      PictureDown     =   "BaSSaM Media.frx":DC38
      PictureOver     =   "BaSSaM Media.frx":DF4B
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button Minimize 
      Height          =   180
      Left            =   4720
      TabIndex        =   11
      ToolTipText     =   "Minimize Prog"
      Top             =   0
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":E25C
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdExit 
      Height          =   180
      Left            =   4940
      TabIndex        =   12
      ToolTipText     =   "Close Prog"
      Top             =   0
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":E4F8
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdRemove 
      Height          =   300
      Left            =   2980
      TabIndex        =   13
      ToolTipText     =   "Delete Item"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":E7B5
      PictureDown     =   "BaSSaM Media.frx":EAD6
      PictureOver     =   "BaSSaM Media.frx":EDD8
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdNext 
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Move Next"
      Top             =   645
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      ButtonColor     =   0
      PictureBack     =   "BaSSaM Media.frx":F0E7
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdStop 
      Height          =   310
      Left            =   520
      TabIndex        =   17
      ToolTipText     =   "Stop Song"
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      ButtonColor     =   0
      PictureBack     =   "BaSSaM Media.frx":F50E
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdPause 
      Height          =   255
      Left            =   540
      TabIndex        =   18
      ToolTipText     =   "Pause Song"
      Top             =   455
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      ButtonColor     =   0
      PictureBack     =   "BaSSaM Media.frx":F8FE
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdPlay 
      Default         =   -1  'True
      Height          =   405
      Left            =   660
      TabIndex        =   19
      ToolTipText     =   "Play Song"
      Top             =   770
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
      ButtonColor     =   0
      PictureBack     =   "BaSSaM Media.frx":FD03
      Style           =   1
   End
   Begin VB.ListBox Playlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1980
      Left            =   285
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   5240
      Width           =   6360
   End
   Begin BaSSaM_MuSiC_PLaYeR.Slider Slider1 
      Height          =   120
      Left            =   2430
      TabIndex        =   20
      ToolTipText     =   "Balance"
      Top             =   1360
      Width           =   615
      _ExtentX        =   212
      _ExtentY        =   1085
      Bar             =   "BaSSaM Media.frx":10072
      BarOver         =   "BaSSaM Media.frx":10314
      BarDown         =   "BaSSaM Media.frx":105B4
      BackColor       =   0
      Value           =   50
      Position        =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdFull 
      Height          =   300
      Left            =   6000
      TabIndex        =   21
      ToolTipText     =   "Full Screen"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":10856
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button Shuffle 
      Height          =   300
      Left            =   3470
      TabIndex        =   22
      ToolTipText     =   "Shuffle Play"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":10BA4
      PictureDown     =   "BaSSaM Media.frx":10EBF
      PictureOver     =   "BaSSaM Media.frx":111DA
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button Looop 
      Height          =   300
      Left            =   3990
      TabIndex        =   23
      ToolTipText     =   "Continous Play"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":114F0
      PictureDown     =   "BaSSaM Media.frx":117FD
      PictureOver     =   "BaSSaM Media.frx":11B0A
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdMoveUp1 
      Height          =   300
      Left            =   4580
      TabIndex        =   24
      ToolTipText     =   "Move Sel Up"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":11E1B
      PictureDown     =   "BaSSaM Media.frx":12127
      PictureOver     =   "BaSSaM Media.frx":12629
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button cmdMoveDown 
      Height          =   300
      Left            =   5170
      TabIndex        =   25
      ToolTipText     =   "Move Sel Dn"
      Top             =   330
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":1295A
      PictureDown     =   "BaSSaM Media.frx":12C71
      PictureOver     =   "BaSSaM Media.frx":13173
      Style           =   1
   End
   Begin BaSSaM_MuSiC_PLaYeR.Button tray 
      Height          =   180
      Left            =   4545
      TabIndex        =   31
      ToolTipText     =   "Send To Tray"
      Top             =   15
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      ButtonColor     =   16777215
      PictureBack     =   "BaSSaM Media.frx":134A3
      Style           =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   3480
      TabIndex        =   32
      Top             =   765
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   30
      Top             =   7275
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--<Play-List>--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2925
      TabIndex        =   29
      Top             =   4845
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visualation && Vedios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   28
      Top             =   2145
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programed && Designed By BaSSaM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3600
      TabIndex        =   27
      Top             =   7395
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image cont 
      Height          =   255
      Left            =   5385
      Picture         =   "BaSSaM Media.frx":13778
      Stretch         =   -1  'True
      Top             =   1290
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image shuff 
      Height          =   255
      Left            =   5385
      Picture         =   "BaSSaM Media.frx":13C6A
      Stretch         =   -1  'True
      Top             =   1290
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2160
      Picture         =   "BaSSaM Media.frx":1415C
      Stretch         =   -1  'True
      ToolTipText     =   "Balance"
      Top             =   1300
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3045
      Picture         =   "BaSSaM Media.frx":1459E
      Stretch         =   -1  'True
      ToolTipText     =   "Balance"
      Top             =   1300
      Width           =   255
   End
   Begin WMPLibCtl.WindowsMediaPlayer Media 
      Height          =   1995
      Left            =   1680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2370
      Width           =   3585
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6324
      _cy             =   3519
   End
   Begin VB.Label SongTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Song Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   5
      Top             =   1050
      Width           =   3975
   End
   Begin VB.Label SongTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "\00:00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   760
      Width           =   495
   End
   Begin VB.Label SongDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   3
      Top             =   760
      Width           =   495
   End
   Begin VB.Label lblVolume 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume 100 %"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4845
      TabIndex        =   2
      Top             =   760
      Width           =   975
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iRet As Integer
Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean
Dim Fso As New FileSystemObject
Dim CurRgn, TempRgn As Long  ' Region variables
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Fast binary Data
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Dim PicInfo As BITMAP

Private Type BITMAP
 bmType As Long
 bmWidth As Long
 bmHeight As Long
 bmWidthBytes As Long
 bmPlanes As Integer
 bmBitsPixel As Integer
 bmBits As Long
End Type

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBack.Top = 645 + 15
Label5.Caption = cmdBack.ToolTipText
End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = cmdClear.ToolTipText
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.Top = 15
Minimize.Top = 0
tray.Top = 15
Label5.Caption = cmdExit.ToolTipText
End Sub

Private Sub cmdFull_Click()
Media.fullScreen = True
End Sub

Private Sub cmdFull_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFull.Top = 330 + 15
Label5.Caption = cmdFull.ToolTipText
End Sub

Private Sub cmdLoadList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdLoadList.Top = 1310 + 15
Label5.Caption = cmdLoadList.ToolTipText
End Sub

Private Sub cmdMoveDown_Click()
On Error GoTo b
iRet = MoveDown_Click(Playlist)
b:
End Sub

Private Sub cmdMoveDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = cmdMoveDown.ToolTipText
End Sub

Private Sub cmdMoveUp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = cmdMoveUp1.ToolTipText
End Sub

Private Sub cmdMoveUp1_Click()
On Error GoTo b
iRet = MoveUp_Click(Playlist)
b:
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNext.Top = 645 + 15
Label5.Caption = cmdNext.ToolTipText
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOpen.Top = 300 + 15
Label5.Caption = cmdOpen.ToolTipText
End Sub

Private Sub cmdPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPause.Top = 455 + 15
Label5.Caption = cmdPause.ToolTipText
End Sub

Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPlay.Top = 770 + 15
Label5.Caption = cmdPlay.ToolTipText
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = cmdRemove.ToolTipText
End Sub

Private Sub cmdSaveList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSaveList.Top = 330 + 15
Label5.Caption = cmdSaveList.ToolTipText
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdStop.Top = 1200 + 10
cmdStop.Left = 520 + 20
Label5.Caption = cmdStop.ToolTipText
End Sub

Private Sub Form_Load()
Horizental
Dim Region As Long
Dim ByteCtr As Long
Dim ByteData(18559) As Byte

ByteCtr = 18560
'Get the Data
GetObject PicHiddenData.Image, Len(PicInfo), PicInfo
GetBitmapBits PicHiddenData.Image, ByteCtr, ByteData(0)

'Shape The Form
Region = ExtCreateRegion(ByVal 0&, ByteCtr, ByteData(0))
SetWindowRgn Me.hwnd, Region, True

If Timer2.Enabled = True Then
shuff.Visible = True
cont.Visible = False
End If

If Timer4.Enabled = True Then
shuff.Visible = False
cont.Visible = True
End If
VolumeSlider.Value = 100
Dim file As String
file = App.Path & "\" & "Registry.dat"
Dim A As String
Dim X As String
On Error GoTo Error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Playlist.AddItem A$
Loop
Close 1
Exit Sub
Error:
''---------------------------------------
End Sub

Private Sub cmdBack_Click()
On Error GoTo b:
Playlist.ListIndex = Playlist.ListIndex - 1
Media.URL = SongTitle.Caption
Media.URL = Playlist.Text
On Error Resume Next
Media.Controls.play
SongTitle.Caption = Playlist.Text
b:
End Sub

Private Sub cmdClear_Click()
Playlist.Clear
SongTitle.Caption = ""
End Sub

Private Sub cmdExit_Click()
Unload frmmain
Unload frm_Open_Dialog
End Sub

Private Sub CmdLoadList_Click()
Dim file As String
Dialog.DialogTitle = "Load Bassam PlayList."
Dialog.MaxFileSize = 16384
Dialog.FileName = ""
Dialog.Filter = "Bassam PlayList Files|*.Bassam"
Dialog.ShowOpen     ' = 1
If Dialog.FileName = "" Then Exit Sub
file = Dialog.FileName
Dim A As String
Dim X As String
On Error GoTo Error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Playlist.AddItem A$
Loop
Close 1
Exit Sub
Call xListKillDupes(Playlist) 'calls sub from module
Error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub cmdNext_Click()
On Error GoTo b:
Playlist.ListIndex = Playlist.ListIndex + 1
Media.URL = SongTitle.Caption
Media.URL = Playlist.Text
On Error Resume Next
Media.Controls.play
SongTitle.Caption = Playlist.Text
b:
End Sub

Private Sub cmdOpen_Click()
frm_Open_Dialog.Show vbModal
End Sub

Private Sub CmdPause_Click()
On Error GoTo b
If Playlist.ListCount = 0 Then Exit Sub
If SongTitle.Caption = "" Then Exit Sub
If cmdPause.ToolTipText = "Pause Song" Then
Media.Controls.pause
'cmdPause.ToolTipText = "Resume"
Else
'Media.Controls.play
'cmdPause.ToolTipText = "Pause"
End If
b:
End Sub

Private Sub CmdPlay_Click()
SongTitle.Caption = Playlist.Text
On Error Resume Next
Media.URL = SongTitle.Caption
If SongTitle.Caption <> "" Then
Media.Controls.play
Media.Controls.currentPosition = TimeSlider.Value
cmdPause.ToolTipText = "Pause Song"
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub cmdRemove_Click()
If Playlist.ListIndex = -1 Then
MsgBox "No file selected", vbExclamation, "Error"
Else
Playlist.RemoveItem Playlist.ListIndex
SongTitle.Caption = ""
End If
End Sub

Private Sub cmdSaveList_Click()
On Error Resume Next
Dim intRecord As Integer
    Dim strFilePath As String
    Dim ListData As Variant
    With Dialog
        .Flags = cdlOFNOverwritePrompt
       '.InitDir = App.Path
        .DefaultExt = "Bassam"
        .Filter = "Bassam Media PlayList Files|*.Bassam"
        .ShowSave
        strFilePath = .FileName
    End With
    If strFilePath <> "" Then
        Open strFilePath For Output As #1
        For intRecord = 0 To Playlist.ListCount - 1
            Write #1, Playlist.List(intRecord)
        Next intRecord
        Close #1
    End If
End Sub

Private Sub cmdStop_Click()
Media.Controls.pause
TimeSlider.Value = 0
Media.Controls.currentPosition = TimeSlider.Value
SongTitle.Caption = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMe = True
OldX = X
OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveMe = True Then
  frmmain.Left = frmmain.Left + (X - OldX)
  frmmain.Top = frmmain.Top + (Y - OldY)
End If
Minimize.Top = 0
cmdExit.Top = 0
tray.Top = 15
cmdFull.Top = 330
cmdSaveList.Top = 330
cmdOpen.Top = 300
cmdLoadList.Top = 1310
cmdStop.Top = 1200
cmdStop.Left = 520
cmdNext.Top = 645
cmdPlay.Top = 770
cmdPause.Top = 455
cmdBack.Top = 645
Label5.Caption = ""
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmmain.Left = frmmain.Left + (X - OldX)
frmmain.Top = frmmain.Top + (Y - OldY)
MoveMe = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmmain = Nothing 'good practice to free resources VB doesn't normally free when you unload a form!

On Error GoTo b
Open (App.Path & "\" & "Registry.dat") For Output As #1
       Dim i%
       For i = 0 To Playlist.ListCount - 1
       Print #1, Playlist.List(i)
       Next
       Close #1
b:
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMe = True
OldX = X
OldY = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveMe = True Then
  frmmain.Left = frmmain.Left + (X - OldX)
  frmmain.Top = frmmain.Top + (Y - OldY)
End If
cmdExit.Top = 0
Minimize.Top = 0
tray.Top = 15
cmdFull.Top = 330
cmdSaveList.Top = 330
Label5.Caption = ""
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmmain.Left = frmmain.Left + (X - OldX)
frmmain.Top = frmmain.Top + (Y - OldY)
MoveMe = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMe = True
OldX = X
OldY = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveMe = True Then
  frmmain.Left = frmmain.Left + (X - OldX)
  frmmain.Top = frmmain.Top + (Y - OldY)
End If
Minimize.Top = 0
cmdExit.Top = 0
tray.Top = 0
cmdFull.Top = 330
cmdSaveList.Top = 330
cmdOpen.Top = 300
cmdLoadList.Top = 1310
cmdStop.Top = 1200
cmdStop.Left = 520
cmdNext.Top = 645
cmdPlay.Top = 770
cmdPause.Top = 455
cmdBack.Top = 645
Label5.Caption = "The Author"
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmmain.Left = frmmain.Left + (X - OldX)
frmmain.Top = frmmain.Top + (Y - OldY)
MoveMe = False
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMe = True
OldX = X
OldY = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveMe = True Then
  frmmain.Left = frmmain.Left + (X - OldX)
  frmmain.Top = frmmain.Top + (Y - OldY)
End If
Minimize.Top = 0
cmdExit.Top = 0
tray.Top = 0
cmdFull.Top = 330
cmdSaveList.Top = 330
cmdOpen.Top = 300
cmdLoadList.Top = 1310
cmdStop.Top = 1200
cmdStop.Left = 520
cmdNext.Top = 645
cmdPlay.Top = 770
cmdPause.Top = 455
cmdBack.Top = 645
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmmain.Left = frmmain.Left + (X - OldX)
frmmain.Top = frmmain.Top + (Y - OldY)
MoveMe = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveMe = True
OldX = X
OldY = Y
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveMe = True Then
  frmmain.Left = frmmain.Left + (X - OldX)
  frmmain.Top = frmmain.Top + (Y - OldY)
End If
Minimize.Top = 0
cmdExit.Top = 0
tray.Top = 0
cmdFull.Top = 330
cmdSaveList.Top = 330
cmdOpen.Top = 300
cmdLoadList.Top = 1310
cmdStop.Top = 1200
cmdStop.Left = 520
cmdNext.Top = 645
cmdPlay.Top = 770
cmdPause.Top = 455
cmdBack.Top = 645
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmmain.Left = frmmain.Left + (X - OldX)
frmmain.Top = frmmain.Top + (Y - OldY)
MoveMe = False
End Sub

Private Sub Looop_Click()
Timer2.Enabled = False
If Timer4.Enabled = False Then
Timer4.Enabled = True
shuff.Visible = False
cont.Visible = True
Exit Sub
End If
If Timer4.Enabled = True Then
Timer4.Enabled = False
Exit Sub
End If
End Sub

Private Sub Looop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = Looop.ToolTipText
End Sub

Private Sub Media_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
Label5.Caption = "Vedio Screen"
End Sub

Private Sub Minimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Minimize.Top = 15
cmdExit.Top = 0
tray.Top = 15
Label5.Caption = Minimize.ToolTipText
End Sub

Private Sub Playlist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Playlist.ToolTipText = SongTitle.Caption
Label5.Caption = "Play-List"
End Sub

Private Sub Shuffle_Click()
Timer4.Enabled = False
If Timer2.Enabled = False Then
Timer2.Enabled = True
shuff.Visible = True
cont.Visible = False
Exit Sub
End If
If Timer2.Enabled = True Then
Timer2.Enabled = False
Exit Sub
End If
End Sub

Private Sub Media_OpenStateChange(ByVal NewState As Long)
If Timer2.Enabled = True Then
shuff.Visible = True
cont.Visible = False
End If

If Timer4.Enabled = True Then
shuff.Visible = False
cont.Visible = True
End If

On Error GoTo b:
Timer1.Enabled = True
b:
End Sub

Private Sub Minimize_Click()
frmmain.WindowState = 1
End Sub

Private Sub Playlist_Click()
SongTitle.Caption = Playlist.Text
Horizental
End Sub

Private Sub Playlist_DblClick()
SongTitle.Caption = Playlist.Text
On Error Resume Next
Media.URL = SongTitle.Caption
If SongTitle.Caption <> "" Then
Media.Controls.play
TimeSlider.Max = Media.currentMedia.duration
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub Shuffle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = Shuffle.ToolTipText
End Sub

Private Sub Slider1_Change(Value As Long)
On Error GoTo b
If Slider1.Value > -500 And Slider1.Value < 500 Then
End If
If Slider1.Value < -500 Then
End If
If Slider1.Value > 500 Then
End If
Media.settings.balance = Slider1.Value
Exit Sub
b:
MsgBox "Err"
Exit Sub
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Balance Bar"
End Sub

Private Sub Timer1_Timer()
On Error GoTo F
TimeSlider.Value = Media.Controls.currentPosition
TimeSlider.Max = Media.currentMedia.duration
If Media.currentMedia.duration > 0 Then
Else
Exit Sub
End If
Dim i As Integer
Dim min As Integer
Dim sec As Integer
i = Val(Format(Media.Controls.currentPosition, "###"))
If i > 59 Then
min = i \ 60
sec = i Mod 60
SongDuration.Caption = Format(min, "0#") & ":" & Format(sec, "00")
Else
If i > -1 Then
SongDuration.Caption = "00" & ":" & Format(i, "0#")
End If
End If

i = Val(Format(frmmain.Media.currentMedia.duration, "###"))
If i > 59 Then
min = i \ 60
sec = i Mod 60
SongTime.Caption = "/" & Format(min, "0#") & ":" & Format(sec, "00")
Else
If i > -1 Then
End If
End If
F:
End Sub

Private Sub Timer2_Timer()
On Error GoTo b:
Dim rand$
Dim blah$
If Media.playState = wmppsStopped Then
On Error Resume Next
Playlist.ListIndex = Module1.RandomNumber(Playlist.ListCount)
rand$ = Playlist.Text
On Error Resume Next
Media.URL = rand$
Media.Controls.play
Playlist.ListIndex = Playlist.Text
blah$ = Module1.ReplaceString(Playlist.Text, ".mp3 ", "")
Playlist.Text = Playlist.ListIndex
SongTitle.Caption = Media.URL
Timer1.Enabled = True
End If
b:
End Sub

Private Sub Timer4_Timer()
On Error GoTo b:
If Media.playState = wmppsStopped Then
Playlist.ListIndex = Playlist.ListIndex + 1
Media.URL = Playlist.Text
On Error Resume Next
Media.Controls.play
End If
b:
End Sub

Private Sub TimeSlider_Change(Value As Long)
Media.Controls.currentPosition = TimeSlider.Value
End Sub

Private Sub TimeSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Time Bar"
End Sub

Private Sub tray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Minimize.Top = 0
cmdExit.Top = 0
tray.Top = 30
Label5.Caption = tray.ToolTipText
End Sub

Private Sub VolumeSlider_Change(Value As Long)
Media.settings.volume = VolumeSlider.Value
lblVolume.Caption = "Volume " & VolumeSlider.Value & " %"
End Sub

Private Sub VolumeSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Media.settings.volume = VolumeSlider.Value
lblVolume.Caption = "Volume " & VolumeSlider.Value & " %"
End Sub

Private Sub VolumeSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Volume Bar"
End Sub

Function Horizental()
On Error GoTo b
   Dim c As Long
   Dim rcText As RECT
   Dim newWidth As Long
   Dim itemWidth As Long
   Dim sysScrollWidth As Long
   Me.Font.Name = Playlist.Font.Name
   Me.Font.Bold = Playlist.Font.Bold
   Me.Font.Size = Playlist.Font.Size
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   For c = 0 To Playlist.ListCount - 1
      Call DrawText(frmmain.hDC, (Playlist.List(c)), -1&, rcText, DT_CALCRECT)
            itemWidth = rcText.Right + sysScrollWidth
      If itemWidth >= newWidth Then
         newWidth = itemWidth
      End If
   Next
      Call SendMessage(Playlist.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)
b:
End Function

Public Function MoveUp_Click(lstMove As listbox) As Integer
On Error GoTo b
 'not by source
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    iCnt = lstMove.ListIndex
    If iCnt > -1 Then
         strTemp1 = lstMove.List(iCnt)
        '-- Add the item selected to one position above the current position
        lstMove.AddItem strTemp1, (iCnt - 1)
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstMove.RemoveItem (iCnt + 1)
        '-- Reselect the item that was moved.
             lstMove.Selected(iCnt - 1) = True
    End If
b:
End Function
Public Function MoveDown_Click(lstMove As listbox) As Integer
On Error GoTo b
    Dim strTemp1 As String    '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer    '-- holds the index of the item to be moved
    '-- Assign the first index
    iCnt = lstMove.ListIndex
    If iCnt > -1 Then
         strTemp1 = lstMove.List(iCnt)
        '-- Add the item selected to below the current position
        lstMove.AddItem strTemp1, (iCnt + 2)
        lstMove.RemoveItem (iCnt)
        '-- Reselect the item that was moved.
        lstMove.Selected(iCnt + 1) = True
   End If
b:
End Function

