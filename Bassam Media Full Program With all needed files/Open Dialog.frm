VERSION 5.00
Begin VB.Form frm_Open_Dialog 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Media File"
   ClientHeight    =   5250
   ClientLeft      =   3270
   ClientTop       =   1110
   ClientWidth     =   5385
   Icon            =   "Open Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox lstFiles 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   2400
      Width           =   5145
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   1890
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5145
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4430
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   315
      ItemData        =   "Open Dialog.frx":030A
      Left            =   960
      List            =   "Open Dialog.frx":032F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Choose the file type"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   4380
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Look in:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   190
      Width           =   585
   End
End
Attribute VB_Name = "frm_Open_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cmdAdd_Click()
frmmain.Playlist.Refresh
'adds the selected files to the Playlist
    Dim i As Integer
    Dim J As Integer
    For J = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(J) Then
        frmmain.Playlist.AddItem lstFiles.List(J)
            i = i + 1
        End If
    Next J
Call xListKillDupes(frmmain.Playlist) 'calls sub from module
    Unload Me
End Sub

Private Sub Combo1_Click()

If Combo1.ListIndex = 0 Then File1.Pattern = "*.mp3"
If Combo1.ListIndex = 1 Then File1.Pattern = "*.avi"
If Combo1.ListIndex = 2 Then File1.Pattern = "*.asf"
If Combo1.ListIndex = 3 Then File1.Pattern = "*.mpeg"
If Combo1.ListIndex = 4 Then File1.Pattern = "*.mpg"
If Combo1.ListIndex = 5 Then File1.Pattern = "*.wav"
If Combo1.ListIndex = 6 Then File1.Pattern = "*.wmv"
If Combo1.ListIndex = 7 Then File1.Pattern = "*.wma"
If Combo1.ListIndex = 8 Then File1.Pattern = "*.cda"
If Combo1.ListIndex = 9 Then File1.Pattern = "*.mid"
If Combo1.ListIndex = 10 Then File1.Pattern = "*.midi"
'lstFiles.Clear
Dir1_Change
End Sub

Private Sub Dir1_Change()
lstFiles.Clear
File1.Path = Dir1.Path
Dim tel
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        If Len(Dir1.Path) > 3 Then
 lstFiles.AddItem Dir1.Path & "\" & File1.FileName
          Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        lstFiles.AddItem Dir1.Path & File1.FileName
        End If
    Next tel
Else
'    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
End If
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir1.ToolTipText = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
lstFiles.Refresh
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstFiles.ToolTipText = lstFiles.Text
Horizental1
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstFiles.ToolTipText = lstFiles.Text
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstFiles.ToolTipText = lstFiles.Text
End Sub

Function Horizental1()
On Error GoTo b
   Dim c As Long
   Dim rcText As RECT
   Dim newWidth As Long
   Dim itemWidth As Long
   Dim sysScrollWidth As Long
   Me.Font.Name = lstFiles.Font.Name
   Me.Font.Bold = lstFiles.Font.Bold
   Me.Font.Size = lstFiles.Font.Size
   sysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   For c = 0 To lstFiles.ListCount - 1
      Call DrawText(frm_Open_Dialog.hDC, (lstFiles.List(c)), -1&, rcText, DT_CALCRECT)
            itemWidth = rcText.Right + sysScrollWidth
      If itemWidth >= newWidth Then
         newWidth = itemWidth
      End If
   Next
      Call SendMessage(lstFiles.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&)
b:
End Function

