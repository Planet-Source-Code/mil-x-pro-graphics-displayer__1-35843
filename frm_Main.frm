VERSION 5.00
Begin VB.Form frm_Main 
   Caption         =   "Graphics_Displayer"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   7515
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPattern 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2160
      Width           =   2400
   End
   Begin VB.Timer TimerSlideShow 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3000
      Top             =   4080
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   30
      TabIndex        =   11
      Top             =   390
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      Pattern         =   "*.bmp;*.gif;*.jpeg;*.jpg;*.emf;*.wmf;*.ico"
      TabIndex        =   12
      Top             =   2640
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   50
      Left            =   5520
      SmallChange     =   10
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   1215
      LargeChange     =   50
      Left            =   6840
      SmallChange     =   10
      TabIndex        =   8
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   840
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   3
      Top             =   4680
      Width           =   4395
      Begin VB.CheckBox chkExact 
         Height          =   255
         Left            =   2235
         TabIndex        =   17
         ToolTipText     =   "Always display picture at it's original size"
         Top             =   120
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox cboInterval 
         Height          =   315
         ItemData        =   "frm_Main.frx":08CA
         Left            =   3705
         List            =   "frm_Main.frx":08F5
         TabIndex        =   14
         Text            =   "2"
         ToolTipText     =   "SlideShow Interval (Second)"
         Top             =   90
         Width           =   525
      End
      Begin VB.CheckBox chkSlide 
         Height          =   495
         Left            =   2880
         Picture         =   "frm_Main.frx":0925
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Play SlideShow"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cExact 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Picture         =   "frm_Main.frx":0A2E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Original Size"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cMinus 
         Height          =   495
         Left            =   1440
         Picture         =   "frm_Main.frx":0B5B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Zoom Out"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cEqual 
         Height          =   495
         Left            =   720
         Picture         =   "frm_Main.frx":0BE2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fit"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cPlus 
         Height          =   495
         Left            =   0
         Picture         =   "frm_Main.frx":0CAF
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Zoom In"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox picCont 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   3960
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
      Begin VB.PictureBox picTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox picFrom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   3720
      Left            =   2880
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnurename 
         Caption         =   "&Rename"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Public Function GetFile() As String
Dim File As String

If File1.FileName = "" Then Exit Function
    'CHECK IF THE FILE IN ROOT DIR
If Len(File1.Path) > 3 Then
    File = File1.Path & "\" & File1.FileName
Else
    File = File1.Path & File1.FileName
End If
GetFile = File
End Function

Public Sub ResizeViewPort()
On Error Resume Next

Dim to_x As Single
Dim to_y As Single
Dim wid As Single
Dim hgt As Single

If picFrom.Picture = 0 Then Exit Sub
    
    picCont.Cls
    
    wid = picFrom.ScaleWidth
    hgt = picFrom.ScaleHeight
    
If wid > picCont.ScaleWidth Then
    hgt = hgt * picCont.ScaleWidth / wid
    wid = picCont.ScaleWidth
End If
    
If hgt > picCont.ScaleHeight Then
    wid = wid * picCont.ScaleHeight / hgt
    hgt = picCont.ScaleHeight
End If

    to_x = (picCont.ScaleWidth - wid) / 2
    to_y = (picCont.ScaleHeight - hgt) / 2

picTo.Move to_x, to_y, wid, hgt
End Sub

Private Sub StretchBltPix()
On Error GoTo limit
Dim num_trials As Integer
Dim trial As Integer
Dim start_time As Single
Dim fr_wid As Single
Dim fr_hgt As Single
Dim to_wid As Single
Dim to_hgt As Single

    num_trials = CInt(2)
    fr_wid = picFrom.ScaleWidth
    fr_hgt = picFrom.ScaleHeight
    to_wid = picTo.ScaleWidth
    to_hgt = picTo.ScaleHeight

    picTo.Cls
    MousePointer = vbHourglass
    start_time = Timer
    For trial = 1 To num_trials
        StretchBlt picTo.hDC, 0, 0, to_wid, to_hgt, _
        picFrom.hDC, 0, 0, fr_wid, fr_hgt, SRCCOPY
        DoEvents
    Next trial
    MousePointer = vbDefault
Exit Sub
limit:
MsgBox "This is a limit"
MousePointer = vbDefault
End Sub

Private Sub CheckForScrolls()
  If (picTo.Width < picCont.Width) Then
    HScroll1.Value = HScroll1.Min
    HScroll1.Enabled = False
    picTo.Left = (picCont.Width - picTo.Width) / 2
  Else
    With HScroll1
      .Visible = True
      .Enabled = True
      .Min = 0
      .Max = -(picTo.Width - (picCont.Width) + 4)
      .Value = (.Max - .Min) / 2
    End With
  End If
  If (picTo.Height < picCont.Height) Then
    VScroll1.Value = VScroll1.Min
    VScroll1.Enabled = False
    picTo.Top = (picCont.Height - picTo.Height) / 2
  Else
    With VScroll1
      .Visible = True
      .Enabled = True
      .Min = 0
      .Max = -(picTo.Height - (picCont.Height) + 4)
      .Value = (.Max - .Min) / 2
    End With
  End If
End Sub

Private Sub cboInterval_Click()
File1.SetFocus
End Sub


Private Sub cboPattern_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer

pat = cboPattern.List(cboPattern.ListIndex)
p1 = InStr(pat, "(")
p2 = InStr(pat, ")")
File1.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub

Private Sub cEqual_Click()
picTo.Visible = False
ResizeViewPort
StretchBltPix
CheckForScrolls
picTo.Visible = True
File1.SetFocus
End Sub

Private Sub cExact_Click()
On Error Resume Next
Dim to_x As Single
Dim to_y As Single
picTo.Visible = False

picTo.Width = picFrom.Width
picTo.Height = picFrom.Height
to_x = (picCont.ScaleWidth - picTo.Width) / 2
to_y = (picCont.ScaleHeight - picTo.Height) / 2
picTo.Move to_x, to_y
StretchBltPix
CheckForScrolls
picTo.Visible = True
File1.SetFocus
End Sub

Private Sub chkExact_Click()
File1.SetFocus
End Sub

Private Sub chkSlide_Click()

If chkSlide.Value = 1 Then
    If File1.ListCount <= 2 Then
        Call MsgBox("There are not enough files in this directory", vbCritical + vbOKOnly, "Error!")
        Exit Sub
    End If
    If TimerSlideShow.Enabled = True Then
        TimerSlideShow.Enabled = False
        Exit Sub
    Else
        If File1.ListIndex = -1 Then File1.ListIndex = 0
        TimerSlideShow.Enabled = True
    End If
ElseIf chkSlide.Value = 0 Then
    TimerSlideShow.Enabled = False
End If
File1.SetFocus
End Sub

Private Sub cPlus_Click()
Dim to_x As Single
Dim to_y As Single
picTo.Visible = False
picTo.Width = picTo.Width * 1.25
picTo.Height = picTo.Height * 1.25

to_x = (picCont.ScaleWidth - picTo.Width) / 2
to_y = (picCont.ScaleHeight - picTo.Height) / 2
picTo.Move to_x, to_y

StretchBltPix
CheckForScrolls
picTo.Visible = True
File1.SetFocus
End Sub

Private Sub cMinus_Click()
Dim to_x As Single
Dim to_y As Single
picTo.Visible = False
picTo.Width = picTo.Width / 1.25
picTo.Height = picTo.Height / 1.25

to_x = (picCont.ScaleWidth - picTo.Width) / 2
to_y = (picCont.ScaleHeight - picTo.Height) / 2
picTo.Move to_x, to_y

StretchBltPix
CheckForScrolls
picTo.Visible = True
File1.SetFocus
End Sub

Private Sub Dir1_Change()
Me.Caption = "Graphics_Displayer"
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo er
Dir1.Path = Drive1.Drive
Exit Sub
er:
MsgBox "Drive reading error. Error " & err.Number & " : " & err.Description, vbCritical, App.Title
End Sub

Private Sub File1_Click()
On Error GoTo invalidpic
Dim SizePic As Long, size As Long, unit As String

picFrom.Picture = LoadPicture(GetFile)
picTo.Visible = False

If chkExact.Value = 1 Then
    cExact_Click
Else
    ResizeViewPort
    StretchBltPix
End If

CheckForScrolls
picTo.Visible = True
SizePic = FileLen(GetFile)
size = IIf(SizePic \ 1024 > 1, SizePic \ 1024, SizePic)
unit = IIf(SizePic \ 1024 > 1, " KB", " Bytes")
Caption = "Graphics_Displayer - " & File1.FileName & " [" & picFrom.Width & " x " & picFrom.Height & " pixels - " & Format(size, "#,##0") & unit & "]"
Exit Sub
invalidpic:
MsgBox "error " & err.Number & " : " & err.Description, vbCritical, App.Title
picTo.Visible = True
End Sub

Private Sub Form_Load()
ScaleMode = vbPixels
picFrom.ScaleMode = vbPixels
picTo.ScaleMode = vbPixels
picFrom.AutoRedraw = True
picTo.AutoRedraw = True

cboPattern.AddItem "All Graphics Files (*.bmp;*.jpg;*.gif;*.wmf;*.emf;*.ico;*.cur)"
cboPattern.AddItem "Bitmap (*.bmp)"
cboPattern.AddItem "GIF (*.gif)"
cboPattern.AddItem "JPG (*.jpg)"
cboPattern.AddItem "All Files (*.*)"
cboPattern.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frm_Main = Nothing
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
Dim to_x As Single
Dim to_y As Single
Drive1.Move 2, 2
Dir1.Move 2, 26, Dir1.Width, ScaleHeight / 2 - 40
cboPattern.Move 3, Drive1.Height + Dir1.Height + 8
File1.Move 2, Drive1.Height + Dir1.Height + cboPattern.Height + 10, File1.Width, ScaleHeight - Drive1.Height - Dir1.Height - cboPattern.Height - 5
picCont.Move 4 + Drive1.Width, 2, ScaleWidth - Drive1.Width - VScroll1.Width - 4, ScaleHeight - Picture1.Height - HScroll1.Height - 2
HScroll1.Move picCont.Left, picCont.Height + 2, picCont.Width, HScroll1.Height
VScroll1.Move Drive1.Width + picCont.Width + 4, 0, VScroll1.Width, picCont.Height
to_x = (picCont.ScaleWidth - picTo.Width) / 2
to_y = (picCont.ScaleHeight - picTo.Height) / 2
picTo.Move to_x, to_y

Picture1.Move (Me.ScaleWidth - Picture1.Width + Drive1.Width + 4) / 2, picCont.Height + HScroll1.Height + 4
CheckForScrolls
End Sub

Private Sub HScroll1_Change()
picTo.Left = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub mnuAbout_Click()
MsgBox "Graphics_Displayer v 0.3" & vbCrLf & vbCrLf & _
"By Mil-X Pro" & vbCrLf & _
"14 June 2002", vbInformation, "About..."
End Sub

Private Sub mnucopy_Click()
On Error GoTo err
Clipboard.Clear
Clipboard.SetData picFrom.Picture
Exit Sub
err:
Beep
End Sub

Private Sub mnudelete_Click()
On Error GoTo err
Dim sNewName As String
Dim sOldName As String
Dim LastIndex As Long
Dim oldIndex As Long

If Len(GetFile) > 0 Then
    With File1
        If vbOK = MsgBox("Delete " + .FileName + "?", vbOKCancel + vbQuestion, "Confirm Delete") Then
            oldIndex = .ListIndex
            Kill (GetFile)
            .Refresh
            .ListIndex = oldIndex
        End If
    End With
End If
Exit Sub
err:
Beep
End Sub

Private Sub mnuexit_Click()
Set frm_Main = Nothing
End
End Sub

Private Sub mnurename_Click()
Dim sNewName As String
Dim sOldName As String
Dim LastIndex As Long
Dim oldIndex As Long

If Len(GetFile) > 0 Then

    With File1
        sNewName = InputBox("Please type a new name for " & "(" & File1.FileName & ")", "ImejDisplayer", .FileName)
        File1.SetFocus
        If sNewName = "" Then Exit Sub
        sNewName = .Path + "\" + sNewName
        sOldName = .Path + "\" + .FileName
        MoveFile sOldName, sNewName
    End With
    
    LastIndex = File1.ListIndex
    File1.Refresh
    File1.ListIndex = LastIndex
    File1.SetFocus
End If
End Sub


Private Sub TimerSlideShow_Timer()
On Error GoTo wrongtime
If File1.ListIndex = (File1.ListCount - 1) Then File1.ListIndex = 0
TimerSlideShow.Interval = (cboInterval.Text * 1000)
File1.ListIndex = File1.ListIndex + 1
Exit Sub
wrongtime:
cboInterval.Text = "2"
End Sub

Private Sub VScroll1_Change()
picTo.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
picTo.Top = VScroll1.Value
End Sub
