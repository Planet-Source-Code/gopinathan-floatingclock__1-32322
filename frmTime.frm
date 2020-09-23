VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   60
   ClientTop       =   -225
   ClientWidth     =   2490
   ControlBox      =   0   'False
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Hide"
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   1125
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Skins"
      Height          =   555
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   2490
      Begin VB.OptionButton Option1 
         Caption         =   "Default"
         Height          =   285
         Left            =   45
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Green"
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   225
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Gold"
         Height          =   285
         Left            =   1755
         TabIndex        =   1
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3915
      Top             =   855
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   2610
      Top             =   90
      _ExtentX        =   3651
      _ExtentY        =   397
      _Version        =   393216
      Cols            =   13
      Picture         =   "frmTime.frx":0742
   End
   Begin VB.Image imgGoldNum 
      Height          =   225
      Left            =   6120
      Picture         =   "frmTime.frx":1FF4
      Top             =   2880
      Width           =   2070
   End
   Begin VB.Image imgGoldMain 
      Height          =   570
      Left            =   5985
      Picture         =   "frmTime.frx":3898
      Top             =   2205
      Width           =   2520
   End
   Begin VB.Image imgGreenNum 
      Height          =   225
      Left            =   3375
      Picture         =   "frmTime.frx":83AA
      Top             =   2925
      Width           =   2070
   End
   Begin VB.Image imgGreenMain 
      Height          =   540
      Left            =   3285
      Picture         =   "frmTime.frx":9C4E
      Top             =   2205
      Width           =   2505
   End
   Begin VB.Image imgDefaultNum 
      Height          =   225
      Left            =   585
      Picture         =   "frmTime.frx":E370
      Top             =   2835
      Width           =   2070
   End
   Begin VB.Image imgDefaultMain 
      Height          =   555
      Left            =   630
      Picture         =   "frmTime.frx":FC14
      Top             =   2205
      Width           =   2520
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   12
      Left            =   2235
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   11
      Left            =   2070
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   10
      Left            =   1935
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   9
      Left            =   1785
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   8
      Left            =   1620
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   7
      Left            =   135
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   6
      Left            =   300
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   5
      Left            =   450
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   585
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   750
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   900
      Top             =   135
      Width           =   180
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   1080
      Top             =   135
      Width           =   225
   End
   Begin VB.Image imgDisplay 
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   1245
      Top             =   135
      Width           =   225
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   555
      Left            =   0
      Picture         =   "frmTime.frx":1452E
      Top             =   -10
      Width           =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
' The above Declare statement must appear on one line.

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
'Put global constants in bas module.
Dim lsDate As String
Dim Press As Boolean
   Dim cx, cy

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form1.Height = 525
Form1.Width = 2500
End Sub

Private Sub Form_Load()
success% = SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

subInitialize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Form1.Height = 1485
   Form1.Width = 2565
Else
    Press = True
    Form1.MousePointer = 5
    cx = X
    cy = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Press = True Then
Form1.Left = Form1.Left + X - cx
Form1.Top = Form1.Top + Y - cy
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Press = True Then
Press = False
Form1.MousePointer = 0
End If
End Sub

Private Sub Image3_Click()

End Sub


Private Sub mnuExit_Click()
End
End Sub




Private Sub Option1_Click()
Image1.Picture = imgDefaultMain.Picture
PictureClip1.Picture = imgDefaultNum.Picture
subInitialize

End Sub

Private Sub Option2_Click()
Image1.Picture = imgGreenMain.Picture
PictureClip1.Picture = imgGreenNum.Picture
subInitialize

End Sub

Private Sub Option3_Click()
Image1.Picture = imgGoldMain.Picture
PictureClip1.Picture = imgGoldNum.Picture
subInitialize

End Sub

Private Sub Timer1_Timer()
Dim lsTime As String
lsTime = Format(Time, "HH:MM:SS")
'For Seconds
imgDisplay(0).Picture = PictureClip1.GraphicCell(Mid(lsTime, 8, 1))
imgDisplay(1).Picture = PictureClip1.GraphicCell(Mid(lsTime, 7, 1))

'For Minutes
imgDisplay(3).Picture = PictureClip1.GraphicCell(Mid(lsTime, 5, 1))
imgDisplay(4).Picture = PictureClip1.GraphicCell(Mid(lsTime, 4, 1))

'For Hour
imgDisplay(6).Picture = PictureClip1.GraphicCell(Mid(lsTime, 2, 1))
imgDisplay(7).Picture = PictureClip1.GraphicCell(Mid(lsTime, 1, 1))

End Sub

Private Sub subInitialize()

lsDate = Format(Date, "DD-MM")
'Day
imgDisplay(8).Picture = PictureClip1.GraphicCell(Mid(lsDate, 1, 1))
imgDisplay(9).Picture = PictureClip1.GraphicCell(Mid(lsDate, 2, 1))
'Month
imgDisplay(11).Picture = PictureClip1.GraphicCell(Mid(lsDate, 4, 1))
imgDisplay(12).Picture = PictureClip1.GraphicCell(Mid(lsDate, 5, 1))

' Separators for Date and Time
imgDisplay(2).Picture = PictureClip1.GraphicCell(10)
imgDisplay(5).Picture = PictureClip1.GraphicCell(10)
imgDisplay(10).Picture = PictureClip1.GraphicCell(10)
Timer1_Timer

End Sub
