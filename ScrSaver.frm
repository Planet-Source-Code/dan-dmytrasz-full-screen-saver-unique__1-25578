VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6648
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6648
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD1 
      Left            =   480
      Top             =   360
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer PictureChanger 
      Interval        =   3000
      Left            =   120
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   1680
      Top             =   1200
      Width           =   6000
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Byte
Private Const SPI_SCREENSAVERRUNNING = 97


Private Sub Form_Initialize()
  On Error GoTo Errors
    AppActivate "Display Properties"
    Unload Me
    Exit Sub
  
Errors:
  Err.Clear
  Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  ToEnd
End Sub

Private Sub Form_Load()
  Dim X As Long
  Dim MyVar As Long
  X = ShowCursor(False)
  If App.PrevInstance Then ToEnd
  MyVar = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1, ByVal 1&, False)
  MainForm.Show
  PictureChanger_Timer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Timer1.Enabled = False Then ToEnd
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Timer1.Enabled = False Then ToEnd
End Sub

Private Sub PictureChanger_Timer()
  Static X As Byte
  X = X + 1
  If X > 35 Then X = 1
  Image1.Picture = LoadPicture(App.Path & "\" & X & ".jpg")
  CenterPic
End Sub

Sub CenterPic()
    Image1.Top = (MainForm.Height - Image1.Height) / 2
    Image1.Left = (MainForm.Width - Image1.Width) / 2
End Sub

Private Sub Timer1_Timer()
  Static W As Integer
  W = W + 1
  If W > 2 Then Timer1.Enabled = False
End Sub

Sub ToEnd()
  Dim X As Long
  Dim MyVar As Long
  MyVar = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 0, ByVal 1&, False)
  X = ShowCursor(True)
  End
End Sub
