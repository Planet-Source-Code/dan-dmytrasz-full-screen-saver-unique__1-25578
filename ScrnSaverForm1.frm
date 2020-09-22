VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   $"ScrnSaverForm1.frx":0000
   ClientHeight    =   9390
   ClientLeft      =   1680
   ClientTop       =   765
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ListBox MP3Playlist 
      Height          =   1425
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   240
      Top             =   240
   End
   Begin VB.ListBox PlayList 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   5940
      Left            =   120
      Normal          =   0   'False
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.jpg;*.Gif;*.Jpeg"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   3300
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8880
      Visible         =   0   'False
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   9300
      Left            =   120
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim beenclicked As Long
Dim Pause As Boolean
Dim MyPath As String
Dim MyPatterns As String
Dim MySpeed As String
Dim Mode As Integer
Dim ListDone As Boolean
' First, all the Win32 video declares and whatnot.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hwnd&) As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)

    If RunMode = rmScreenSaver Then
        Unload Me
        End
    End If
    
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim hGLRC As Long
Dim TempVar As String
Dim Temp As String
Dim Counter As Long
Dim X As Long
    If Mid$(Command, 1, 2) <> "/p" Then
        LockOn Me
    End If
On Error GoTo Hell
    'Type your initialization code here
  Open App.Path & "\Picview.ini" For Input As #1
    Input #1, Temp
    
    Select Case Temp
      Case "1000"
        Mode = 0
        Input #1, MyPath
      Case "0100"
        Mode = 1
        Input #1, Temp
        For X = 0 To Val(Temp) - 1
          Input #1, Temp
          PlayList.AddItem Temp
        Next X
        ListDone = True
      Case "0010"
        Input #1, Temp
        For X = 0 To Val(Temp) - 1
          Input #1, Temp
          PlayList.AddItem Temp
        Next X
        Mode = 2
      Case "0001"
        Mode = 3
    End Select
    
    Input #1, Temp
    If Val(Temp) > 0 Then
      For X = 0 To Val(Temp) - 1
        Input #1, Temp
        MP3Playlist.AddItem Temp
      Next X
    End If
    
    Input #1, Temp
    MySpeed = Val(Temp)
    Input #1, TempVar
    
    If Val(TempVar) = 1 Then MyPatterns = "*.JPG;"
    Input #1, TempVar
    If Val(TempVar) = 1 Then MyPatterns = MyPatterns & "*.GIF;"
    Input #1, TempVar
    If Val(TempVar) = 1 Then MyPatterns = MyPatterns & "*.BMP;"
  Close #1
  
  If Mode = 0 Then LoadDirectory
  MediaPlayer1_PlayStateChange 0, 0
  Timer1.Interval = Val(MySpeed)
  Timer1.Enabled = True
  'File1.Pattern = MyPatterns

Exit Sub
Hell:
  Close #1
  Timer1.Interval = 1000
  Timer1.Enabled = True
  MyPath = "C:\Windows\system"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If RunMode = rmScreenSaver Then
        Unload Me
        End
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Put the focus away from the screensaver
    If RunMode = rmScreenSaver Then
        LockOff Me
    End If

End Sub
Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static Count As Integer
    Count = Count + 1 ' Give enough time for program to run
    
    If Count > 5 Then
        If RunMode = rmScreenSaver Then
            Unload Me
            End
        End If
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = False

    'If Windows is shut down close this application too
    If UnloadMode = vbAppWindows Then
        Exit Sub
    End If
    
    'if a password is beeing used ask for it and check its validity
    If RunMode = rmScreenSaver And UsePassword Then
        ShowCursor True
        If (VerifyScreenSavePwd(Me.hwnd)) = False Then
            Cancel = True
        End If
        ShowCursor False
    End If

End Sub

Private Sub LoadPic(WhatPic As String)
  Image1.Stretch = False
  Image1.Visible = False
  
  Label1.Caption = "NAME:  " & WhatPic
  Image1.Picture = LoadPicture(WhatPic)
  FitToScreen
End Sub

Private Sub File1_Click()
  If Right(File1.Path, 1) = "\" Then
    LoadPic File1.Path & File1.List(File1.ListIndex)
  Else
    LoadPic File1.Path & "\" & File1.List(File1.ListIndex)
  End If
End Sub

Sub FitToScreen()
  Dim Resp As Single
  If Image1.Width >= Image1.Height Then
    If Image1.Width > MainForm.Width Then
      Resp = MainForm.Width / Image1.Width
      Image1.Stretch = True
      Image1.Width = Image1.Width * Resp
      Image1.Height = Image1.Height * Resp
    End If
  Else
    If Image1.Height > MainForm.Height Then
      Resp = MainForm.Height / Image1.Height
      Image1.Stretch = True
      Image1.Height = Image1.Height * Resp
      Image1.Width = Image1.Width * Resp
    End If
  End If
  CenterPic
  Image1.Visible = True
End Sub

Sub CenterPic()
  Image1.Left = (MainForm.Width - Image1.Width) / 2
  Image1.Top = (MainForm.Height - Image1.Height) / 2
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
  If MediaPlayer1.PlayState = mpClosed Or MediaPlayer1.PlayState = mpStopped Then
    With MP3Playlist
      If .ListCount > 0 Then
        If .ListIndex = .ListCount - 1 Then
          .ListIndex = 0
        Else
          .ListIndex = .ListIndex + 1
        End If
        MediaPlayer1.FileName = .List(.ListIndex)
        MediaPlayer1.Play
      End If
    End With
  End If
End Sub

Private Sub Timer1_Timer()
  Select Case Mode
    Case 0
      CyclcleList1
    Case 1
      If ListDone Then
        If PlayList.ListCount - 1 = PlayList.ListIndex Then
          PlayList.Selected(0) = True
          File1.Path = PlayList.List(PlayList.ListIndex)
        Else
          PlayList.Selected(PlayList.ListIndex + 1) = True
          File1.Path = PlayList.List(PlayList.ListIndex)
        End If
        ListDone = False
      End If
      CyclcleList1
    Case 2
      If PlayList.ListCount - 1 = PlayList.ListIndex Then
        PlayList.Selected(0) = True
      Else
        PlayList.Selected(PlayList.ListIndex + 1) = True
      End If
      LoadPic PlayList.List(PlayList.ListIndex)
    Case 3
  End Select
End Sub

Private Sub CyclcleList1()
  If File1.ListCount - 1 = File1.ListIndex Then
    ListDone = True
    File1.Selected(0) = True
  Else
    File1.Selected(File1.ListIndex + 1) = True
  End If
End Sub

Private Sub LoadDirectory()
  File1.Path = MyPath
  File1.Selected(0) = True
End Sub
