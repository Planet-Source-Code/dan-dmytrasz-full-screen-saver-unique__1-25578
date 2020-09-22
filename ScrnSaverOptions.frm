VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Options 
   Caption         =   "Options"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pictures"
      TabPicture(0)   =   "ScrnSaverOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ModeData(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ModeData(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "MP3"
      TabPicture(1)   =   "ScrnSaverOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "File3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Dir3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Drive3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "List2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   255
         Left            =   -67820
         TabIndex        =   31
         Top             =   615
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Accept"
         Height          =   375
         Left            =   -68160
         TabIndex        =   30
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "test Song"
         Height          =   735
         Left            =   -74760
         TabIndex        =   26
         Top             =   2640
         Width           =   6405
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Height          =   615
            Left            =   5880
            TabIndex        =   29
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   5775
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
            Height          =   375
            Left            =   5880
            TabIndex        =   27
            Top             =   240
            Width           =   495
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
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   7695
      End
      Begin VB.DriveListBox Drive3 
         Height          =   315
         Left            =   -74880
         TabIndex        =   24
         Top             =   3480
         Width           =   3435
      End
      Begin VB.DirListBox Dir3 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   23
         Top             =   3840
         Width           =   3435
      End
      Begin VB.FileListBox File3 
         Height          =   3405
         Left            =   -71400
         Pattern         =   "*.MP3"
         TabIndex        =   22
         Top             =   3495
         Width           =   4515
      End
      Begin VB.Frame ModeData 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   6615
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   4815
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   720
            Width           =   4750
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   0
            TabIndex        =   18
            Top             =   1080
            Width           =   4750
         End
         Begin VB.FileListBox File1 
            Height          =   2820
            Left            =   0
            Pattern         =   "*.jpg;*.gif;*.jpeg"
            TabIndex        =   17
            Top             =   3720
            Width           =   4750
         End
         Begin VB.Label DirHolder 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   240
            Width           =   4750
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Directory Location"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   0
            Width           =   4215
         End
      End
      Begin VB.Frame ModeData 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   6700
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
         Begin VB.FileListBox File2 
            Height          =   4380
            Left            =   2400
            Pattern         =   "*.jpg;*.gif;*.jpeg"
            TabIndex        =   14
            Top             =   2230
            Width           =   2415
         End
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   0
            TabIndex        =   13
            Top             =   1850
            Width           =   4800
         End
         Begin VB.DirListBox Dir2 
            Height          =   4365
            Left            =   0
            TabIndex        =   12
            Top             =   2230
            Width           =   2355
         End
         Begin VB.ListBox List1 
            Height          =   1425
            Left            =   0
            MultiSelect     =   2  'Extended
            OLEDropMode     =   1  'Manual
            TabIndex        =   11
            Top             =   360
            Width           =   4780
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear"
            Height          =   255
            Left            =   3920
            TabIndex        =   10
            Top             =   380
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Directory Image List"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   4695
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Accept"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   6600
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Mode"
         Height          =   1935
         Left            =   5280
         TabIndex        =   3
         Top             =   480
         Width           =   2535
         Begin VB.OptionButton Modes 
            Caption         =   "Multi-Directory Display"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Modes 
            Caption         =   "Custom Image List"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton Modes 
            Caption         =   "Single Directory Display"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Modes 
            Caption         =   "Complete Drive Display"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   4
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Preview"
         Height          =   3615
         Left            =   5280
         TabIndex        =   1
         Top             =   2760
         Width           =   2535
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   3255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2295
            Begin VB.Image Image1 
               Height          =   3255
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   2295
            End
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnuType 
         Caption         =   "&Picture Types"
         Begin VB.Menu MnuJpg 
            Caption         =   "JPG"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuGif 
            Caption         =   "GIF"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuBmp 
            Caption         =   "BMP"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MnuDuration 
         Caption         =   "&Picture Duration"
         Begin VB.Menu MyDuration 
            Caption         =   "700 Ms"
         End
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function SaveOptions()
  Static Count As Boolean
  Dim Temp As String
  
  If Count = False Then
    Open App.Path & "\Picview.ini" For Output As #1
     If Modes(0).Value Then
       Temp = "1"
     Else
       Temp = "0"
     End If
     If Modes(1).Value Then
       Temp = Temp & "1"
     Else
       Temp = Temp & "0"
     End If
     If Modes(2).Value Then
       Temp = Temp & "1"
     Else
       Temp = Temp & "0"
     End If
     If Modes(3).Value Then
       Temp = Temp & "1"
     Else
       Temp = Temp & "0"
     End If
     
     Print #1, Temp         '*************************
     
     Select Case Temp
       Case "1000"
         Temp = DirHolder.Caption
         Print #1, Temp     '*************************
       Case "0100", "0010"
         Temp = List1.ListCount
         Print #1, Temp     '*************************
         
         For X = 0 To Val(Temp) - 1
           Temp = List1.List(X)
           Print #1, Temp   '*************************
         Next X
         
       Case "0001"
     End Select
     
     Temp = List2.ListCount
     Print #1, Temp         '*************************
     
     If Val(Temp) > 0 Then
       For X = 0 To Val(Temp) - 1
         Temp = List2.List(X)
         Print #1, Temp     '*************************
       Next X
     End If
     
     Temp = Left(MyDuration.Caption, Len(MyDuration.Caption) - 3)
     Print #1, Temp         '*************************
     Temp = MnuJpg.Checked
     Print #1, Temp         '*************************
     Temp = MnuGif.Checked
     Print #1, Temp         '*************************
     Temp = MnuBmp.Checked
     Print #1, Temp         '*************************
    Close #1
    Count = True
  End If
End Function

Private Sub Command1_Click()
  SaveOptions
  Form_QueryUnload 0, 0
End Sub

Private Sub Command2_Click()
  List1.Clear
End Sub

Private Sub Command3_Click()
  Command1_Click
End Sub

Private Sub Command4_Click()
  List2.Clear
End Sub

Private Sub Dir1_Change()
  DirHolder.Caption = Dir1.Path
  File1.Path = Dir1.Path
End Sub

Private Sub Dir2_Change()
  File2.Path = Dir2.Path
End Sub

Private Sub Dir2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Modes(1).Value = True Then
    If Button = 2 Then
      List1.AddItem Dir2.List(Dir2.ListIndex)
    End If
  End If
End Sub

Private Sub Dir3_Change()
  File3.Path = Dir3.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive3_Change()
  Dir3.Path = Drive3.Drive
End Sub

Private Sub File1_Click()
  ImageManipulation File1.Path, File1.List(File1.ListIndex)
End Sub

Private Sub File2_Click()
  ImageManipulation File2.Path, File2.List(File2.ListIndex)
End Sub

Private Sub ImageManipulation(MyPath As String, MyVal As String)
  
  Image1.Visible = False
  Image1.Stretch = False
  If Right(MyPath, 1) = "\" Then
    Image1.Picture = LoadPicture(MyPath & MyVal)
  Else
    Image1.Picture = LoadPicture(MyPath & "\" & MyVal)
  End If
  
  Dim Resp As Single
  
  If Image1.Width >= Image1.Height Then
    If Image1.Width > Frame2.Width Then
      Resp = Frame2.Width / Image1.Width
      Image1.Stretch = True
      Image1.Width = Image1.Width * Resp
      Image1.Height = Image1.Height * Resp
    End If
  Else
    If Image1.Height > Frame2.Height Then
      Resp = Frame2.Height / Image1.Height
      Image1.Stretch = True
      Image1.Height = Image1.Height * Resp
      Image1.Width = Image1.Width * Resp
    End If
  End If
  
  Image1.Left = (Frame2.Width - Image1.Width) / 2
  Image1.Top = (Frame2.Height - Image1.Height) / 2

  Image1.Visible = True

End Sub

Private Sub File2_DblClick()
  If Modes(2).Value = True Then
    If Right(App.Path, 1) = "\" Then
      List1.AddItem File2.Path & File2.List(File2.ListIndex)
    Else
      List1.AddItem File2.Path & "\" & File2.List(File2.ListIndex)
    End If
  End If
End Sub

Private Sub File3_Click()
  If File3.ListIndex <> -1 Then Label2.Caption = File3.List(File3.ListIndex)
End Sub

Private Sub File3_DblClick()
  If Right(File3.Path, 1) = "\" Then
    List2.AddItem File3.Path & File3.List(File3.ListIndex)
  Else
    List2.AddItem File3.Path & "\" & File3.List(File3.ListIndex)
  End If
End Sub

Private Sub Form_Load()
Dim Temp As String

On Error GoTo Hell
  Open App.Path & "\Picview.ini" For Input As #1
    Input #1, Temp              '*************************
    Select Case Trim(Temp)
      Case "1000"
        Modes(0).Value = True
        Input #1, Temp          '*************************
        Dir1.Path = Temp
      Case "0100"
        Modes(1).Value = True
        Input #1, Temp          '*************************
        For X = 0 To Val(Temp) - 1
          Input #1, Temp        '*************************
          List1.AddItem Temp
        Next X
      Case "0010"
        Modes(2).Value = True
        Input #1, Temp          '*************************
        For X = 0 To Val(Temp) - 1
          Input #1, Temp        '*************************
          List1.AddItem Temp
        Next X
      Case "0001"
        Modes(3).Value = True
        Input #1, Temp          '*************************
    End Select
    
    Input #1, Temp
    If Val(Temp) > 0 Then
      For X = 0 To Val(Temp) - 1
        Input #1, Temp          '*************************
        List2.AddItem Temp
      Next X
    End If
    
    Input #1, Temp              '*************************
    MyDuration.Caption = Temp & " Ms"
    Input #1, Temp              '*************************
    MnuJpg.Checked = Temp
    Input #1, Temp              '*************************
    MnuGif.Checked = Temp
    Input #1, Temp              '*************************
    MnuBmp.Checked = Temp
    
    If DirHolder.Caption = "" Then Dir1.Path = "C:\"
    Dir2.Path = "C:\"
    Dir3.Path = "C:\"
    Close #1
Exit Sub
Hell:
  Close #1
  MyDuration.Caption = "1000 Ms"
  Modes(0).Value = True
  MnuJpg.Checked = True
  MnuGif.Checked = True
  MnuBmp.Checked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Label3_Click()
  If Right(File3.Path, 1) = "\" Then
    MediaPlayer1.FileName = File3.Path & Label2.Caption
  Else
    MediaPlayer1.FileName = File3.Path & "\" & Label2.Caption
  End If
  MediaPlayer1.Play
End Sub

Private Sub List1_DblClick()
  List1.RemoveItem List1.ListIndex
End Sub

Private Sub List2_DblClick()
  List2.RemoveItem List2.ListIndex
End Sub

Private Sub MnuExit_Click()
  Unload Me
End Sub

Private Sub MnuJpg_Click()
  If MnuJpg.Checked Then
    MnuJpg.Checked = False
  Else
    MnuJpg.Checked = True
  End If
End Sub

Private Sub Mnugif_Click()
  If MnuGif.Checked Then
    MnuGif.Checked = False
  Else
    MnuGif.Checked = True
  End If
End Sub

Private Sub Mnubmp_Click()
  If MnuBmp.Checked Then
    MnuBmp.Checked = False
  Else
    MnuBmp.Checked = True
  End If
End Sub

Private Sub Modes_Click(Index As Integer)
  List1.Clear
  Select Case Index
    Case 0
      ModeData(0).Visible = True
      ModeData(1).Visible = False
    Case 1
      Label4.Caption = "Directory View Listings"
      ModeData(1).Visible = True
      ModeData(0).Visible = False
    Case 2
      Label4.Caption = "Image Listings"
      ModeData(1).Visible = True
      ModeData(0).Visible = False
    Case 3
    
  End Select
End Sub

Private Sub MyDuration_Click()
  MyDuration.Caption = InputBox("Enter how long you want each picture to be shown.  The Time is in Mili-seconds....1000 = 1 second", "Change Picture Duration Period") & " Ms"
End Sub
