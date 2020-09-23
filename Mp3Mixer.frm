VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F081C29F-491A-11D4-9BEC-AB56882D7A01}#1.0#0"; "PICSCROLLER.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mp3j Testing BETA"
   ClientHeight    =   7050
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10275
   Begin VB.PictureBox Picture1 
      Height          =   3705
      Left            =   1830
      ScaleHeight     =   3645
      ScaleWidth      =   7260
      TabIndex        =   33
      Top             =   1170
      Width           =   7320
      Begin VB.Label Label1 
         Caption         =   $"Mp3Mixer.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   7050
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   1620
      Left            =   30
      TabIndex        =   32
      Top             =   690
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   2858
      ButtonWidth     =   1296
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Directory"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+ Dir"
      Height          =   495
      Left            =   300
      TabIndex        =   20
      Top             =   1470
      Width           =   405
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   240
      Left            =   285
      TabIndex        =   8
      Top             =   1245
      Width           =   405
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   870
      ScaleHeight     =   360
      ScaleWidth      =   9060
      TabIndex        =   29
      Top             =   4995
      Width           =   9060
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Deck 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   90
         TabIndex        =   31
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Deck 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4500
         TabIndex        =   30
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   75
      ScaleHeight     =   360
      ScaleWidth      =   9855
      TabIndex        =   26
      Top             =   165
      Width           =   9855
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Play List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   795
         TabIndex        =   27
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   7155
      Pattern         =   "*.mp3;*.wav;*.class"
      TabIndex        =   25
      Top             =   8370
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1485
      Left            =   855
      TabIndex        =   12
      Top             =   5370
      Width           =   4440
      Begin VB.PictureBox Scroll3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2595
         ScaleHeight     =   255
         ScaleWidth      =   1740
         TabIndex        =   21
         Top             =   195
         Width           =   1740
         Begin PicScroller.PicScroll PicScroll3 
            Height          =   225
            Left            =   45
            TabIndex        =   22
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   397
            Bar             =   "Mp3Mixer.frx":008B
            BackColor       =   14737632
            Max             =   2500
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "=>"
         Height          =   255
         Left            =   3870
         TabIndex        =   18
         Top             =   465
         Width           =   405
      End
      Begin PicScroller.PicScroll PicScroll1 
         Height          =   225
         Left            =   90
         TabIndex        =   13
         Top             =   1140
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   397
         Bar             =   "Mp3Mixer.frx":036F
         BackColor       =   8421504
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Track"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   165
         TabIndex        =   17
         Top             =   450
         Width           =   3630
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   195
         Width           =   2325
      End
      Begin VB.Label Label11 
         Caption         =   "(remaining 0 min 0 sec)"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2475
         TabIndex        =   15
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "0 min 00 sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   855
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1485
      Left            =   5340
      TabIndex        =   2
      Top             =   5370
      Width           =   4575
      Begin VB.CommandButton Command6 
         Caption         =   "<="
         Height          =   255
         Left            =   4050
         TabIndex        =   19
         Top             =   465
         Width           =   405
      End
      Begin VB.PictureBox Scroll4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1740
         TabIndex        =   23
         Top             =   180
         Width           =   1740
         Begin PicScroller.PicScroll PicScroll4 
            Height          =   225
            Left            =   -15
            TabIndex        =   24
            Top             =   30
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   397
            Bar             =   "Mp3Mixer.frx":0653
            BackColor       =   14737632
            Max             =   2500
         End
      End
      Begin PicScroller.PicScroll PicScroll2 
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   397
         Bar             =   "Mp3Mixer.frx":0937
         BackColor       =   8421504
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Track"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   135
         TabIndex        =   7
         Top             =   435
         Width           =   3810
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Artist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   150
         TabIndex        =   6
         Top             =   195
         Width           =   2430
      End
      Begin VB.Label Label4 
         Caption         =   "0 min 00 sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   165
         TabIndex        =   5
         Top             =   855
         Width           =   1590
      End
      Begin VB.Label Label12 
         Caption         =   "(remaining 0 min 0 sec)"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2565
         TabIndex        =   4
         Top             =   840
         Width           =   1875
      End
   End
   Begin VB.Timer Tmr12 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2955
      Top             =   8295
   End
   Begin VB.Timer Tmr21 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3420
      Top             =   8295
   End
   Begin VB.Timer Tmr1End 
      Interval        =   1000
      Left            =   2535
      Top             =   8310
   End
   Begin VB.Timer Tmr2End 
      Interval        =   1000
      Left            =   4305
      Top             =   8325
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   3855
      Top             =   8325
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   8220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mp3Mixer.frx":0C1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mp3Mixer.frx":11B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mp3Mixer.frx":1753
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstPlay 
      Height          =   4380
      Left            =   855
      TabIndex        =   10
      Top             =   570
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7726
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Track"
         Object.Width           =   3880
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LstHistory 
      Height          =   4380
      Left            =   5340
      TabIndex        =   11
      Top             =   570
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   7726
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Track"
         Object.Width           =   3880
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "|>"
      Height          =   285
      Left            =   285
      TabIndex        =   9
      Top             =   975
      Width           =   405
   End
   Begin MediaPlayerCtl.MediaPlayer am1 
      Height          =   630
      Left            =   6495
      TabIndex        =   1
      Top             =   8265
      Visible         =   0   'False
      Width           =   315
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   30
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   0
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
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer am2 
      Height          =   660
      Left            =   6075
      TabIndex        =   0
      Top             =   8250
      Visible         =   0   'False
      Width           =   315
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   30
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
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuLstHistory 
      Caption         =   "LstHistory"
      Visible         =   0   'False
      Begin VB.Menu mnuSendToPlaylist 
         Caption         =   "Se&t as Next"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuRemoveFromHistory 
         Caption         =   "Re&move From History"
      End
   End
   Begin VB.Menu mnuLstPLay 
      Caption         =   "LstPLay"
      Visible         =   0   'False
      Begin VB.Menu mnuSetAsNext 
         Caption         =   "Set As &Next"
      End
      Begin VB.Menu mnuPlay2 
         Caption         =   "P&lay"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "R&emove From Playlist"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pimp As Integer, pump As Integer, foop As String, loopy As Boolean
Dim a1
Dim a2


Public Deck As String
Public PS4Value As Integer
Public PS3Value As Integer



Private Sub Command2_Click()
Command5_Click
End Sub

Private Sub Command3_Click()
Set id = New Id3


On Error GoTo Er:

    Dim ReturnValue As String 'Keeps up the return
    Dim WithFiles As Long 'Just for this project, to add browsing with files or not
    Dim SelectedFolder
   
    ReturnValue = BrowseForFolder(Me.hwnd, "Choose a directory to add:", WithFiles, RecycleBin)
    If ReturnValue <> "" Then
      SelectedFolder = ReturnValue: GoTo 123
    Else
Exit Sub
    End If
    
123:
    
    If Right(SelectedFolder, 1) = "\" Then
    File1.Path = SelectedFolder

    Else
    File1.Path = SelectedFolder & "\"
    End If
    

    
id.ClearAll
File1.ListIndex = -1
Dim aa
Dim bb
Dim i
Dim l

For i = 1 To File1.ListCount

File1.ListIndex = File1.ListIndex + 1

id.Filename = File1.Path & "\" & File1.Filename

ab = RTrim(id.Artist)
bc = RTrim(id.Title)

Set r = LstPlay.ListItems.Add(, File1.Path & "\" & File1.Filename, bc, , 1)
r.SubItems(1) = ab
DoEvents
Next i
Exit Sub
Er:
MsgBox Err.Description, vbCritical, "File/Path Error"
End Sub




Private Sub ClearDeck1()
am1.Filename = ""
Label7.Caption = "Artist"
Label8.Caption = "Track"
PicScroll1.Value = 0
End Sub

Private Sub ClearDeck2()
am2.Filename = ""
Label9.Caption = "Track"
Label10.Caption = "Artist"
PicScroll2.Value = 0
End Sub


Private Sub LoadDeck1()
Deck = "am2"

Set am = LstPlay.ListItems(1)
amTitle = am.Text
amArtist = am.SubItems(1)
amKey = am.Key

am1.Filename = amKey

Tmr21.Enabled = True 'fade 2 to 1 and play
PS4Value = PicScroll4.Value - 2500
Scroll4.Enabled = False

On Error Resume Next

Set s = LstHistory.ListItems.Add(1, amKey, amTitle, , 1)
s.SubItems(1) = amArtist

LstPlay.ListItems.Remove amKey

'''''''''''''''''''''''''

Label7.Caption = amArtist
Label8.Caption = amTitle
End Sub

Private Sub LoadDeck2()

Deck = "am1"



Set am = LstPlay.ListItems(1)
amTitle = am.Text
amArtist = am.SubItems(1)
amKey = am.Key

am2.Filename = amKey
Tmr12.Enabled = True 'fade 1 to 2 and play 2
PS3Value = PicScroll3.Value - 2500
Scroll3.Enabled = False

On Error Resume Next

Set l = LstHistory.ListItems.Add(1, amKey, amTitle, , 1)
l.SubItems(1) = amArtist

LstPlay.ListItems.Remove amKey


'''''''''''''''''''''''''''


Label10.Caption = amArtist
Label9.Caption = amTitle
End Sub


Private Sub Fade1to2()
a1 = am1.Volume
a2 = am2.Volume
Timer1.Enabled = True
am1.Volume = a1
Timer4.Enabled = True
End Sub

Private Sub Fade2to1()
a1 = am1.Volume
a2 = am2.Volume
Timer2.Enabled = True
am2.Volume = a2
Timer5.Enabled = True
End Sub



Private Sub am1_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
PicScroll1.Max = am1.Duration
PicScroll1.Value = am1.CurrentPosition
End Sub

Private Sub AddFilesToList()
Dim r As Node
Dim M As Node
File1.ListIndex = -1
Dim aa
Dim bb
Dim i
Dim l



Set r = TreeView1.Nodes.Add(, , "Top", "My Music", 1)

For i = 1 To File1.ListCount

File1.ListIndex = File1.ListIndex + 1

Set id = New Id3
id.Filename = Path & "\" & File1.Filename

aa = RTrim(id3Info.Artist)
bb = RTrim(id3Info.Title)
cc = RTrim(id3Info.Album)
dd = RTrim(id3Info.sYear)


On Error Resume Next
'Set l = ListView1.ListItems.Add(, , aa, , 1)
Set aaa = TreeView1.Nodes.Add(r, tvwChild, Trim(aa), aa, 3) ' Artist Name

hell:
  'l.SubItems(1) = bb

DoEvents
Next i
r.Expanded = True
End Sub

Private Sub Command1_Click()
Tmr12.Enabled = True

PS3Value = PicScroll3.Value - 2500
Scroll3.Enabled = False

Deck = "am2"
Command5_Click
End Sub

Private Sub Command4_Click()
Set id = New Id3






Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo e_Trap
    
    FileDialog.sFilter = "MP3 Music Files (*.mp3)" & Chr$(0) & "*.mp3" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Add Files To Playlist"
    sOpen = ShowOpen(Me.hwnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        FileList = vbCr
        
        For Count = 1 To sOpen.nFilesSelected
            'FileList = FileList & sOpen.sFiles(Count) & vbCr
        
id.Filename = sOpen.sFiles(Count)
Set r = LstPlay.ListItems.Add(, id.Filename, id.Title, , 1)
r.SubItems(1) = id.Artist
        
        
        Next Count
        

    End If
    Exit Sub
e_Trap:
    Exit Sub
    Resume
End Sub

Private Sub Command5_Click()
If LstPlay.ListItems.Count = 0 Then Exit Sub

Command1.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
mnuPlay2 = False
mnuPlay.Enabled = False

If Deck = "am2" Then
Command6.Enabled = False
LoadDeck2
Exit Sub
End If

If Deck = "am1" Then
Command1.Enabled = True
LoadDeck1
Exit Sub
End If

End Sub

Private Sub Command6_Click()
Tmr21.Enabled = True

PS4Value = PicScroll4.Value - 2500
Scroll4.Enabled = False

Deck = "am1"
Command5_Click
End Sub













Private Sub Command8_Click()
AddFilesToList
End Sub



Private Sub Form_Load()

'EnhListView_Toggle_FlatColumnHeaders frmMain, LstPlay
'EnhListView_Toggle_FlatColumnHeaders frmMain, LstHistory


LstPlay.ColumnHeaders(1).Width = LstPlay.Width / 1.7
LstPlay.ColumnHeaders(2).Width = LstPlay.Width - (LstPlay.ColumnHeaders(1).Width + 58)

LstHistory.ColumnHeaders(1).Width = LstHistory.Width / 1.7
LstHistory.ColumnHeaders(2).Width = LstHistory.Width - (LstHistory.ColumnHeaders(1).Width + 58)



PicScroll3.Value = PicScroll3.Max
PicScroll4.Value = PicScroll4.Max

SetTimer Me.hwnd, 0, 1000, AddressOf UpdateBars
SetTimer Me.hwnd, 1, 200, AddressOf UpdateTimes

Deck = "am2"

End Sub



Private Sub Form_Unload(Cancel As Integer)
KillTimer Me.hwnd, 0
End Sub



Private Sub HScroll3_Change()
Dim pim, sha
sha = HScroll3.Value - 2500
am2.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll3.min
foo = HScroll3.Value

hell:
Exit Sub
End Sub



Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
'PopupMenu mnuMain
Else
End If
End Sub





Private Sub LstHistory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
On Error Resume Next
ListViewDelSelectedItems frmMain, LstHistory
End If
End Sub

Private Sub LstHistory_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If LstHistory.ListItems.Count = 0 Then
Exit Sub
End If

If Button = 2 Then
PopupMenu mnuLstHistory
End If
End Sub

Private Sub LstPlay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
On Error Resume Next
ListViewDelSelectedItems frmMain, LstPlay
End If
End Sub

Private Sub LstPlay_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If LstPlay.ListItems.Count = 0 Then
Exit Sub
End If

If Button = 2 Then
PopupMenu mnuLstPLay
End If
End Sub







Private Sub LstPlay_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
ListViewGetSelectedItems frmMain, LstHistory
End Sub

Private Sub mnuDelete_Click()
On Error Resume Next
ListViewDelSelectedItems frmMain, LstPlay
End Sub



Private Sub mnuPlay_Click()
ListViewGetSelectedItems frmMain, LstHistory
Command5_Click
End Sub

Private Sub mnuPlay2_Click()
ListViewSetAsNext frmMain, LstPlay
Command5_Click
End Sub

Private Sub mnuRemoveFromHistory_Click()
On Error Resume Next
ListViewDelSelectedItems frmMain, LstHistory
End Sub

Private Sub mnuSendToPlaylist_Click()
ListViewGetSelectedItems frmMain, LstHistory
End Sub

Private Sub mnuSetAsNext_Click()
ListViewSetAsNext frmMain, LstPlay
End Sub

Private Sub PicScroll1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
KillTimer Me.hwnd, 0
End Sub

Private Sub PicScroll1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
SetTimer Me.hwnd, 0, 1, AddressOf UpdateBars
am1.CurrentPosition = PicScroll1.Value
End Sub

Private Sub PicScroll2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
KillTimer Me.hwnd, 0
End Sub

Private Sub PicScroll2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
SetTimer Me.hwnd, 0, 1, AddressOf UpdateBars
am2.CurrentPosition = PicScroll2.Value
End Sub





Private Sub PicScroll3_Change()
Dim pim, sha
sha = PicScroll3.Value - 2500
am1.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = PicScroll3.min
foo = PicScroll3.Value
hell:
Exit Sub
End Sub

Private Sub PicScroll3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim foo As Integer, poo As Integer
On Error GoTo hell
Dim pim, sha
sha = PicScroll3.Value - 2500
am1.Volume = sha
poo = PicScroll3.min
foo = PicScroll3.Value

hell:
Exit Sub
End Sub

Private Sub PicScroll4_Change()
Dim pim, sha
sha = PicScroll4.Value - 2500
am2.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = PicScroll4.min
foo = PicScroll4.Value

hell:
Exit Sub
End Sub

Private Sub PicScroll4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim foo As Integer, poo As Integer
On Error GoTo hell
Dim pim, sha
sha = PicScroll4.Value - 2500
am2.Volume = sha
poo = PicScroll4.min
foo = PicScroll4.Value

hell:
Exit Sub
End Sub







Private Sub Picture23_Click()
On Error GoTo hell
am2.Stop
hell:
Exit Sub
End Sub


Private Sub Picture25_Click()
On Error GoTo hell:
If Pause2 = True Then

Pause2 = False
End If
am2.Play
Deck = "am2"
hell:
Exit Sub
End Sub

Private Sub Picture3_Click()
On Error GoTo hell:
If Pause = True Then
Pause = False
End If
am1.Play
Deck = "am1"
hell:
Exit Sub
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'BitBlt Picture3.hdc, 0, 0, 50, 50, Picture2.hdc, 31, 0, SRCCOPY
Picture3.Refresh
End Sub







Private Sub Picture5_Click()
On Error GoTo hell
am1.Stop
hell:
Exit Sub
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'BitBlt Picture5.hdc, 0, 0, 50, 50, Picture2.hdc, 154.5, 0, SRCCOPY
Picture5.Refresh
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'BitBlt Picture5.hdc, 0, 0, 50, 50, Picture2.hdc, 123.7, 0, SRCCOPY
Picture5.Refresh
End Sub







Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'BitBlt Picture8.hdc, 0, 0, 50, 50, Picture2.hdc, 339, 0, SRCCOPY
Picture8.Refresh
End Sub

Private Sub Picture8_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'BitBlt Picture8.hdc, 0, 0, 50, 50, Picture2.hdc, 309, 0, SRCCOPY
Picture8.Refresh
End Sub

Private Sub Picture9_Click()
If loopy = False Then
'BitBlt Picture9.hdc, 0, 0, 50, 50, Picture10.hdc, 31, 0, SRCCOPY
Picture9.Refresh
loopy = True
Else
'BitBlt Picture9.hdc, 0, 0, 50, 50, Picture10.hdc, 0, 0, SRCCOPY
Picture9.Refresh
loopy = False
End If
End Sub





















Private Sub Timer1_Timer()

End Sub

Private Sub Timer2_Timer()
Tmr1End.Enabled = True
Tmr2End.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Tmr12_Timer()
On Error Resume Next
frmMain.am1.Volume = frmMain.am1.Volume - 100

If frmMain.am1.Volume < (frmMain.PicScroll3.Value / 2500) / 2 Then
frmMain.am2.Play
End If

If frmMain.am1.Volume < -5000 Then
Tmr12.Enabled = False
frmMain.am1.Stop

frmMain.Command1.Enabled = True
frmMain.Command6.Enabled = True
frmMain.Command5.Enabled = True
mnuPlay2 = True
mnuPlay.Enabled = True
Command6.Enabled = True
Command1.Enabled = False

ClearDeck1

frmMain.am1.Volume = PS3Value
Scroll3.Enabled = True






End If
End Sub

Private Sub Tmr1End_Timer()
If frmMain.am1.PlayState = 6 Then Exit Sub
If frmMain.am1.PlayState = 0 Then Exit Sub

If frmMain.am1.CurrentPosition > frmMain.am1.Duration - 7 Then

Command5_Click

Tmr1End.Enabled = False

Timer2.Enabled = True
End If
End Sub

Private Sub Tmr21_Timer()

On Error Resume Next
frmMain.am2.Volume = frmMain.am2.Volume - 100

If frmMain.am2.Volume < (frmMain.PicScroll4.Value / 2500) / 2 Then
frmMain.am1.Play
End If

If frmMain.am2.Volume < -5000 Then
Tmr21.Enabled = False

frmMain.am2.Stop


frmMain.Command1.Enabled = True
frmMain.Command6.Enabled = True
frmMain.Command5.Enabled = True
mnuPlay2 = True
mnuPlay.Enabled = True
Command6.Enabled = False
Command1.Enabled = True

ClearDeck2

frmMain.am2.Volume = PS4Value
Scroll4.Enabled = True


End If
End Sub

Private Sub Tmr2End_Timer()
If am2.PlayState = 6 Then Exit Sub
If am2.PlayState = 0 Then Exit Sub

If am2.CurrentPosition > am2.Duration - 7 Then
Command5_Click

Tmr2End.Enabled = False



Timer2.Enabled = True
End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
Command5_Click
End If
If Button.Index = 2 Then
Command4_Click
End If
If Button.Index = 3 Then
Command3_Click
End If
End Sub

Private Sub TreeView1_Click()
On Error Resume Next
ListView1.ListItems.Clear
File1.ListIndex = -1

Picture1.Visible = True

For i = 1 To File1.ListCount - 1
File1.ListIndex = File1.ListIndex + 1


'getId3 File1.Path & "\" & File1.Filename

aaa = RTrim(id3Info.Artist)
bb = RTrim(id3Info.Title)
cc = RTrim(id3Info.Album)
dd = RTrim(id3Info.sYear)

If aaa = TreeView1.SelectedItem.Text Then
Set x = ListView1.ListItems.Add(, File1.Path & "\" & File1.Filename, bb)
x.SubItems(1) = cc
x.SubItems(2) = dd

Else
End If
DoEvents
Next i

Picture1.Visible = False
End Sub

'''''''''''''''''
