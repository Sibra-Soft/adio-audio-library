VERSION 5.00
Object = "{40F6D89D-D6BF-4EAD-B885-E1869BDF4E31}#36.0#0"; "AdioLibrary.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Test 
   Caption         =   "Adio Test"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin AdioLibrary.AdioCore AdioCore1 
      Left            =   9120
      Top             =   6960
      _ExtentX        =   8493
      _ExtentY        =   873
      Begin AdioLibrary.AdioRecorder AdioRecorder1 
         Left            =   4320
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioTagging AdioTagging1 
         Left            =   3720
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioPlaylist AdioPlaylist1 
         Left            =   3120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         AllowDuplicateItems=   0   'False
      End
      Begin AdioLibrary.AdioCDPlayer AdioCDPlayer1 
         Left            =   2520
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioAudioPeak AdioAudioPeak1 
         Left            =   1920
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioMidiPlayer AdioMidiPlayer1 
         Left            =   1320
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin AdioLibrary.AdioMediaPlayer AdioMediaPlayer1 
         Left            =   720
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Devices "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6480
      TabIndex        =   32
      Top             =   4320
      Width           =   7335
      Begin VB.CommandButton Button_SetMidiDevice 
         Caption         =   "SET MIDI"
         Height          =   375
         Left            =   2520
         TabIndex        =   45
         Top             =   2520
         Width           =   1095
      End
      Begin MSComctlLib.ListView Listview_Devices 
         Height          =   2055
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   7939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CommandButton Button_SetOutDevice 
         Caption         =   "SET OUT"
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Button_SetInDevice 
         Caption         =   "SET IN"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Playlist "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   6480
      TabIndex        =   15
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   5880
         TabIndex        =   27
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add Directory"
         Height          =   375
         Left            =   5880
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Button_AddFileToPlaylist 
         Caption         =   "Add File"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Button_PlaylistQuery 
         Caption         =   "Exec Query"
         Height          =   375
         Left            =   5880
         TabIndex        =   29
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Button_LoadPlaylist 
         Caption         =   "Load Playlist"
         Height          =   375
         Left            =   5880
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Button_RemoveFromPlaylist 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5880
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.ListBox ListBox_Playlist 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Total runtime: 0"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3525
         Width           =   4215
      End
   End
   Begin VB.CommandButton Button_Rewind 
      Caption         =   "Rewind"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Button_Prev 
      Caption         =   "Prev"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " Audio Player "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Timer Timer_Controls 
         Interval        =   10
         Left            =   5640
         Top             =   2400
      End
      Begin VB.CommandButton Button_Record 
         Caption         =   "Record"
         Height          =   375
         Left            =   4800
         TabIndex        =   31
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Button_Stream 
         Caption         =   "Stream"
         Height          =   615
         Left            =   3480
         TabIndex        =   30
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option_CdPlayer 
         Caption         =   "Use Adio CD Player"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   5775
      End
      Begin VB.OptionButton Option_MidiPlayer 
         Caption         =   "Use Adio Midi Player"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   5775
      End
      Begin VB.OptionButton Option_MediaPlayer 
         Caption         =   "Use Adio Audio Player"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Value           =   -1  'True
         Width           =   5775
      End
      Begin VB.CommandButton Button_Playlist 
         Caption         =   "Playlist"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox CheckBox_UsePlaylist 
         Caption         =   "Use Playlist"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton Button_Fade 
         Caption         =   "Fade"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Button_Forward 
         Caption         =   "Forward"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Button_Next 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5280
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Button_Stop 
         Caption         =   "Stop"
         Height          =   615
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Button_Pause 
         Caption         =   "Pause"
         Height          =   615
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Button_Play 
         Caption         =   "Play"
         Height          =   615
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Button_OpenAudioFile 
         Caption         =   "Open"
         Height          =   615
         Left            =   240
         Picture         =   "Form_Test.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   6255
      Begin VB.CommandButton Button_Balance 
         Height          =   195
         Left            =   4540
         TabIndex        =   48
         Top             =   2830
         Width           =   135
      End
      Begin VB.HScrollBar Slider_Balance 
         Height          =   255
         Left            =   3240
         Max             =   1000
         Min             =   -1000
         TabIndex        =   46
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CheckBox CheckBox_MutePlayer 
         Caption         =   "Mute"
         Height          =   255
         Left            =   3240
         TabIndex        =   41
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox CheckBox_MuteMaster 
         Caption         =   "Mute"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar ProgressBar_LeftVolume 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.HScrollBar Slider_Player 
         Height          =   255
         Left            =   3240
         Max             =   100
         TabIndex        =   36
         Top             =   1560
         Value           =   100
         Width           =   2775
      End
      Begin VB.HScrollBar Slider_Master 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   35
         Top             =   1560
         Value           =   100
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar ProgressBar_RightVolume 
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar_MasterVolume 
         Height          =   255
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         Height          =   195
         Left            =   3240
         TabIndex        =   47
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Player:"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Master:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Left            =   3240
         TabIndex        =   24
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   42
      Top             =   3600
      Width           =   6255
      Begin VB.Label Label_Runtime 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   5955
      End
   End
End
Attribute VB_Name = "Form_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AdioAudioPeak1_ChannelAudioLevelChange(leftValue As Integer, rightValue As Integer)
ProgressBar_LeftVolume.Value = leftValue
ProgressBar_RightVolume.Value = rightValue
End Sub

Private Sub AdioAudioPeak1_MasterAudioPeakLevelChange(Value As Integer)
ProgressBar_MasterVolume.Value = Value
End Sub

Private Sub AdioMediaPlayer1_MediaEnded()
Debug.Print "Ended"
End Sub

Private Sub AdioMediaPlayer1_StartPlay()
Call AdioAudioPeak1.SetChannel(AdioMediaPlayer1.Channel)
End Sub

Private Sub AdioMediaPlayer1_Stopped()
Debug.Print "Stopped"
End Sub

Private Sub AdioMediaPlayer1_StreamBuffering(Percent As Integer)
Debug.Print Percent
End Sub

Private Sub AdioMediaPlayer1_StreamTitleChange(Title As String)
Debug.Print "Stream: " & Title
End Sub

Private Sub AdioMidiPlayer1_MediaEnded()
Debug.Print "Midi media ended"
End Sub

Private Sub AdioMidiPlayer1_MidiTrack(Name As String, TrackNr As Integer)
Debug.Print "Midi track (" & TrackNr & "): " & Name
End Sub

Private Sub AdioMidiPlayer1_NewMediaFile(File As String)
Debug.Print "New midi file loaded: " & File
End Sub

Private Sub AdioMidiPlayer1_Playing()
Debug.Print "Start playing midi"
End Sub

Private Sub AdioMidiPlayer1_Ready()
Debug.Print "Midi component is ready"
End Sub

Private Sub AdioPlaylist1_Error(Description As String, code As Long)
Debug.Print Description
End Sub

Private Sub AdioPlaylist1_ListChanged()
Label1.Caption = "Item(s): " & AdioPlaylist1.ListCount & vbNewLine & "Total runtime: " & AdioPlaylist1.GetPlaylistRuntimeString

Dim Item As mdlAdioPlaylistItem

ListBox_Playlist.Clear

For Each Item In AdioPlaylist1.GetList
    ListBox_Playlist.AddItem Item.FileName
Next
End Sub

Private Sub AdioPlaylist1_TrackChanged(track As AdioLibrary.mdlAdioPlaylistItem)
AdioMediaPlayer1.LoadFile track.LocalFile
AdioMediaPlayer1.StartPlay

ListBox_Playlist.Selected(track.Nr - 1) = True
End Sub

Private Sub AdioRecorder1_Error(ErrorCode As Long, ErrorDescription As String)
Debug.Print "RecError: " & ErrorDescription
End Sub

Private Sub Button_AddFileToPlaylist_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .DialogTitle = "Open audio file"
    .Filter = "MPEG Layer 3 (*.mp3)|*.mp3|Midi File (*.mid)|*.mid|Microsoft Wave File (*.wav)|*.wav"
    
    .ShowOpen
    
    AdioPlaylist1.AddFile .FileName
End With

Exit Sub
ErrorHandler:
Select Case Err.Number
    Case 0
    Case cdlCancel
End Select
End Sub

Private Sub Button_Balance_Click()
Slider_Balance.Value = 0
End Sub

Private Sub Button_Fade_Click()
AdioMediaPlayer1.Fade AdioOut
End Sub

Private Sub Button_Forward_Click()
If Option_MediaPlayer.Value Then
    AdioMediaPlayer1.SeekBySeconds AdioForward
ElseIf Option_MidiPlayer.Value Then
    AdioMidiPlayer1.SeekBySeconds AdioForward
End If
End Sub

Private Sub Button_LoadPlaylist_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .DialogTitle = "Open audio file"
    .Filter = "Audiostation Playlist (*.apl)|*.apl|Playlist (.m3u)|*.m3u|ShoutCast Playlist (*.pls)|*.pls|Windows Media Player Playlist (*.wpl)|*.wpl"
    
    .ShowOpen
    
    AdioPlaylist1.LoadPlaylist .FileName, PLAYLIST_APL
End With

Exit Sub
ErrorHandler:
Select Case Err.Number
    Case 0
    Case cdlCancel
End Select
End Sub

Private Sub Button_Next_Click()
Call AdioPlaylist1.GetTrack(PLS_NEXT)
End Sub

Private Sub Button_OpenAudioFile_Click()
On Error GoTo ErrorHandler
With CommonDialog
    .CancelError = True
    .DialogTitle = "Open audio file"
    .Filter = "MPEG Layer 3 (*.mp3)|*.mp3|Midi File (*.mid)|*.mid|Microsoft Wave File (*.wav)|*.wav"
    
    .ShowOpen
    
    If Option_MidiPlayer.Value = True Then
        AdioMidiPlayer1.LoadFile .FileName
    Else
        AdioMediaPlayer1.LoadFile .FileName
    End If
End With

Exit Sub
ErrorHandler:
Select Case Err.Number
    Case 0
    Case cdlCancel
End Select
End Sub

Private Sub Button_Pause_Click()
If Option_MediaPlayer.Value Then
    AdioMediaPlayer1.PausePlay
ElseIf Option_MidiPlayer.Value Then
    AdioMidiPlayer1.PausePlay
End If
End Sub

Private Sub Button_Play_Click()
If Option_MediaPlayer.Value Then
    AdioMediaPlayer1.StartPlay
ElseIf Option_MidiPlayer.Value Then
    AdioMidiPlayer1.StartPlay
End If
End Sub

Private Sub Button_Prev_Click()
Call AdioPlaylist1.GetTrack(PLS_PREV)
End Sub

Private Sub Button_Record_Click()
Call AdioRecorder1.StartRecording
End Sub

Private Sub Button_RemoveFromPlaylist_Click()
AdioPlaylist1.Remove ListBox_Playlist.ListIndex
End Sub

Private Sub Button_Rewind_Click()
If Option_MediaPlayer.Value Then
    AdioMediaPlayer1.SeekBySeconds AdioRewind
ElseIf Option_MidiPlayer.Value Then
    AdioMidiPlayer1.SeekBySeconds AdioRewind
End If
End Sub

Private Sub Button_SetMidiDevice_Click()
AdioMidiPlayer1.InitComponent (Listview_Devices.SelectedItem.SubItems(2) - 1)
End Sub

Private Sub Button_Stop_Click()
If Option_MediaPlayer.Value Then
    AdioMediaPlayer1.StopPlay
ElseIf Option_MidiPlayer.Value Then
    AdioMidiPlayer1.StopPlay
End If
End Sub

Private Sub Button_Stream_Click()
AdioMediaPlayer1.LoadStream "http://somafm.com/secretagent.pls"
End Sub

Private Sub CheckBox_MuteMaster_Click()
Call AdioAudioPeak1.MuteMasterVolume
End Sub

Private Sub CheckBox_MutePlayer_Click()
AdioMediaPlayer1.MuteAudio
End Sub

Private Sub Command1_Click()
AdioPlaylist1.Clear
End Sub

Private Sub Button_PlaylistQuery_Click()
Dim DInput As String

If Button_PlaylistQuery.Caption = "Clear Query" Then
    AdioPlaylist1.ClearQuery
Else
    DInput = InputBox("Enter query to execute", "Query")
    
    If DInput <> vbNullString Then
        Call AdioPlaylist1.ExecQuery(DInput)
    End If
End If
End Sub

Private Sub Form_Load()
Dim AdioDeviceItem As mdlAdioDevice
Dim AdioMidiDeviceItem As mdlAdioMidiDevice

Dim LstItem As ListItem

AdioCore1.Initialize
AdioAudioPeak1.Run

Label1.Caption = "Item(s): 0" & vbNewLine & "Total runtime: 0"

For Each AdioDeviceItem In AdioCore1.GetListOfDevices
    Set LstItem = Listview_Devices.ListItems.Add(, , IIf(AdioDeviceItem.DInput, "INPUT", "OUTPUT"))
        LstItem.SubItems(1) = AdioDeviceItem.dName
Next

For Each AdioMidiDeviceItem In AdioMidiPlayer1.GetListOfMidiDevices
    Set LstItem = Listview_Devices.ListItems.Add(, , "MIDI")
        LstItem.SubItems(1) = AdioMidiDeviceItem.mName
        LstItem.SubItems(2) = AdioMidiDeviceItem.Mid
Next
End Sub

Private Sub ListBox_Playlist_DblClick()
Call AdioPlaylist1.GetTrack(PLS_GOTO, ListBox_Playlist.ListIndex + 1)
End Sub

Private Sub Listview_Devices_DblClick()
Debug.Print Listview_Devices.SelectedItem.Key
End Sub

Private Sub Slider_Balance_Change()
Call AdioMediaPlayer1.SetBalance(Slider_Balance.Value)
End Sub

Private Sub Slider_Master_Scroll()
AdioAudioPeak1.SetMasterVolume Slider_Master.Value
End Sub

Private Sub Slider_Player_Scroll()
AdioMediaPlayer1.SetVolume Slider_Player.Value
End Sub

Private Sub Timer_Controls_Timer()
If AdioPlaylist1.QueryActive Then
    Button_PlaylistQuery.Caption = "Clear Query"
Else
    Button_PlaylistQuery.Caption = "Exec Query"
End If

If Option_MidiPlayer.Value Then
    Label_Runtime.Caption = "Elapsed: " & AdioMidiPlayer1.GetProperties.ElapsedString & " - Remaining: " & AdioMidiPlayer1.GetProperties.RemainingString
Else
    Label_Runtime.Caption = "Elapsed: " & AdioMediaPlayer1.GetProperties.ElapsedString & " - Remaining: " & AdioMediaPlayer1.GetProperties.RemainingString
End If

If CheckBox_UsePlaylist.Value = vbChecked Then
    Button_Prev.Enabled = True
    Button_Next.Enabled = True
Else
    Button_Prev.Enabled = False
    Button_Next.Enabled = False
    
    If ListBox_Playlist.ListCount = 0 Then
        CheckBox_UsePlaylist.Enabled = False
    Else
        CheckBox_UsePlaylist.Enabled = True
    End If
End If
End Sub
