VERSION 5.00
Begin VB.UserControl AdioMediaPlayer 
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6165
   Begin VB.Timer Timer_Playing 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   1350
   End
   Begin VB.Timer Timer_Stream 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1620
      Top             =   1350
   End
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioMediaPlayer.ctx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label_StreamTitle 
      Height          =   285
      Left            =   810
      TabIndex        =   0
      Top             =   1485
      Visible         =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "AdioMediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////
'// FileName        : AdioMediaPlayer.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 28-10-2023
'// Last Modified   : 30-01-2026
'// Copyright       : Sibra-Soft
'// Description     : Usercontrol for audio playback
'////////////////////////////////////////////////////////////////

Option Explicit

'// Private vars
Private MediaChannel As Long
Private StreamEnded As Boolean

'// Public vars
Public State As enumAdioPlayState
Public LoadedFile As String
Public RepeatMode As enumAdioRepeatMode

'// Enums
Public Event Paused()
Public Event Stopped()
Public Event Playing()
Public Event StartPlay()
Public Event MediaEnded()
Public Event NewMediaFile(File As String)
Public Event NewStream()
Public Event Error(Description As String, Code As Long)
Public Event Fading(Progress As Integer)
Public Event StreamBuffering(Percent As Integer)
Public Event StreamTitleChange(Title As String)
Public Function Channel() As Long
Channel = MediaChannel
End Function
'*
'* Set the balance of the speaker audio
'* @param Integer Value: Balance value between -1000 and 1000
'*
Public Sub SetBalance(Value As Integer)
Call modAdio.SetBalance(MediaChannel, Value)
End Sub
Public Function SetDeviceById(id As Long) As Boolean

End Function
Public Function SetDevice(device As mdlAdioDevice) As Boolean

End Function
Public Function LoadStream(StreamUrl As String, Optional ProxyAddress As String) As Boolean
If OpenStreamByUrl(StreamUrl) Then
    Timer_Stream.Enabled = True
    
    State = AdioPlaying
    
    RaiseEvent NewStream
End If
End Function
Public Sub Fade(FadeType As enumAdioFadeType, Optional Duration As Integer = 5)
modAdio.AdioFade MediaChannel, FadeType, Duration
End Sub
Public Sub SetVolume(Value As Integer)
Call modAdio.SetVolume(MediaChannel, Value)
End Sub
Public Function GetVolume() As Integer
GetVolume = modAdio.GetVolume(MediaChannel)
End Function
Public Function MuteAudio() As Boolean
If modAdio.Mute Then
    Call modAdio.AdioMuteOff(MediaChannel)
    Mute = False
Else
    Call modAdio.AdioMuteOn(MediaChannel)
    Mute = True
End If
End Function
Public Sub SeekBySeconds(Direction As enumAdioSeekDirection, Optional Seconds As Integer = 10)
Call modAdio.AdioSeekBySeconds(MediaChannel, Direction, Seconds)
End Sub
Public Sub StartPlay()
Call modAdio.AdioPlay(MediaChannel)

StreamEnded = False
modAdio.State = AdioPlaying
State = AdioPlaying

RaiseEvent StartPlay
RaiseEvent Playing

Timer_Playing.Enabled = True
End Sub
Public Sub StopPlay()
If Not State = AdioPlaying Then: Exit Sub

Call modAdio.AdioStop(MediaChannel)

modAdio.State = AdioStopped
State = AdioStopped

RaiseEvent Stopped

Timer_Stream.Enabled = False
Timer_Playing.Enabled = False
End Sub
Public Sub PausePlay()
Call modAdio.AdioPause(MediaChannel)

modAdio.State = AdioPaused
State = AdioPaused

RaiseEvent Paused

Timer_Playing.Enabled = False
End Sub
Public Function GetProperties() As mdlAdioProperties
Set GetProperties = modAdio.GetProperties(MediaChannel)
End Function
Public Function LoadFile(File As String) As Boolean
Dim Fso As New FileSystemObject

If Not Ext.FileExists(File) Then: RaiseEvent Error("File not found", 100)
If Not CheckFileSupport(File) Then: RaiseEvent Error("File not supported", 110)

Call BASS_ChannelFree(MediaChannel)

' Check the extension
Select Case Fso.GetExtensionName(File)
    Case "flac": MediaChannel = BASS_FLAC_StreamCreateFile(0&, StrPtr(File), 0&, 0&, BASS_SAMPLE_FX)
    Case Else: MediaChannel = BASS_StreamCreateFile(0&, StrPtr(File), 0&, 0&, BASS_SAMPLE_FX)
End Select

If MediaChannel Then
    State = AdioReady
    LoadedFile = File
    
    RaiseEvent NewMediaFile(File)
Else
    RaiseEvent Error("Problem while loading file: " & File, BASS_ErrorGetCode)
End If
End Function

Private Sub Label_StreamTitle_Change()
RaiseEvent StreamTitleChange(Label_StreamTitle.Caption)
End Sub

Private Sub Timer_Playing_Timer()
If GetProperties.RemainingInSeconds <= 0 Then: StreamEnded = True

If StreamEnded = True Then
    State = AdioEnded
    Timer_Playing.Enabled = False
    
    RaiseEvent MediaEnded
Else
    RaiseEvent Playing
End If
End Sub

Private Sub Timer_Stream_Timer()
Call TimerProc

If StreamState = Buffering Then: RaiseEvent StreamBuffering(StreamBufferProgress)

Label_StreamTitle.Caption = modAdioNetRadio.StreamMeta
End Sub

Private Sub UserControl_Resize()
Width = Image_Main.Width
Height = Image_Main.Height
End Sub
