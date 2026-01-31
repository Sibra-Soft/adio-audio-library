Attribute VB_Name = "modAdio"
'///////////////////////////////////////////////////////////////
'// FileName        : modAdio.bas
'// FileType        : Microsoft Visual Basic 6 - Module
'// Author          : Alex van den Berg
'// Created         : 17-08-2023
'// Last Modified   : 31-01-2026
'// Copyright       : Sibra-Soft
'// Description     : Simplified functions for using Bass
'////////////////////////////////////////////////////////////////

Option Explicit

'// Public vars
Public State As enumAdioPlayState
Public Mute As Boolean
'*
'* Set the balance of the speaker audio
'* @param Long Channel: The channel you want to interact with
'* @param Integer Value: Balance value between -1000 and 1000
'*
Public Sub SetBalance(Channel As Long, Value As Integer)
Dim mBalance As Single

mBalance = Value / 1000
Call BASS_ChannelSetAttribute(Channel, BASS_ATTRIB_PAN, mBalance)
End Sub
'*
'* Make a negative number always 0
'*
Private Function DPosNr(number As Long) As Long
If number < 0 Then
    DPosNr = 0
Else
    DPosNr = number
End If
End Function
'*
'* Write MP3 meta tags to a specified MP3 file
'* @param String File: The file you want to change
'* @param String Value: The value you want to write
'* @param enumAdioTags TagType: The tag you want to change
'* @param enumAdioTagVersion AdioTags: The version of the tag you want to write (V1, V2)
'* @return Boolean: Tells if the tags has been changed
'*
Public Function AdioWriteTag(File As String, Value As String, TagType As enumAdioTags, TagVersion As enumAdioTagVersion) As Boolean
Dim OutputTags As ID3Tag

OutputTags.Artist = AdioReadTag(File, tArtist, TagVersion)
OutputTags.Title = AdioReadTag(File, tTitle, TagVersion)
OutputTags.Album = AdioReadTag(File, tAlbum, TagVersion)
OutputTags.OrigArtist = AdioReadTag(File, tAlbumArtist, TagVersion)
OutputTags.Composer = AdioReadTag(File, tComposer, TagVersion)
OutputTags.Copyright = AdioReadTag(File, tCopyright, TagVersion)
OutputTags.Genre = AdioReadTag(File, tGenre, TagVersion)
OutputTags.SongYear = AdioReadTag(File, tYear, TagVersion)
OutputTags.TrackNr = AdioReadTag(File, tTrackNumber, TagVersion)

Select Case TagType
    Case enumAdioTags.tArtist: OutputTags.Artist = Value
    Case enumAdioTags.tTitle: OutputTags.Title = Value
    Case enumAdioTags.tAlbum: OutputTags.Album = Value
    Case enumAdioTags.tAlbumArtist: OutputTags.OrigArtist = Value
    Case enumAdioTags.tComposer: OutputTags.Composer = Value
    Case enumAdioTags.tCopyright: OutputTags.Copyright = Value
    Case enumAdioTags.tGenre: OutputTags.Genre = Value
    Case enumAdioTags.tYear: OutputTags.SongYear = Value
    Case enumAdioTags.tTrackNumber: OutputTags.TrackNr = Value
End Select

If TagVersion = v1 Then
    AdioWriteTag = TagLib.WriteID3v1(File, OutputTags, True)
Else
    AdioWriteTag = TagLib.WriteID3v2(File, OutputTags, VERSION_2_4, True)
End If
End Function
'*
'* Get file properties of a specified file
'* @param String File: The file to get the property of
'* @param enumAdioProperty PropertyType: The property you want to get
'* @return String: The value of the property
'*
Public Function AdioReadAudioProperty(File As String, PropertyType As enumAdioProperty) As String
Dim OutputInfo As MPEGInfo
Dim Fso As New FileSystemObject

If File = vbNullString Then: Exit Function

If Fso.GetExtensionName(File) <> "mp3" Then
    ' When it's not a MP3 file
    Select Case PropertyType
        Case enumAdioProperty.pDurationInSeconds: AdioReadAudioProperty = modFileProp.GetFileDurationInSeconds(File)
        Case enumAdioProperty.pDurationString: AdioReadAudioProperty = Helpers.SecondsToTimeSerial(modFileProp.GetFileDurationInSeconds(File), SmallTimeSerial)
        Case enumAdioProperty.pFileSize: AdioReadAudioProperty = FileLen(File)
    End Select
Else
    ' When it's a MP3 file
    TagLib.ReadMPEGInfo File, OutputInfo
    
    Select Case PropertyType
        Case enumAdioProperty.pDurationInSeconds: AdioReadAudioProperty = modFileProp.GetFileDurationInSeconds(File)
        Case enumAdioProperty.pBitrate: AdioReadAudioProperty = OutputInfo.Bitrate
        Case enumAdioProperty.pFileSize: AdioReadAudioProperty = FileLen(File)
        Case enumAdioProperty.pDurationString: AdioReadAudioProperty = Helpers.SecondsToTimeSerial(modFileProp.GetFileDurationInSeconds(File), SmallTimeSerial)
        Case enumAdioProperty.pChannels: AdioReadAudioProperty = OutputInfo.ChannelMode
        Case enumAdioProperty.pFrequency: AdioReadAudioProperty = OutputInfo.Frequency
    End Select
End If
End Function
'*
'* Gets MP3 meta tags of a specified file
'* @param String File: The file you want to get tags of
'* @param enumAdioTags TagType: The tag you want to get
'* @param enumAdioTagVersion TagVersion: The version of the tag you want to get (V1, V2)
'* @return String: The value of the tag
'*
Public Function AdioReadTag(File As String, TagType As enumAdioTags, TagVersion As enumAdioTagVersion) As String
Dim Output As ID3Tag

If File = vbNullString Then: Exit Function

If TagVersion = v2 Then
    TagLib.ReadID3v2 File, Output
Else
    TagLib.ReadID3v1 File, Output
End If

Select Case TagType
    Case enumAdioTags.tArtist: AdioReadTag = Output.Artist
    Case enumAdioTags.tTitle: AdioReadTag = Output.Title
    Case enumAdioTags.tAlbum: AdioReadTag = Output.Album
    Case enumAdioTags.tAlbumArtist: AdioReadTag = Output.OrigArtist
    Case enumAdioTags.tComposer: AdioReadTag = Output.Composer
    Case enumAdioTags.tCopyright: AdioReadTag = Output.Copyright
    Case enumAdioTags.tGenre: AdioReadTag = Output.Genre
    Case enumAdioTags.tYear: AdioReadTag = Output.SongYear
    Case enumAdioTags.tTrackNumber: AdioReadTag = Output.TrackNr
End Select
End Function
'*
'* Get the current channel volume
'* @param Long Channel: The channel you want to interact with
'*
Public Function GetVolume(Channel As Long) As Integer
Dim mVolume As Single
Call BASS_ChannelGetAttribute(Channel, BASS_ATTRIB_VOL, mVolume)

GetVolume = mVolume * 100
End Function
'*
'* Set the current channel volume
'* @param Long Channel: The channel you want to interact with
'* @param Integer Value: 0 to 100 meaning 0 no sound, 100 max volume
'*
Public Sub SetVolume(Channel As Long, Value As Integer)
Dim mVolume As Single

mVolume = Value / 100
Call BASS_ChannelSetAttribute(Channel, BASS_ATTRIB_VOL, mVolume)
End Sub
'*
'* Gets player properties (songlength, duration, etc.)
'* @param Long Channel: The channel you want to interact with
'* @return mdlAdioProperties: The properties of the current player
'*
Public Function GetProperties(Channel As Long) As mdlAdioProperties
Dim ReturnValue As New mdlAdioProperties
Dim SongLengthInBytes As Long
Dim SongPosInBytes As Long

SongLengthInBytes = BASS_ChannelGetLength(Channel, BASS_POS_BYTE)
SongPosInBytes = BASS_ChannelGetPosition(Channel, BASS_POS_BYTE)

' Get properties in seconds
ReturnValue.DurationInSeconds = BASS_ChannelBytes2Seconds(Channel, SongLengthInBytes)
ReturnValue.ElapsedInSeconds = DPosNr(BASS_ChannelBytes2Seconds(Channel, SongPosInBytes))
ReturnValue.RemainingInSeconds = DPosNr(ReturnValue.DurationInSeconds - ReturnValue.ElapsedInSeconds)

' Get properties from seconds to string
ReturnValue.DurationString = Helpers.SecondsToTimeSerial(ReturnValue.DurationInSeconds, SmallTimeSerial)
ReturnValue.ElapsedString = Helpers.SecondsToTimeSerial(ReturnValue.ElapsedInSeconds, SmallTimeSerial)
ReturnValue.RemainingString = Helpers.SecondsToTimeSerial(ReturnValue.RemainingInSeconds, SmallTimeSerial)

Set GetProperties = ReturnValue
End Function
'*
'* Stop the player
'* @param Long Channel: The channel you want to interact with
'*
Public Sub AdioStop(Channel As Long)
Call BASS_ChannelStop(Channel)
End Sub
'*
'* Start the player
'* @param Long Channel: The channel you want to interact with
'*
Public Sub AdioPlay(Channel As Long)
Dim SyncHandle As Long

If State = AdioPaused Then
    Call BASS_ChannelPlay(Channel, 0&)
Else
    Call BASS_ChannelPlay(Channel, 1&)
End If
End Sub
'*
'* Pause the player
'* @param Long Channel: The channel you want to interact with
'*
Public Sub AdioPause(Channel As Long)
Call BASS_ChannelPause(Channel)
End Sub
'*
'* Seek the player forward or backwards a specified number of seconds
'* @param Long Channel: The channel you want to interact with
'* @param enumAdioSeekDirection Direction: The direction you want to seek (rewind, forward)
'* @param Integer Seconds: The amount of seconds you want to seek (default = 10)
'*
Public Sub AdioSeekBySeconds(Channel As Long, Direction As enumAdioSeekDirection, Optional Seconds As Integer = 10)
Dim Pos As Long
Dim AtPos As Long

Pos = BASS_ChannelGetPosition(Channel, BASS_POS_BYTE)
AtPos = BASS_ChannelSeconds2Bytes(Channel, Seconds)

Select Case Direction
    Case enumAdioSeekDirection.AdioForward: Call BASS_ChannelSetPosition64(Channel, Pos + AtPos, 0&, BASS_POS_BYTE)
    Case enumAdioSeekDirection.AdioRewind: Call BASS_ChannelSetPosition64(Channel, Pos - AtPos, 0&, BASS_POS_BYTE)
End Select
End Sub
'*
'* Fade the current player
'* @param Long Channel: The channel you want to interact with
'* @param enumAdioFadeType FadeType: The type of fade (IN or OUT)
'* @param Integer Duration: The duration of the fade in seconds (default = 10)
'*
Public Sub AdioFade(Channel As Long, FadeType As enumAdioFadeType, Optional Duration As Integer = 5)
Select Case FadeType
    Case enumAdioFadeType.AdioIn: Call BASS_ChannelSlideAttribute(Channel, BASS_ATTRIB_VOL, 1, Duration * 1000)
    Case enumAdioFadeType.AdioOut: Call BASS_ChannelSlideAttribute(Channel, BASS_ATTRIB_VOL, -1, Duration * 1000)
End Select
End Sub
'*
'* Mute the current player
'* @param Long Channel: The channel you want to interact with
'*
Public Sub AdioMuteOn(Channel As Long)
Call BASS_ChannelSetAttribute(Channel, BASS_ATTRIB_VOL, 0)
Mute = True
End Sub
'*
'* Disable mute of the current player
'* @param Long Channel: The channel you want to interact with
'*
Public Sub AdioMuteOff(Channel As Long)
Call BASS_ChannelSetAttribute(Channel, BASS_ATTRIB_VOL, 1)
Mute = False
End Sub
