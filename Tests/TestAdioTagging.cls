VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - Tagging Component
'
Option Explicit

Private strArtist As String
Private strTitle As String
Private strYear As String
Private strGenre As String
Private strAlbum As String
Private strDurationInSeconds As String
Public Sub Setup()
With Form_Test.AdioTagging
    strArtist = .ReadTag(App.Path & "\test.mp3", tArtist, v2)
    strTitle = .ReadTag(App.Path & "\test.mp3", tTitle, v2)
    strYear = .ReadTag(App.Path & "\test.mp3", tYear, v2)
    strGenre = .ReadTag(App.Path & "\test.mp3", tGenre, v2)
    strAlbum = .ReadTag(App.Path & "\test.mp3", tAlbum, v2)
    
    strDurationInSeconds = .ReadProperty(App.Path & "\test.mp3", pDurationInSeconds)
End With
End Sub
Public Sub TestDurationInSecondsPropertyRead()
Assert.AreEqual "0", strDurationInSeconds, "Wrong duration amount found"
End Sub
Public Sub TestAlbumTagRead()
Assert.AreEqual "My Suno Songs", strAlbum, "Wrong album found"
End Sub
Public Sub TestGenreTagRead()
Assert.AreEqual "Country", strGenre, "Wrong genre found"
End Sub
Public Sub TestArtistTagRead()
Assert.AreEqual "Alex", strArtist, "Wrong artist found"
End Sub
Public Sub TestTitleTagRead()
Assert.AreEqual "Neon Rodeo", strTitle, "Wrong title found"
End Sub
Public Sub TestYearTagRead()
Assert.AreEqual "2026", strYear, "Wrong year found"
End Sub
