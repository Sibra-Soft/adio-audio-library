VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestAdioPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestAdioLibrary - Playlist Component
'
Option Explicit
Public Sub Setup()
Form_Test.AdioPlaylist.Clear
End Sub
Public Sub TestAddItemToPlaylist()
Assert.AreEqual Form_Test.AdioPlaylist.ListCount, 0, "Playlist must be empty before adding"
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
Assert.AreEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
End Sub
Public Sub TestRemoveItemFromPlaylist()
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
Assert.AreEqual Form_Test.AdioPlaylist.ListCount, 1, "Playlist must contain 1 item"
Form_Test.AdioPlaylist.Remove 1
Assert.AreEqual Form_Test.AdioPlaylist.ListCount, 0, "Playlist should be empty"
End Sub
Public Sub TestGetFirstPlaylistTrack()
Form_Test.AdioPlaylist.AddFile App.Path & "\test.mp3"
Assert.AreEqual App.Path & "\test.mp3", Form_Test.AdioPlaylist.GetTrack(PLS_FIRST).LocalFile
End Sub
