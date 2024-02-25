VERSION 5.00
Begin VB.UserControl AdioPlaylist 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image_Main 
      Height          =   480
      Left            =   0
      Picture         =   "AdioPlaylist.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "AdioPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////
'// FileName        : AdioPlaylist.ctl
'// FileType        : Microsoft Visual Basic 6 - Usercontrol
'// Author          : Alex van den Berg
'// Created         : 16-08-2023
'// Last Modified   : 27-10-2023
'// Copyright       : Sibra-Soft
'// Description     : Playlist component
'////////////////////////////////////////////////////////////////

Option Explicit

'// Enums
Public Enum enumAdioPlaylistRepeatMode
    PLS_NO_REPEAT
    PLS_REPEAT
    PLS_SHUFFLE
End Enum

Public Enum enumAdioPlaylistDirection
    PLS_NEXT
    PLS_PREV
    PLS_FIRST
    PLS_LAST
    PLS_GOTO
End Enum

Public Enum enumAdioPlaylistType
    PLAYLIST_WPL
    PLAYLIST_M3U
    PLAYLIST_PLS
    PLAYLIST_APL
End Enum

'// Private vars
Private CurTrack As Integer
Private CurList As New Collection
Private TempList As New Collection
Private AddingMultipleFiles As Boolean

'// Public vars
Public RepeatMode As enumAdioPlaylistRepeatMode
Public QueryActive As Boolean

'// Events
Public Event ListChanged()
Public Event ListLoadFinished()
Public Event ListSaveFinished()
Public Event TrackRemoved(track As mdlAdioPlaylistItem)
Public Event TrackChanged(track As mdlAdioPlaylistItem)
Public Event NewFileAdded(item As mdlAdioPlaylistItem, position As Long)
Public Event Error(Description As String, Code As Long)
Public Event Progress(item As Long, TotalItems As Long, Finished As Boolean)

'Default Property Values:
Const m_def_AllowDuplicateItems = 1

'Property Variables:
Dim m_AllowDuplicateItems As Boolean
'*
'* Clear the last executed query
'*
Public Sub ClearQuery()
QueryActive = False
RaiseEvent ListChanged
End Sub
'*
'* Checks if the specified file already exists within the playlist
'* @param String File: The full file path of the file
'* @return mdlAdioPlaylistItem: The item of the playlist that was found as duplicate
'*
Private Function CheckIfItemExists(File As String) As Variant
Dim PlaylistItem As mdlAdioPlaylistItem

For Each PlaylistItem In CurList
    If PlaylistItem.LocalFile = File Then
        Set CheckIfItemExists = PlaylistItem
        Exit Function
    End If
Next

Set CheckIfItemExists = Nothing
End Function
'*
'* Remove a item from the current playlist
'* @param Integer Index: The index of the item you want to delete
'*
Public Sub Remove(Index As Long)
RaiseEvent TrackRemoved(CurList(Index))

CurList.Remove Index

RaiseEvent ListChanged
End Sub
'*
'* Get the current amount of items added to your playlist
'* @return Long: The amount of items added to the playlist
'*
Public Function ListCount() As Long
ListCount = CurList.count
End Function
'*
'* Add multiple files to the current playlist, using a newline-separated string
'* @param String Files: Newline-separated string containing the full file paths of the files you want to add
'*
Public Function AddMultipleFiles(Files As String)
Dim Lines() As String
Dim i As Long
Dim ItemsCount As Long

Lines = Split(Files, vbNewLine)
ItemsCount = UBound(Lines)

AddingMultipleFiles = True

For i = 0 To UBound(Lines)
    Call AddFile(Lines(i))
    
    RaiseEvent Progress(i, ItemsCount, False)
    
    DoEvents
Next

AddingMultipleFiles = False

RaiseEvent ListChanged
RaiseEvent Progress(ItemsCount, ItemsCount, True)
End Function
'*
'* Add a single file to the playlist
'* @param String File: The full file path to the file you want to add
'* @param Long InsertAt: The position you want to add the file to the playlist
'* @return mdlAdioPlaylistItem: A model containging all the properties of the playlist item
'*
Public Function AddFile(File As String, Optional InsertAt As Long = 0) As mdlAdioPlaylistItem
Dim PlaylistItem As New mdlAdioPlaylistItem
Dim Fso As New FileSystemObject

If StrExt.IsNullOrWhiteSpace(File) Then: Exit Function
If Not Ext.FileExists(File) Then: RaiseEvent Error("File not found: " & File, 100): Exit Function
If Not CheckFileSupport(File) Then: RaiseEvent Error("File not supported: " & File, 110): Exit Function
If Not AllowDuplicateItems Then
    Dim Exists As Variant
    
    Set Exists = CheckIfItemExists(File)
    
    If Not Exists Is Nothing Then
        Set AddFile = Exists
        RaiseEvent Error("Duplicates not allowed, file already in playlist on position: " & Exists.nR, 202)
        Exit Function
    End If
End If

PlaylistItem.LocalFile = File
PlaylistItem.FileExtension = Fso.GetExtensionName(File)
PlaylistItem.FileName = Fso.GetFileName(File)
PlaylistItem.RuntimeInSeconds = modAdio.AdioReadAudioProperty(File, pDurationInSeconds)
PlaylistItem.RuntimeString = modAdio.AdioReadAudioProperty(File, pDurationString)
PlaylistItem.nR = CurList.count + 1

CurList.Add PlaylistItem

RaiseEvent NewFileAdded(PlaylistItem, InsertAt)
If Not AddingMultipleFiles Then: RaiseEvent ListChanged

Set AddFile = PlaylistItem
End Function
'*
'* Get the current playlist
'* @return Collection: A collection of all the playlist items of the current playlist
'*
Public Function GetList() As Collection
If QueryActive Then
    Set GetList = TempList
Else
    Set GetList = CurList
End If
End Function
'*
'* Clear the current playlist, and remove all the items
'*
Public Sub Clear()
Set CurList = New Collection
End Sub
'*
'* Execute a query on the current playlist
'* @param String Query: The query you want to execute
'*
Public Function ExecQuery(Query As String)
Dim PlaylistItem As New mdlAdioPlaylistItem
Dim Column As String
Dim Operator As String
Dim Value As String

Set TempList = New Collection

Column = StrExt.SplitStr(Query, Space(1), 0)
Operator = StrExt.SplitStr(Query, Space(1), 1)
 
' Check if query contains string
If StrExt.Contains(Query, Chr(39)) Then
    Value = StrExt.Between("'", "'", StrExt.SplitStr(Query, Space(1), 2))
Else
    Value = StrExt.SplitStr(Query, Space(1), 2)
End If

' Check for errors
If (LCase(Operator) <> "eq") And (LCase(Operator) <> "like") Then
    RaiseEvent Error("Query syntax error, operator not supported", 10)
    Exit Function
End If

' Execute the query
For Each PlaylistItem In CurList
    ' Exec for extension
    If (Column = "extension" And Operator = "eq") And PlaylistItem.FileExtension = Value Then TempList.Add PlaylistItem
    If (Column = "extension" And Operator = "like") And StrExt.Contains(PlaylistItem.FileExtension, Value) Then TempList.Add PlaylistItem
Next

QueryActive = True

RaiseEvent ListChanged
End Function
'*
'* Get the next, previous, etc. track from the playlist
'* @param enumAdioPlaylistDirection Direction: The direction you want to go within the playlist (next, previous, etc.)
'* @param Long TrackNr: The tracknumber when using Direction: Goto
'* @return mdlAdioPlaylistItem: Model containg the current selected playlist item
'*
Public Function GetTrack(Direction As enumAdioPlaylistDirection, Optional TrackNr As Long = 0) As mdlAdioPlaylistItem
Randomize

If RepeatMode = PLS_NO_REPEAT Then
    Select Case Direction
        Case enumAdioPlaylistDirection.PLS_FIRST: CurTrack = 1
        Case enumAdioPlaylistDirection.PLS_LAST: CurTrack = CurList.count
        Case enumAdioPlaylistDirection.PLS_NEXT: CurTrack = CurTrack + 1
        Case enumAdioPlaylistDirection.PLS_PREV: CurTrack = CurTrack - 1
        Case enumAdioPlaylistDirection.PLS_GOTO: CurTrack = TrackNr
    End Select
Else
    If RepeatMode = PLS_REPEAT And Direction = PLS_GOTO Then: CurTrack = TrackNr
    If RepeatMode = PLS_SHUFFLE And Direction = PLS_GOTO Then: CurTrack = Ext.RandomNumber(1, ListCount)
End If

' Skip if end or begin of playlist
If CurTrack = 0 Then: Exit Function
If CurTrack > CurList.count Then: Exit Function

' Check if the tracknumber exists
If Not Ext.Exists(CurList, CurTrack) Then: RaiseEvent Error("Specified tracknumber does not exists within the current playlist", 104): Exit Function

RaiseEvent TrackChanged(CurList(CurTrack))

Set GetTrack = CurList(CurTrack)
End Function
'*
'* Get the total runtime of the playlist in seconds
'* @return Long: The total runtime of the playlist
'*
Public Function GetPlaylistRuntimeInSeconds() As Long
Dim PlaylistItem As mdlAdioPlaylistItem
Dim TotalRuntime As Long

For Each PlaylistItem In GetList
    TotalRuntime = TotalRuntime + PlaylistItem.RuntimeInSeconds
Next

GetPlaylistRuntimeInSeconds = TotalRuntime
End Function
'*
'* Get the total runtime of the playlist in string format
'* @return String: The total runtime in string format
'*
Public Function GetPlaylistRuntimeString() As String
GetPlaylistRuntimeString = Ext.SecondsToTimeSerial(GetPlaylistRuntimeInSeconds(), LongTimeSerial)
End Function
'*
'* Load a playlist file
'* @param String File: The playlist file you want to load
'* @param enumAdioPlaylistType ListType: The playlist type you want to load
'*
Public Sub LoadPlaylist(File As String, ListType As enumAdioPlaylistType)
Select Case ListType
    Case enumAdioPlaylistType.PLAYLIST_APL: Call AddMultipleFiles(modPlaylist.LoadAplFile(File)) ' Audiostation
    Case enumAdioPlaylistType.PLAYLIST_M3U: Call AddMultipleFiles(modPlaylist.LoadM3uFile(File)) ' M3u file
    Case enumAdioPlaylistType.PLAYLIST_PLS: Call AddMultipleFiles(modPlaylist.LoadPlsFile(File)) ' Pls file
    Case enumAdioPlaylistType.PLAYLIST_WPL: Call AddMultipleFiles(modPlaylist.LoadWplFile(File)) ' Windows Media Player
End Select

RaiseEvent ListLoadFinished
RaiseEvent ListChanged
End Sub
'*
'* Save the current playlist to the disk
'* @param String File: The file you want to save the playlist to
'* @param enumAdioPlaylistType ListType: The type of playlist you want to use when saving
'* @return Boolean: Tells if the playlist has been saved
'*
Public Function SavePlaylist(File As String, ListType As enumAdioPlaylistType) As Boolean
Select Case ListType
    Case enumAdioPlaylistType.PLAYLIST_APL: SavePlaylist = modPlaylist.SaveAplPlaylist(File, GetList) ' Audiostation
    Case enumAdioPlaylistType.PLAYLIST_M3U: SavePlaylist = modPlaylist.SaveM3uPlaylist(File, GetList) ' M3u file
    Case enumAdioPlaylistType.PLAYLIST_PLS: SavePlaylist = modPlaylist.SavePlsPlaylist(File, GetList) ' Pls file
    Case enumAdioPlaylistType.PLAYLIST_WPL: SavePlaylist = modPlaylist.SaveWplPlaylist(File, GetList) ' Windows Media Player
End Select

RaiseEvent ListSaveFinished
End Function
'*
'* Usercontrol resize function
'*
Private Sub UserControl_Resize()
width = Image_Main.width
height = Image_Main.height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get AllowDuplicateItems() As Boolean
Attribute AllowDuplicateItems.VB_Description = "When set to true duplicate items will be allowed within the playlist, if set to false items will only be added once and grouped by there full file path"
    AllowDuplicateItems = m_AllowDuplicateItems
End Property

Public Property Let AllowDuplicateItems(ByVal New_AllowDuplicateItems As Boolean)
    m_AllowDuplicateItems = New_AllowDuplicateItems
    PropertyChanged "AllowDuplicateItems"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowDuplicateItems = m_def_AllowDuplicateItems
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AllowDuplicateItems = PropBag.ReadProperty("AllowDuplicateItems", m_def_AllowDuplicateItems)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AllowDuplicateItems", m_AllowDuplicateItems, m_def_AllowDuplicateItems)
End Sub

