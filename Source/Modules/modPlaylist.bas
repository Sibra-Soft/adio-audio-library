Attribute VB_Name = "modPlaylist"
Public Function SavePlsPlaylist(File As String, list As Collection) As Boolean
Dim PlsItem As mdlAdioPlaylistItem

For Each PlsItem In list
    Call Ext.INIWrite("playlist", "File" & i, PlsItem.LocalFile, File)
Next

Call Extensions.INIWrite("playlist", "NumberOfEntries", lstFormList.ListItems.count, File)
Call Extensions.INIWrite("playlist", "Version", 2, File)

' Check if the playlist has been saved
If Ext.FileExists(File) Then
    SavePlsPlaylist = True
Else
    SavePlsPlaylist = False
End If
End Function
Public Function SaveAplPlaylist(File As String, list As Collection) As Boolean
Dim PlsItem As mdlAdioPlaylistItem

Open File For Output As #FN
    For Each PlsItem In list
        Print #FN, PlsItem.LocalFile
    Next
Close #FN

' Check if the playlist has been saved
If Ext.FileExists(File) Then
    SaveAplPlaylist = True
Else
    SaveAplPlaylist = False
End If
End Function
Public Function SaveM3uPlaylist(File As String, list As Collection) As Boolean
Dim PlsItem As mdlAdioPlaylistItem

Open File For Output As #FN
    Print #FN, "#EXTM3U"
    
    For Each PlsItem In list
      Print #FN, "#EXTINF:0, " & Ext.GetFileNameFromFilePath(PlsItem.LocalFile, False)
      Print #FN, PlsItem.LocalFile
      Print #FN, ""
    Next
Close #FN

' Check if the playlist has been saved
If Ext.FileExists(File) Then
    SaveM3uPlaylist = True
Else
    SaveM3uPlaylist = False
End If
End Function
Public Function SaveWplPlaylist(File As String, list As Collection) As Boolean
Dim PlaylistName As String
Dim PlsItem As mdlAdioPlaylistItem

PlaylistName = Ext.GetFileNameFromFilePath(File, False)

Open File For Output As #1
    Print #1, "<?wpl version="; 1#; "?>"
    Print #1, "<smil>"
    Print #1, "    <head>"
    Print #1, "        <title>" & PlaylistName & "</title>"
    Print #1, "    </head>"
    Print #1, "    <body>"
    Print #1, "        <seq>"
    
    ' Get all the items from the selected playlist
    For Each plstItem In list
      Print #1, "<media src=""" & PlsItem.LocalFile & """/>"
    Next
    
    Print #1, "       </seq>"
    Print #1, "    </body>"
    Print #1, "</smil>"
Close #1

' Check if the playlist has been saved
If Ext.FileExists(File) Then
    SaveWplPlaylist = True
Else
    SaveWplPlaylist = False
End If
End Function
Public Function LoadAplFile(File As String) As String
Dim StringListOfFiles As String
StringListOfFiles = Ext.FileGetContents(File)

LoadAplFile = StringListOfFiles
End Function
Public Function LoadWplFile(File As String) As String
Dim Lines
Dim FileContent As String
Dim i As Integer
Dim Media As String
Dim StringListOfFiles As String

FileContent = Extensions.FileGetContents(FileName)
Lines = Split(FileContent, vbNewLine)

For i = 0 To UBound(Lines)
    If InStr(1, Lines(i), "<media") Then
        Media = StrExt.Between("<media", "/>", Trim(Lines(i)))
        Media = Replace(Media, Chr(34), vbNullString)
        Media = Replace(Media, "media src=", vbNullString)
        
        StringListOfFiles = StringListOfFiles & Media & vbNewLine
    End If
Next

LoadWplFile = StringListOfFiles
End Function
Public Function LoadM3uFile(File As String) As String
Dim TextLine As String, FN As Integer
Dim StringListOfFiles As String

FN = FreeFile

'Add the files to the array
Open strPlaylistFile For Input As #FN
    Do While Not EOF(FN)
        Line Input #FN, TextLine
        If TextLine <> LineToRem Then
            If Left(TextLine, 7) = "#EXTM3U" Then
                Debug.Print "Playlist Type: M3U"
            Else
                If Left(TextLine, 8) = "#EXTINF:" Then
                    Debug.Print "Info Data: " & TextLine
                Else
                    StringListOfFiles = StringListOfFiles & TextLine & vbNewLine
                End If
            End If
        End If
    Loop
Close #FN

LoadM3uFile = StringListOfFiles
End Function
Public Function LoadPlsFile(File As String) As String
Dim i As Integer
Dim strNumberofEntries As Integer
Dim StringListOfFiles As String

strNumberofEntries = Extensions.INIRead("playlist", "NumberOfEntries", strPlaylistFile)

For i = 1 To strNumberofEntries
    StringListOfFiles = StringListOfFiles & Extensions.INIRead("playlist", "File" & i, strPlaylistFile) & vbNewLine
Next

LoadPlsFile = StringListOfFiles
End Function
