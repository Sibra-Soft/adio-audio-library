Attribute VB_Name = "modMain"
Option Explicit

Public SoundFont As Long

Public TagLib As New clsAdioTagLibrary
Public StrExt As New clsStringExtensions
Public Ext As New clsSibraSoft
Public Function CheckFileSupport(File As String) As Boolean
Dim Fso As New FileSystemObject

' Check if the file is supported by Adio
Select Case Fso.GetExtensionName(File)
    Case "mp1", "mp2", "mp3", "wav", "ogg", "aiff", "aac", "wma", "flac": CheckFileSupport = True: Exit Function
    Case "mid", "midi", "kar", "rmi": CheckFileSupport = True: Exit Function
    
    Case Else: CheckFileSupport = False
End Select
End Function
