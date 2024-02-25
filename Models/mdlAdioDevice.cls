VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mdlAdioDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'///////////////////////////////////////////////////////////////
'// FileName        : AdioDevice.cls
'// FileType        : Microsoft Visual Basic 6 - Class Module
'// Type            : Model
'// Author          : Alex van den Berg
'// Created         : 19-08-2023
'// Last Modified   : 14-10-2023
'// Copyright       : Sibra-Soft
'// Description     : Adio device model
'////////////////////////////////////////////////////////////////

Option Explicit

Public dName As String
Public dId As Long
Public dDriver As String
Public dOutput As Boolean
Public dInput As Boolean
Public dType As String
Public dIsLoopback As Boolean
Public dIsEnabled As Boolean
Public dIsDefault As Boolean
