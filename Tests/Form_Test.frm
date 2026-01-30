VERSION 5.00
Object = "{2FA705D9-0A01-4278-9B75-4ACD7CC6E510}#1.0#0"; "SimplyVBUnit.Component.ocx"
Object = "*\A..\Source\AdioAudioLibrary.vbp"
Begin VB.Form Form_Test 
   Caption         =   "Tests"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin AdioLibrary.AdioPlaylist AdioPlaylist 
      Left            =   2835
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin AdioLibrary.AdioTagging AdioTagging 
      Left            =   2295
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin SimplyVBComp.UIRunner UIRunner1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
Call Me.UIRunner1.Init(App)
End Sub
Private Sub Form_Load()
AddTest New TestAdioTagging
AddTest New TestAdioPlaylist
End Sub
