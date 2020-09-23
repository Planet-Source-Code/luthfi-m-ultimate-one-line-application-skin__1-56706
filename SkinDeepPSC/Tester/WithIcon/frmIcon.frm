VERSION 5.00
Begin VB.Form frmIcon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDisabled 
      Caption         =   "Disabled"
      Enabled         =   0   'False
      Height          =   585
      Left            =   3660
      TabIndex        =   6
      Top             =   2460
      Width           =   1215
   End
   Begin VB.OptionButton optIconPos 
      Caption         =   "Center Bottom"
      Height          =   375
      Index           =   3
      Left            =   930
      TabIndex        =   5
      Top             =   1740
      Width           =   1845
   End
   Begin VB.OptionButton optIconPos 
      Caption         =   "Center Top"
      Height          =   375
      Index           =   2
      Left            =   930
      TabIndex        =   4
      Top             =   1340
      Width           =   1245
   End
   Begin VB.OptionButton optIconPos 
      Caption         =   "Right Middle"
      Height          =   375
      Index           =   1
      Left            =   930
      TabIndex        =   3
      Top             =   940
      Width           =   1245
   End
   Begin VB.OptionButton optIconPos 
      Caption         =   "Left Middle"
      Height          =   375
      Index           =   0
      Left            =   930
      TabIndex        =   2
      Top             =   540
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   585
      Left            =   2190
      TabIndex        =   1
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Height          =   585
      Left            =   780
      TabIndex        =   0
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Image imgCancel 
      Height          =   240
      Left            =   600
      Picture         =   "frmIcon.frx":0000
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSave 
      Height          =   240
      Left            =   150
      Picture         =   "frmIcon.frx":058A
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Example of handling icon request
Option Explicit

Dim WithEvents iconHandler As Skindeep.IconRequester
Attribute iconHandler.VB_VarHelpID = -1
Dim iconPos As enIconPosition

Private Sub Form_Load()
  Set iconHandler = Skindeep.IconRequester
End Sub

'only handle one of the event version
Private Sub iconHandler_RequestIcon(icon As stdole.StdPicture, position As Skindeep.enIconPosition, ByVal hWnd As Long, ByVal caption As String)
  
  position = iconPos
  
  Select Case hWnd
    Case cmdSave.hWnd, cmdDisabled.hWnd
      Set icon = imgSave.Picture
      
    Case cmdCancel.hWnd
      Set icon = imgCancel.Picture
  End Select
End Sub

Private Sub optIconPos_Click(Index As Integer)
  iconPos = Index
  Skindeep.RedrawAll
End Sub
