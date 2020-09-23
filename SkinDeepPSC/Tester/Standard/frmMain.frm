VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   2
      Left            =   450
      TabIndex        =   1
      Text            =   "No Border"
      Top             =   1230
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   1
      Left            =   450
      TabIndex        =   5
      Text            =   "Normal Textbox"
      Top             =   780
      Width           =   2445
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   1830
      Width           =   1245
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable"
      Enabled         =   0   'False
      Height          =   405
      Left            =   3090
      TabIndex        =   3
      Top             =   1830
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   450
      TabIndex        =   2
      Text            =   "Disabled"
      Top             =   330
      Width           =   2445
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   405
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
'optional code, can be omitted entirely but its better to be safe
'this is the last form to unload so remove our skin
  Set activeskin = Nothing
End Sub

'Example of a simple project, see mMain for initialization routine
Private Sub mnuExit_Click()
  Unload Me
End Sub
