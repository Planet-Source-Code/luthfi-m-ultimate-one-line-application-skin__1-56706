VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DrawInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'DrawInfo class, holds all information needed to draw an Item

'History:
'- Sept 24, 2004 : Created

Option Explicit

#If False Then 'ide hack
  Dim enButton, enTextBox
#End If

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Enumeration_____________________________
Public Enum enDrawItemType
  enButton = 1
  enTextBox
End Enum
'Enumeration_____________________________

'Begin Private members___________________________
Private m_hDC As Long
Private m_tR As RECT
Private m_lState As Long
Private m_hfont As Long
Private m_sCaption As String
Private m_eItemType As enDrawItemType
Private m_hIml As Long
Private m_idx As Long
Private m_iconPic As StdPicture
Private m_eIconPos As enIconPosition
'End Private members_____________________________

Property Get ItemType() As enDrawItemType
  ItemType = m_eItemType
End Property

Friend Property Let ItemType(value As enDrawItemType)
  m_eItemType = value
End Property

Property Get hdc() As Long
  hdc = m_hDC
End Property

Friend Property Let hdc(value As Long)
  m_hDC = value
End Property

Property Get Rectangle() As RECT
  Rectangle = m_tR
End Property

Friend Sub SetRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
  With m_tR
    .Left = Left
    .Top = Top
    .Right = Right
    .Bottom = Bottom
  End With
End Sub

Friend Sub SetState(value As Long)
  m_lState = value
End Sub

Property Get IsHover() As Boolean
  IsHover = ((m_lState And ODS_HOTLIGHT) = ODS_HOTLIGHT)
End Property

Property Get IsEnable() As Boolean
  IsEnable = Not ((m_lState And ODS_DISABLED) = ODS_DISABLED)
End Property

Property Get IsFocus() As Boolean
  IsFocus = ((m_lState And ODS_FOCUS) = ODS_FOCUS)
End Property

Property Get IsPushed() As Boolean
  IsPushed = ((m_lState And BST_CHECKED) = BST_CHECKED)
End Property

Property Get IsDefault() As Boolean
  'IsDefault = ((m_lState And BST_CHECKED) = BST_CHECKED)
End Property

Property Get hFont() As Long
  hFont = m_hfont
End Property

Property Get caption() As String
  caption = m_sCaption
End Property

Friend Property Let caption(ByVal value As String)
  m_sCaption = value
End Property

Property Get hImageList() As Long
  hImageList = m_hIml
End Property

Friend Property Let hImageList(ByVal value As Long)
  m_hIml = value
End Property

Property Get IconIndex() As Long
  IconIndex = m_idx
End Property

Friend Property Let IconIndex(ByVal value As Long)
  m_idx = value
End Property

Property Get icon() As StdPicture
  Set icon = m_iconPic
End Property

Friend Property Set icon(ByVal value As StdPicture)
  Set m_iconPic = value
End Property

Property Get IconPosition() As enIconPosition
  IconPosition = m_eIconPos
End Property

Friend Property Let IconPosition(ByVal value As enIconPosition)
  m_eIconPos = value
End Property


Private Sub Class_Terminate()
  Set m_iconPic = Nothing
End Sub
