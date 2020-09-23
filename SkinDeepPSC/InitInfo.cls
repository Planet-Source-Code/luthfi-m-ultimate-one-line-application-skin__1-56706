VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InitInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Skin Initialization Info Class
'Have no use for it now, but provided just in case for future development

Option Explicit

Private m_BorderWidth(enButton To enTextBox) As Long

Property Get BorderWidth(ByVal ItemType As enDrawItemType) As Long
  BorderWidth = m_BorderWidth(ItemType)
End Property

Property Let BorderWidth(ByVal ItemType As enDrawItemType, ByVal value As Long)
  m_BorderWidth(ItemType) = value
End Property

