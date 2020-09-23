VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Custom Skin Interface Definition
'Should move it to TLB, but for now it's resides with the main DLL

Option Explicit

'Creator & Destructor **********************************
Sub InitSkin(info As InitInfo)
'Place all initialization routines here
'such as creating module level brushes, pens
End Sub

Sub DisposeSkin()
'Destroy all module level objects
End Sub
'End Creator & Destructor ******************************

'Drawing Interfaces ************************************
Sub DrawItem(info As DrawInfo)
'the actual drawing codes goes here
End Sub
'End Drawing Interfaces ************************************

'Properties
Property Get Name() As String

End Property
