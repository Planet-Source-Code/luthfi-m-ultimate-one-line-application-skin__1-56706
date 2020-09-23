VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Icon requester
'event for requesting icon from the client
'History:
'Sept 27, 2004 : Created

Option Explicit

#If False Then 'IDE hack
  Dim eipLeftMiddle, eipRightMiddle, eipCenterTop, eipCenterBottom
#End If

Public Enum enIconPosition
  eipLeftMiddle
  eipRightMiddle
  eipCenterTop
  eipCenterBottom
End Enum

Public Event RequestIconList( _
  ByRef hImageList As Long, ByRef index As Long, ByRef position As enIconPosition, _
  ByVal hWnd As Long, ByVal caption As String)

Public Event RequestIcon( _
  ByRef icon As StdPicture, ByRef position As enIconPosition, _
  ByVal hWnd As Long, ByVal caption As String)

Friend Sub RaiseIconList( _
  ByRef hImageList As Long, ByRef index As Long, ByRef position As enIconPosition, _
  ByVal hWnd As Long, ByVal caption As String)
  RaiseEvent RequestIconList(hImageList, index, position, hWnd, caption)
End Sub

Friend Sub RaiseIcon(ByRef icon As StdPicture, ByRef position As enIconPosition, _
  ByVal hWnd As Long, ByVal caption As String)
  
  RaiseEvent RequestIcon(icon, position, hWnd, caption)
End Sub

