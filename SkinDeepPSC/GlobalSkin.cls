VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalSkin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Main Singleton Class **********************************
'Holds the runtime properties, and main entry point
'*******************************************************

'History:**********************************
'Sep 20, 2004 : Created

'******************************************

Option Explicit
Private m_SkinProc As cSkinProc

Public Sub Integrate()
  Set m_SkinProc = New cSkinProc
  If m_SkinProc.ActiveSkin Is Nothing Then
    Set ActiveSkin = New cDefaultXP
  End If
End Sub

Property Set ActiveSkin(ByVal value As ISkin)
  If Not m_SkinProc.ActiveSkin Is Nothing Then m_SkinProc.ActiveSkin.DisposeSkin
  Set m_SkinProc.ActiveSkin = value
  If value Is Nothing Then Exit Property
  value.InitSkin m_SkinProc.InitInfo
End Property

Property Get IconRequester() As IconRequester
  Set IconRequester = IconRequest
End Property

Sub RedrawAll()
  'should have refresh windows belongs to this application only
  'for the moment it's refresh all windows within the desktop
  RedrawWindow ByVal 0, ByVal 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ALLCHILDREN
End Sub

Private Sub Class_Terminate()
  Set m_SkinProc = Nothing
End Sub


