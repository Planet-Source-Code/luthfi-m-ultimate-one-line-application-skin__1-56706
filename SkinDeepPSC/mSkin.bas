Attribute VB_Name = "mSkin"
Option Explicit

Public Enum enMsgWhen
  MSG_AFTER = 1
  MSG_BEFORE = 2
  MSG_BEFORE_AND_AFTER = 3
  ALL_MESSAGES = &HFFFFFFFF
End Enum

Private m_IconRequest As IconRequester
Dim iconInit As Boolean

Property Get IconRequest() As IconRequester
  If iconInit = False Then
    Set m_IconRequest = New IconRequester
    iconInit = True
  End If
  Set IconRequest = m_IconRequest
End Property

Function HookProc(ByVal lHookPtr As Long) As cSkinProc
  Dim oTemp As cSkinProc
   ' Turn the pointer into an illegal, uncounted interface
  If lHookPtr = 0 Then Exit Function
  CopyMemory oTemp, lHookPtr, 4
  ' Do NOT hit the End button here! You will crash!
  ' Assign to legal reference
  Set HookProc = oTemp
  ' Still do NOT hit the End button here! You will still crash!
  ' Destroy the illegal reference
  CopyMemory oTemp, 0&, 4
End Function

Sub GetItemIcon( _
  ByRef hIml As Long, _
  ByRef lIdx As Long, _
  ByRef icon As StdPicture, _
  ByRef position As enIconPosition, _
  ByVal hWnd As Long, ByVal caption As String)
  
  If iconInit = False Then Exit Sub
  
  With m_IconRequest
    'raise the imagelist version first
    .RaiseIconList hIml, lIdx, position, hWnd, caption
    
    'if no one interested raise the stdpic version
    If hIml = 0 Then
      .RaiseIcon icon, position, hWnd, caption
    End If
  End With
End Sub
