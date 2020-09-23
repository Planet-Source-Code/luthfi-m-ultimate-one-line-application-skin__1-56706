VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSkinProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Hook Loop Processor
'Manage All Messages dispatched by any windows in our application

'History:**********************************
'Sep 20, 2004 : Created

'******************************************

Option Explicit

Private m_HookCWP As cHookGate 'Hook Gate for WH_CALLWNDPROC
Private m_HookGM As cHookGate 'Hook Gate for WH_getmessage
Private m_Subclass As cSubclass

Private m_DrawInfo As DrawInfo
Private m_InitInfo As InitInfo
Private m_Skin As ISkin
Private m_hWndHover As Long
Private m_lPtr As Long

Dim lastClass As String

Implements ISubclass

Private Sub Class_Initialize()
  Set m_Subclass = New cSubclass
  Set m_DrawInfo = New DrawInfo
  Set m_InitInfo = New InitInfo
  m_lPtr = ObjPtr(Me)
  'Create callback gate for WH_CWP
  Set m_HookCWP = New cHookGate
  m_HookCWP.StartHook ehtCWP, m_lPtr
  'Create callback gate for WH_GetMessage
  Set m_HookGM = New cHookGate
  m_HookGM.StartHook ehtGetMessage, m_lPtr
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  If Not m_Skin Is Nothing Then
    m_Skin.DisposeSkin
    Set m_Skin = Nothing
  End If
  
  Set m_Subclass = Nothing
  m_HookGM.StopHook
  m_HookCWP.StopHook
  
  Set m_HookGM = Nothing
  Set m_HookCWP = Nothing
  
  Set m_InitInfo = Nothing
  Set m_DrawInfo = Nothing
End Sub

Property Set ActiveSkin(ByVal value As ISkin)
  Set m_Skin = value
End Property

Property Get ActiveSkin() As ISkin
  Set ActiveSkin = m_Skin
End Property

Property Get InitInfo() As InitInfo
  Set InitInfo = m_InitInfo
End Property

Public Sub HookProc(ByVal bBefore As Boolean, bHandled As Boolean, _
  lReturn As Long, nCode As Long, wParam As Long, lParam As Long, _
  ByVal HookType As enHookType)
  
  Dim tCWP As CWPSTRUCT
  Dim tMSG As MSG
  Dim hwndBefore As Long
  
  'early exit
  If m_Skin Is Nothing Then Exit Sub
  If nCode <> HC_ACTION Then Exit Sub
  If bBefore = False Then Exit Sub
  
  lastClass = ""
  Select Case HookType
    Case ehtCWP 'Process WH_CWP
      CopyMemory tCWP, ByVal lParam, Len(tCWP)
      
      Select Case tCWP.message
        Case WM_CREATE
          OnCreate tCWP.hwnd, tCWP.lParam
        
        Case WM_DRAWITEM
          OnDrawItem tCWP.lParam
        
        Case WM_STYLECHANGED
          OnStyleChanged tCWP.hwnd, tCWP.wParam, tCWP.lParam
        
        Case WM_NCPAINT
          lastClass = GetWinText(tCWP.hwnd, True)
          Select Case lastClass
            Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
              If IsBorderLess(tCWP.hwnd) Then Exit Sub 'in case there's wm_erasebkgnd generated
              PlaceSubClass tCWP.hwnd, WM_NCPAINT
              
          End Select
      End Select
  
    Case ehtGetMessage 'Process WH_GETMESSAGE
      CopyMemory tMSG, ByVal lParam, Len(tMSG)
      lastClass = ""
      If tMSG.message = WM_MOUSEMOVE Then
        If m_hWndHover <> tMSG.hwnd Then
          hwndBefore = m_hWndHover
          m_hWndHover = tMSG.hwnd
          
          If hwndBefore <> 0 Then
            
            'redraw previous hovered window
            RequestRedraw hwndBefore
          End If
          
          'redraw new hovered window
          RequestRedraw tMSG.hwnd
        End If
      End If
  End Select
End Sub

Private Sub OnCreate(hwnd As Long, lParam As Long)
  Dim className As String
  Dim tCS As CREATESTRUCT
  
  className = GetWinText(hwnd, True)
  
  Select Case className
    
    Case "ThunderCommandButton", "ThunderRT6CommandButton" 'VB Command Button
      CopyMemory tCS, ByVal lParam, ByVal Len(tCS)
      'clear up all style bits except for BS_OWNERDRAW
       
      Call SetWindowLong(hwnd, GWL_STYLE, (tCS.style And &HFFFF0000) Or BS_OWNERDRAW)
    
    Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
      CopyMemory tCS, ByVal lParam, ByVal Len(tCS)
      
      'make it flat
      If (tCS.dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then
        SetWindowLong hwnd, GWL_EXSTYLE, tCS.dwExStyle And (Not WS_EX_CLIENTEDGE) Or WS_EX_STATICEDGE
        'notify changes
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER
      End If
  End Select
End Sub

Private Sub OnDrawItem(lParam As Long)
  Dim tDI As DRAWITEMSTRUCT
  Dim hIml As Long, lIdx As Long, _
    icon As StdPicture, _
    position As enIconPosition

  
  CopyMemory tDI, ByVal lParam, Len(tDI)
  
  With m_DrawInfo
    Select Case tDI.CtlType
      
      Case ODT_BUTTON
        .ItemType = enButton
        .hdc = tDI.hdc
        .caption = GetWinText(tDI.hwndItem)
        If tDI.hwndItem = m_hWndHover Then tDI.itemState = tDI.itemState Or ODS_HOTLIGHT
        
        .SetState tDI.itemState
        .SetRect tDI.rcItem.Left, tDI.rcItem.Top, tDI.rcItem.Right, tDI.rcItem.Bottom
        
        GetItemIcon hIml, lIdx, icon, position, tDI.hwndItem, .caption
        
        .IconPosition = position
        
        If hIml <> 0 Then
          .hImageList = hIml
          .IconIndex = lIdx
        Else
          Set .icon = icon
        End If
        
        m_Skin.DrawItem m_DrawInfo
    End Select
  End With
End Sub

Private Sub OnStyleChanged(hwnd As Long, wParam As Long, lParam As Long)
  Dim className As String
  Dim tSS As POINTAPI 'Stylestruct
  Static bInStyleChange As Boolean
  
  className = GetWinText(hwnd, True)
  
  Select Case className
    Case "ThunderCommandButton", "ThunderRT6CommandButton", "Edit" 'VB Command Button
    If bInStyleChange = 0 Then
      If (wParam And GWL_STYLE) = GWL_STYLE Then
        bInStyleChange = True
        
        CopyMemory tSS, ByVal lParam, Len(tSS)
        'Debug.Print GetWinText(hWnd, False), tSS.cx, tSS.cy
        tSS.y = (tSS.x And &HFFFF0000) Or BS_OWNERDRAW
        SetWindowLong hwnd, GWL_STYLE, tSS.y
        
        bInStyleChange = False
      End If
    End If
  End Select
End Sub

Private Sub OnNCPaint(hwnd As Long)
  Dim lStyle As Long, tR As RECT
  Select Case lastClass
    Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"

      With m_DrawInfo
        .ItemType = enTextBox
        .hdc = GetWindowDC(hwnd)
        
        If hwnd = m_hWndHover Then lStyle = ODS_HOTLIGHT
        If IsWindowEnabled(hwnd) = 0 Then lStyle = lStyle Or ODS_DISABLED
        If hwnd = GetFocus() Then lStyle = lStyle Or ODS_FOCUS
        GetWindowRect hwnd, tR
        
        .SetState lStyle
        .SetRect 0, 0, tR.Right - tR.Left - 1&, tR.Bottom - tR.Top - 1&
        
        m_Skin.DrawItem m_DrawInfo
        ReleaseDC hwnd, .hdc
      End With
      
  End Select
End Sub

Private Sub RequestRedraw(hwnd As Long)
  
  Select Case GetWinText(hwnd, True)
    Case "ThunderCommandButton", "ThunderRT6CommandButton"
      RedrawWindow hwnd, ByVal 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE
    
    Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
      If IsBorderLess(hwnd) = False Then
        RedrawWindow hwnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_FRAME
      End If
  End Select
End Sub

Private Function GetWinText(hwnd As Long, Optional className As Boolean) As String
  'some static vars to speed up things, this func will be called many times
  Static sBuffer As String * 128& 'is it safe to use 128 bytes? should be enough..
  Static textLength As Long
      
  If className Then
    textLength = GetClassName(hwnd, sBuffer, 129&)
  Else
    textLength = GetWindowText(hwnd, sBuffer, 129&)
  End If
  
  If textLength > 0 Then
    GetWinText = Left$(sBuffer, textLength)
  End If
End Function

Private Sub PlaceSubClass(hwnd As Long, MSG As Long, Optional When As enMsgWhen = MSG_AFTER)
  With m_Subclass
    .UnSubclass
    .AddMsg MSG, When
    .Subclass hwnd, Me
  End With
End Sub

Private Sub ISubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
  Select Case uMsg
    Case WM_NCPAINT
      OnNCPaint hwnd
  End Select
  m_Subclass.UnSubclass
End Sub

Private Function IsBorderLess(hwnd As Long) As Boolean
  IsBorderLess = ((GetWindowLong(hwnd, GWL_EXSTYLE) And WS_EX_STATICEDGE) <> WS_EX_STATICEDGE)
End Function

