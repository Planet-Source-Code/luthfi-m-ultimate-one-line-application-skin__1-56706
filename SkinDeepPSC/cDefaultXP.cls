VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDefaultXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ICON_MARGIN As Long = 6&
Private Const ICON_WIDTH As Long = 16&
Private Const ICON_SPACE As Long = ICON_MARGIN + ICON_WIDTH


Private Declare Function OleTranslateColor2 Lib "oleaut32.dll" Alias "OleTranslateColor" _
  (ByVal Clr As Long, _
  ByVal hPal As Long, _
  ByVal lpcolorref As Long) As Long

Implements ISkin

Private m_lClrSysHilite As Long
Private m_lClrSys3DFace As Long
Private m_lClrSysShadow As Long
Private m_lClrHover As Long
Private m_lClrPush As Long
Private m_lClrPushText As Long

'Private Sub Class_Terminate()
'  ISkin_DisposeSkin
'End Sub

Private Sub ISkin_DisposeSkin()

End Sub

Private Sub ISkin_DrawItem(info As DrawInfo)
  Select Case info.ItemType
    Case enButton
      DrawButton info
      
    Case enTextBox
      DrawTextBox info
  End Select
End Sub

Private Sub ISkin_InitSkin(info As InitInfo)
  With info
    .BorderWidth(enButton) = 1
    .BorderWidth(enTextBox) = 1
  End With
  m_lClrHover = AlphaBlend(vbHighlight, vbWindowBackground, 75&)
  m_lClrPush = AlphaBlend(vbHighlight, vbWindowBackground, 128&)
  OleTranslateColor vbHighlightText, 0, m_lClrPushText
  OleTranslateColor vbHighlight, 0, m_lClrSysHilite
  OleTranslateColor vbButtonShadow, 0, m_lClrSysShadow
  OleTranslateColor vbButtonFace, 0, m_lClrSys3DFace
End Sub

Private Sub DrawButton(info As DrawInfo)
  Dim tR As RECT
  Dim hOldBr As Long, hOldPen As Long
  Dim iconExists As Boolean
  
  tR = info.Rectangle
  With info
    
    iconExists = ((Not (.icon Is Nothing)) Or (.hImageList <> 0))
    
    If .IsEnable Then
    
      If .IsPushed Then
        hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysHilite))
        hOldBr = SelectObject(.hDC, CreateSolidBrush(m_lClrPush))
        SetTextColor .hDC, m_lClrPushText
        
      ElseIf .IsHover Then
        hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysHilite))
        hOldBr = SelectObject(.hDC, CreateSolidBrush(m_lClrHover))
      
      Else
        hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
        hOldBr = SelectObject(.hDC, CreateSolidBrush(m_lClrSys3DFace))
      End If
    Else
      hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
      SetTextColor .hDC, m_lClrSysShadow
    End If
    
    Rectangle .hDC, tR.left, tR.top, tR.right, tR.bottom

    If .IsFocus Then
      InflateRect tR, -3, -3
      DrawFocusRect .hDC, tR
      InflateRect tR, 3, 3
    End If
    
    'icon
    If iconExists Then
      'draw icon and adjust rect for caption
      
      Select Case .IconPosition
        Case eipLeftMiddle
          If .hImageList <> 0 Then
          
          Else
            
            DrawItemIcon .hDC, tR.left + ICON_MARGIN, (tR.bottom - tR.top - ICON_WIDTH) \ 2&, _
              .icon.Handle, .IsEnable
          End If
          tR.left = tR.left + ICON_SPACE
      
        Case eipRightMiddle
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hDC, tR.right - ICON_SPACE, (tR.bottom - tR.top - ICON_WIDTH) \ 2&, _
              .icon.Handle, .IsEnable
          End If
          tR.right = tR.right - ICON_SPACE
            
        Case eipCenterTop
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hDC, (tR.right - tR.left - ICON_WIDTH) \ 2&, 4&, _
              .icon.Handle, .IsEnable
          End If
          tR.top = tR.top + ICON_SPACE
      
        Case eipCenterBottom
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hDC, (tR.right - tR.left - ICON_WIDTH) \ 2&, tR.bottom - ICON_SPACE, _
              .icon.Handle, .IsEnable
          End If
          tR.bottom = tR.bottom - ICON_SPACE
      
      End Select
      
    End If
    
    'caption
    SetBkMode .hDC, TRANSPARENT
    DrawText .hDC, .caption, -1, tR, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    
    If hOldBr <> 0 Then DeleteObject (SelectObject(.hDC, hOldBr))
    If hOldPen <> 0 Then DeleteObject (SelectObject(.hDC, hOldPen))
    
  End With
  
End Sub

Private Sub DrawTextBox(info As DrawInfo)
  Dim tR As RECT, hOldPen As Long
  
  With info
    tR = .Rectangle
    
    If .IsHover Then
      hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysHilite))
    Else
      hOldPen = SelectObject(.hDC, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
    End If
    
    MoveToEx .hDC, tR.left, tR.top, ByVal 0
    LineTo .hDC, tR.right, tR.top
    LineTo .hDC, tR.right, tR.bottom
    LineTo .hDC, tR.left, tR.bottom
    LineTo .hDC, tR.left, tR.top
    
    If hOldPen <> 0 Then DeleteObject (SelectObject(.hDC, hOldPen))
  End With
  
End Sub

Private Property Get ISkin_Name() As String
  ISkin_Name = "Default XP 2002 Skin"
End Property

Private Function AlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  Dim clrFore         As RGBQUAD
  Dim clrBack         As RGBQUAD
  Dim lDif As Long
  On Error Resume Next
  OleTranslateColor2 clrFirst, 0, VarPtr(clrFore)
  OleTranslateColor2 clrSecond, 0, VarPtr(clrBack)
  If lAlpha < 0& Then lAlpha = 0&
  If lAlpha > 255& Then lAlpha = 255&
  With clrFore
    lDif = CLng(.rgbRed) - CLng(clrBack.rgbRed)
    .rgbRed = (lAlpha * lDif) \ 255 + clrBack.rgbRed
    lDif = CLng(.rgbGreen) - CLng(clrBack.rgbGreen)
    .rgbGreen = (lAlpha * lDif) \ 255 + clrBack.rgbGreen
    lDif = CLng(.rgbBlue) - CLng(clrBack.rgbBlue)
    .rgbBlue = (lAlpha * lDif) \ 255 + clrBack.rgbBlue
  End With
  CopyMemory ByVal VarPtr(AlphaBlend), ByVal VarPtr(clrFore), 4
End Function

Private Sub DrawItemIcon(ByVal hDC As Long, ByVal left As Long, ByVal top As Long, _
  ByVal hIcon As Long, Optional IsEnable As Boolean = True)
  
  Dim hBr As Long
  
  If IsEnable Then
    DrawIconEx hDC, left, top, _
      hIcon, ICON_WIDTH, ICON_WIDTH, 0, 0, DI_NORMAL
  Else
    hBr = CreateSolidBrush(m_lClrSysShadow)
    
    DrawState hDC, ByVal hBr, ByVal 0, ByVal hIcon, _
      ByVal 0, left, top, ICON_WIDTH, ICON_WIDTH, DST_ICON Or DSS_MONO
    
    DeleteObject hBr
  End If
End Sub
