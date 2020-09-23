VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMySkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'bitmap source from buttons.jpg
'each cell is 22x22 dimension
'order hover,pushed, normal , disable

Option Explicit

Private Const PS_SOLID As Long = 0
Private Const TRANSPARENT As Long = 1
Private Const DT_CENTER As Long = &H1
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const DI_NORMAL As Long = &H3
Private Const DST_ICON As Long = &H3
Private Const DSS_MONO As Long = &H80

Private Const ICON_MARGIN As Long = 6&
Private Const ICON_WIDTH As Long = 16&
Private Const ICON_SPACE As Long = ICON_MARGIN + ICON_WIDTH

Private Type RECT
  left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private m_lClrSysHilite As Long
'Private m_lClrSys3DFace As Long
Private m_lClrSysShadow As Long
Private m_hBitmapDC As Long
Dim hOldBmp As Long

Implements ISkin

Private Sub Class_Terminate()
  CleanUp
End Sub

Private Sub ISkin_DisposeSkin()
  CleanUp
End Sub

Private Sub ISkin_DrawItem(info As SkinDeep.DrawInfo)
  Select Case info.ItemType
    Case enButton
      DrawButton info
      
    Case enTextBox
      DrawTextBox info
  End Select
End Sub

Private Sub ISkin_InitSkin(info As SkinDeep.InitInfo)
  OleTranslateColor vbHighlight, 0, m_lClrSysHilite
  OleTranslateColor vbButtonShadow, 0, m_lClrSysShadow
  LoadBitmaps
End Sub

Private Property Get ISkin_Name() As String
  ISkin_Name = "My Own Skin"
End Property

Private Sub DrawButton(info As DrawInfo)
  Dim tR As RECT
  Dim iconExists As Boolean
  Dim offset As Long
  
  tR = info.Rectangle
  With info
    
    iconExists = ((Not (.Icon Is Nothing)) Or (.hImageList <> 0))
    
    If .IsEnable Then
    
      If .IsPushed Then
        offset = 23& 'cell 2
      ElseIf .IsHover Then
        'offset = 0 'cell 1
      Else
        offset = 45& 'cell 3
      End If
    
    Else
      SetTextColor .hdc, m_lClrSysShadow
      offset = 67& 'cell 4
    End If
    
    DrawButtonBitmap .hdc, tR, offset
    
    If .IsFocus Then
      InflateRect tR, -3, -3
      DrawFocusRect .hdc, tR
      InflateRect tR, 3, 3
    End If
    
    'icon
    If iconExists Then
      'draw icon and adjust rect for caption
      
      Select Case .IconPosition
        Case eipLeftMiddle
          If .hImageList <> 0 Then
          
          Else
            
            DrawItemIcon .hdc, tR.left + ICON_MARGIN, (tR.Bottom - tR.top - ICON_WIDTH) \ 2&, _
              .Icon.Handle, .IsEnable
          End If
          tR.left = tR.left + ICON_SPACE
      
        Case eipRightMiddle
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hdc, tR.Right - ICON_SPACE, (tR.Bottom - tR.top - ICON_WIDTH) \ 2&, _
              .Icon.Handle, .IsEnable
          End If
          tR.Right = tR.Right - ICON_SPACE
            
        Case eipCenterTop
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hdc, (tR.Right - tR.left - ICON_WIDTH) \ 2&, 4&, _
              .Icon.Handle, .IsEnable
          End If
          tR.top = tR.top + ICON_SPACE
      
        Case eipCenterBottom
          If .hImageList <> 0 Then
          
          Else
            DrawItemIcon .hdc, (tR.Right - tR.left - ICON_WIDTH) \ 2&, tR.Bottom - ICON_SPACE, _
              .Icon.Handle, .IsEnable
          End If
          tR.Bottom = tR.Bottom - ICON_SPACE
      
      End Select
      
    End If
    
    'caption
    SetBkMode .hdc, TRANSPARENT
    DrawText .hdc, .Caption, -1, tR, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    
    
  End With
End Sub

Private Sub DrawTextBox(info As DrawInfo)
  Dim tR As RECT, hOldPen As Long
  
  With info
    tR = .Rectangle
    
    If .IsHover Then
      hOldPen = SelectObject(.hdc, CreatePen(PS_SOLID, 1, m_lClrSysHilite))
    Else
      hOldPen = SelectObject(.hdc, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
    End If
    
    MoveToEx .hdc, tR.left, tR.top, ByVal 0
    LineTo .hdc, tR.Right, tR.top
    LineTo .hdc, tR.Right, tR.Bottom
    LineTo .hdc, tR.left, tR.Bottom
    LineTo .hdc, tR.left, tR.top
    
    If hOldPen <> 0 Then DeleteObject (SelectObject(.hdc, hOldPen))
  End With
  
End Sub

Private Sub LoadBitmaps()
  Dim pic As StdPicture
  Dim screenDC As Long
  Dim hdc As Long
  
  Set pic = LoadPicture(App.Path & "\buttons.jpg")
  screenDC = GetDC(0)
  m_hBitmapDC = CreateCompatibleDC(screenDC)
'  hDC = CreateCompatibleDC(screenDC)
  
  hOldBmp = SelectObject(m_hBitmapDC, pic.Handle)
  
  ReleaseDC 0, screenDC
  
  Set pic = Nothing
End Sub

Private Sub CleanUp()
  
  If hOldBmp <> 0 Then
    SelectObject m_hBitmapDC, hOldBmp
    hOldBmp = 0
  End If
  If m_hBitmapDC <> 0 Then
    DeleteDC m_hBitmapDC
    m_hBitmapDC = 0
  End If
  
End Sub

Private Sub DrawButtonBitmap(ByVal hdc As Long, tR As RECT, offset As Long)
'Left side
  StretchBlt hdc, tR.left, tR.top, 3&, tR.Bottom - tR.top, _
    m_hBitmapDC, offset, 0, 3&, 22&, vbSrcCopy
'Right side
  StretchBlt hdc, tR.Right - 3&, tR.top, 3&, tR.Bottom - tR.top, _
    m_hBitmapDC, offset + 18&, 0, 3&, 22&, vbSrcCopy
'interior
  StretchBlt hdc, tR.left + 3&, tR.top, tR.Right - tR.left - 6&, tR.Bottom - tR.top, _
    m_hBitmapDC, offset + 3&, 0, 16&, 22&, vbSrcCopy
End Sub

Private Sub DrawItemIcon(ByVal hdc As Long, ByVal left As Long, ByVal top As Long, _
  ByVal hIcon As Long, Optional IsEnable As Boolean = True)
  
  Dim hBr As Long
  
  If IsEnable Then
    DrawIconEx hdc, left, top, _
      hIcon, ICON_WIDTH, ICON_WIDTH, 0, 0, DI_NORMAL
  Else
    hBr = CreateSolidBrush(m_lClrSysShadow)
    
    DrawState hdc, ByVal hBr, ByVal 0, ByVal hIcon, _
      ByVal 0, left, top, ICON_WIDTH, ICON_WIDTH, DST_ICON Or DSS_MONO
    
    DeleteObject hBr
  End If
End Sub


