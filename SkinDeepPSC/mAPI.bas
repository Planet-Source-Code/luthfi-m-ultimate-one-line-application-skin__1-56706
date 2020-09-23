Attribute VB_Name = "mAPI"
Option Explicit

Public Const WH_CALLWNDPROC As Long = 4
Public Const WH_GETMESSAGE As Long = 3
Public Const WH_JOURNALPLAYBACK As Long = 1
Public Const WH_JOURNALRECORD As Long = 0
Public Const WH_SYSMSGFILTER As Long = 6
Public Const HC_ACTION As Long = 0

Public Const RDW_ALLCHILDREN As Long = &H80
Public Const RDW_INVALIDATE As Long = &H1
Public Const RDW_UPDATENOW As Long = &H100
Public Const RDW_FRAME As Long = &H400

Public Const WM_CREATE As Long = &H1
Public Const WM_DRAWITEM As Long = &H2B
Public Const WM_STYLECHANGED As Long = &H7D
Public Const WM_NCPAINT As Long = &H85
Public Const WM_MOUSEMOVE As Long = &H200

Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_STYLE As Long = -16
Public Const GWL_WNDPROC As Long = -4

Public Const TRANSPARENT As Long = 1
Public Const DI_NORMAL As Long = &H3
Public Const DST_ICON As Long = &H3
Public Const DSS_MONO As Long = &H80
Public Const DT_CENTER As Long = &H1
Public Const DT_VCENTER As Long = &H4
Public Const DT_SINGLELINE As Long = &H20

Public Const WS_EX_CLIENTEDGE As Long = &H200&
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const BS_OWNERDRAW As Long = &HB&
Public Const BST_CHECKED As Long = &H1

Public Const ODS_HOTLIGHT As Long = &H40
Public Const ODT_BUTTON As Long = 4
Public Const ODS_DISABLED As Long = &H4
Public Const ODS_FOCUS As Long = &H10
Public Const PS_SOLID As Long = 0

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_NOZORDER As Long = &H4

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type CWPSTRUCT
  lParam As Long
  wParam As Long
  message As Long
  hwnd As Long
End Type

Public Type msg
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type

Public Type CREATESTRUCT
  lpCreateParams As Long
  hInstance As Long
  hMenu As Long
  hWndParent As Long
  cy As Long
  cx As Long
  y As Long
  x As Long
  style As Long
  lpszName As Long
  lpszClass As Long
  dwExStyle As Long
End Type

Public Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  hDC As Long
  rcItem As RECT
  itemData As Long
End Type

Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type


Public Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

