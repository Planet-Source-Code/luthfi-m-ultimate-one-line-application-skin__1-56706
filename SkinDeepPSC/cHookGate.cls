VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHookGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Hook Engine Callback gating

Option Explicit

'Stupid IDE Hack
#If False Then
  Dim ehtCWP, ehtGetMessage, ehtSubClass
#End If

Public Enum enHookType
  ehtCWP = WH_CALLWNDPROC
  ehtGetMessage = WH_GETMESSAGE
  ehtSubClass = -10
End Enum

Private m_Hook As cHook
Private lHookPtr As Long
Private m_hookType As enHookType

Implements IHook

Private Sub Class_Terminate()
  StopHook
End Sub

Private Sub IHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, nCode As Long, wParam As Long, lParam As Long)
  
  Dim Proc As cSkinProc
  
  'this line below is to prevent cyclic references, a bit slower but otherwise
  'the globalskin would never be destroyed in compiled runtime
  Set Proc = HookProc(lHookPtr)
  'callback to cSkinProc
  Proc.HookProc bBefore, bHandled, lReturn, nCode, wParam, lParam, m_hookType
End Sub

Public Sub StartHook(ByVal HookType As enHookType, lHookProc As Long)
  'Start a new Hook and save the callback info
  lHookPtr = lHookProc 'only save the ptr
  Set m_Hook = New cHook
  m_hookType = HookType
  
  m_Hook.Hook Me, HookType
End Sub

Public Sub StopHook()
  If Not m_Hook Is Nothing Then
    m_Hook.UnHook
    Set m_Hook = Nothing
  End If
  lHookPtr = 0
End Sub


