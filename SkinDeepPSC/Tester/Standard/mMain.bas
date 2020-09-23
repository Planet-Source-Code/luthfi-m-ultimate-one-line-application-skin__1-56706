Attribute VB_Name = "mMain"
Option Explicit

Sub Main()

  Dim mainForm As frmMain
  
  'start skin engine, call this before showing/loading any form
  SkinDeep.Integrate
  
  'show our main form
  Set mainForm = New frmMain
  mainForm.Show
  Set mainForm = Nothing

'Or if you prefer an even shorter version, only two lines of code!!
'  Integrate
'  frmMain.Show
  
End Sub
