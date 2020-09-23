Attribute VB_Name = "mMain"
Option Explicit

Sub Main()

  Integrate
  'attach our own skin
  Set activeskin = New cMySkin
  frmMain.Show
  
End Sub
