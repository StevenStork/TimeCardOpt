VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHookCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Routes Change from a runtime-created ComboBox to a form method.

Public WithEvents ComboBox As MSForms.ComboBox

Private mHandler As Object
Private mProc As String

Public Sub Init(ByVal cbo As MSForms.ComboBox, ByVal handler As Object, ByVal procName As String)
    Set ComboBox = cbo
    Set mHandler = handler
    mProc = procName
End Sub

Private Sub ComboBox_Change()
    If mHandler Is Nothing Then Exit Sub
    If Len(mProc) = 0 Then Exit Sub
    CallByName mHandler, mProc, VbMethod
End Sub
