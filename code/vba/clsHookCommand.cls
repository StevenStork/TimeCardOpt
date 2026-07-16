VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHookCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Routes Click from a runtime-created CommandButton to a form method.

Public WithEvents CommandButton As MSForms.CommandButton

Private mHandler As Object
Private mProc As String

Public Sub Init(ByVal btn As MSForms.CommandButton, ByVal handler As Object, ByVal procName As String)
    Set CommandButton = btn
    Set mHandler = handler
    mProc = procName
End Sub

Private Sub CommandButton_Click()
    If mHandler Is Nothing Then Exit Sub
    If Len(mProc) = 0 Then Exit Sub
    CallByName mHandler, mProc, VbMethod
End Sub
