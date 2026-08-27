VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHookListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Routes Click/Change from a runtime-created ListBox to a form method.

Public WithEvents ListBox As MSForms.ListBox

Private mHandler As Object
Private mProc As String

Public Sub Init(ByVal lst As MSForms.ListBox, ByVal handler As Object, ByVal procName As String)
    Set ListBox = lst
    Set mHandler = handler
    mProc = procName
End Sub

Private Sub ListBox_Click()
    Notify
End Sub

Private Sub ListBox_Change()
    Notify
End Sub

Private Sub Notify()
    If mHandler Is Nothing Then Exit Sub
    If Len(mProc) = 0 Then Exit Sub
    CallByName mHandler, mProc, VbMethod
End Sub
