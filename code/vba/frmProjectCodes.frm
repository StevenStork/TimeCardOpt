VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectCodes
   Caption         =   "Project Codes"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjectCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Editor for the Project Codes sheet.
'
' Left:  maintain the master list of direct charge codes (column B)
' Right: maintain projects (columns D+) with title + associated direct codes
'
' Import notes:
'   1. Import modProjectCodes.bas, clsHookCommand.cls, clsHookListBox.cls
'   2. In VBA: Insert > UserForm, rename it to frmProjectCodes
'   3. Paste this form's code into the form module (below Option Explicit)
'   4. Or import this .frm if your Excel accepts text-only UserForm exports
'
' The UI is built at runtime in UserForm_Initialize, so an empty UserForm
' shell named frmProjectCodes is enough.
'==============================================================================

Private Const CTRL_DIRECT_LIST As String = "lstDirectCodes"
Private Const CTRL_DIRECT_TEXT As String = "txtDirectCode"
Private Const CTRL_PROJECT_LIST As String = "lstProjects"
Private Const CTRL_PROJECT_TITLE As String = "txtProjectTitle"
Private Const CTRL_ASSOC_LIST As String = "lstAssociatedCodes"

Private mLoading As Boolean
Private mHooks As Collection

'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set mHooks = New Collection
    EnsureUi
    EnsureProjectCodesLayout
    RefreshAll
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set mHooks = Nothing
End Sub

'------------------------------------------------------------------------------
Private Sub EnsureUi()
    Me.Caption = "Project Codes"
    Me.Width = 520
    Me.Height = 400

    If ControlExists(CTRL_DIRECT_LIST) Then
        HookExistingControls
        Exit Sub
    End If

    AddLabel "lblDirectHeader", "Direct Charge Codes", 12, 12, 200, 18
    AddTextBox CTRL_DIRECT_TEXT, 12, 36, 150, 22
    HookButton AddButton("cmdAddDirect", "Add", 168, 36, 50, 22), "OnAddDirect"
    HookButton AddButton("cmdRemoveDirect", "Remove", 222, 36, 60, 22), "OnRemoveDirect"
    AddListBox CTRL_DIRECT_LIST, 12, 66, 270, 250, False

    AddLabel "lblProjectHeader", "Projects", 300, 12, 200, 18
    AddTextBox CTRL_PROJECT_TITLE, 300, 36, 150, 22
    HookButton AddButton("cmdAddProject", "Add", 456, 36, 50, 22), "OnAddProject"
    HookList AddListBox(CTRL_PROJECT_LIST, 300, 66, 206, 90, False), "OnProjectSelected"

    AddLabel "lblAssocHeader", "Associated Direct Codes", 300, 168, 200, 18
    AddListBox CTRL_ASSOC_LIST, 300, 192, 206, 114, True

    HookButton AddButton("cmdSaveProject", "Save Project", 300, 318, 100, 24), "OnSaveProject"
    HookButton AddButton("cmdDeleteProject", "Delete Project", 406, 318, 100, 24), "OnDeleteProject"
    HookButton AddButton("cmdClose", "Close", 300, 348, 206, 24), "OnCloseForm"
End Sub

Private Sub HookExistingControls()
    On Error Resume Next
    HookButton Me.Controls("cmdAddDirect"), "OnAddDirect"
    HookButton Me.Controls("cmdRemoveDirect"), "OnRemoveDirect"
    HookButton Me.Controls("cmdAddProject"), "OnAddProject"
    HookButton Me.Controls("cmdSaveProject"), "OnSaveProject"
    HookButton Me.Controls("cmdDeleteProject"), "OnDeleteProject"
    HookButton Me.Controls("cmdClose"), "OnCloseForm"
    HookList Me.Controls(CTRL_PROJECT_LIST), "OnProjectSelected"
    On Error GoTo 0
End Sub

Private Function ControlExists(ByVal controlName As String) As Boolean
    Dim ctl As Object
    On Error Resume Next
    Set ctl = Me.Controls(controlName)
    ControlExists = Not ctl Is Nothing
    On Error GoTo 0
End Function

Private Sub AddLabel(ByVal name As String, ByVal caption As String, _
    ByVal leftPos As Single, ByVal topPos As Single, _
    ByVal widthPos As Single, ByVal heightPos As Single)

    Dim ctl As MSForms.Label
    Set ctl = Me.Controls.Add("Forms.Label.1", name, True)
    ctl.Caption = caption
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
End Sub

Private Sub AddTextBox(ByVal name As String, _
    ByVal leftPos As Single, ByVal topPos As Single, _
    ByVal widthPos As Single, ByVal heightPos As Single)

    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls.Add("Forms.TextBox.1", name, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
End Sub

Private Function AddButton(ByVal name As String, ByVal caption As String, _
    ByVal leftPos As Single, ByVal topPos As Single, _
    ByVal widthPos As Single, ByVal heightPos As Single) As MSForms.CommandButton

    Dim ctl As MSForms.CommandButton
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", name, True)
    ctl.Caption = caption
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    Set AddButton = ctl
End Function

Private Function AddListBox(ByVal name As String, _
    ByVal leftPos As Single, ByVal topPos As Single, _
    ByVal widthPos As Single, ByVal heightPos As Single, _
    ByVal multiSelect As Boolean) As MSForms.ListBox

    Dim ctl As MSForms.ListBox
    Set ctl = Me.Controls.Add("Forms.ListBox.1", name, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    If multiSelect Then
        ctl.MultiSelect = fmMultiSelectMulti
    Else
        ctl.MultiSelect = fmMultiSelectSingle
    End If
    Set AddListBox = ctl
End Function

Private Sub HookButton(ByVal btn As MSForms.CommandButton, ByVal procName As String)
    Dim hook As clsHookCommand
    If btn Is Nothing Then Exit Sub
    Set hook = New clsHookCommand
    hook.Init btn, Me, procName
    mHooks.Add hook
End Sub

Private Sub HookList(ByVal lst As MSForms.ListBox, ByVal procName As String)
    Dim hook As clsHookListBox
    If lst Is Nothing Then Exit Sub
    Set hook = New clsHookListBox
    hook.Init lst, Me, procName
    mHooks.Add hook
End Sub

'------------------------------------------------------------------------------
Private Sub txtDirectCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        OnAddDirect
    End If
End Sub

Private Sub txtProjectTitle_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        OnAddProject
    End If
End Sub

' Designer-control fallbacks (used if the form is laid out manually).
Private Sub cmdAddDirect_Click()
    OnAddDirect
End Sub
Private Sub cmdRemoveDirect_Click()
    OnRemoveDirect
End Sub
Private Sub cmdAddProject_Click()
    OnAddProject
End Sub
Private Sub cmdSaveProject_Click()
    OnSaveProject
End Sub
Private Sub cmdDeleteProject_Click()
    OnDeleteProject
End Sub
Private Sub cmdClose_Click()
    OnCloseForm
End Sub
Private Sub lstProjects_Click()
    OnProjectSelected
End Sub
Private Sub lstProjects_Change()
    OnProjectSelected
End Sub

'------------------------------------------------------------------------------
Private Sub RefreshAll()
    Dim selectedProject As String

    selectedProject = SelectedProjectTitle()
    RefreshDirectList
    RefreshProjectList selectedProject
    RefreshAssociationList
End Sub

Private Sub RefreshDirectList()
    Dim codes As Variant
    Dim i As Long
    Dim lst As MSForms.ListBox

    Set lst = Me.Controls(CTRL_DIRECT_LIST)
    lst.Clear

    codes = LoadDirectCodes()
    For i = 1 To VariantLen(codes)
        lst.AddItem CStr(VariantItem(codes, i))
    Next i
End Sub

Private Sub RefreshProjectList(Optional ByVal selectTitle As String = "")
    Dim titles As Variant
    Dim i As Long
    Dim lst As MSForms.ListBox

    Set lst = Me.Controls(CTRL_PROJECT_LIST)
    mLoading = True
    lst.Clear

    titles = LoadProjectTitles()
    For i = 1 To VariantLen(titles)
        lst.AddItem CStr(VariantItem(titles, i))
        If Len(selectTitle) > 0 Then
            If StrComp(CStr(VariantItem(titles, i)), selectTitle, vbTextCompare) = 0 Then
                lst.ListIndex = i - 1
            End If
        End If
    Next i

    mLoading = False

    If lst.ListCount > 0 And lst.ListIndex < 0 Then
        lst.ListIndex = 0
    End If
End Sub

Private Sub RefreshAssociationList()
    Dim lst As MSForms.ListBox
    Dim directs As Variant
    Dim associated As Variant
    Dim projectTitle As String
    Dim i As Long
    Dim j As Long
    Dim code As String
    Dim isSelected As Boolean

    Set lst = Me.Controls(CTRL_ASSOC_LIST)
    lst.Clear

    projectTitle = SelectedProjectTitle()
    directs = LoadDirectCodes()

    If Len(projectTitle) = 0 Or VariantLen(directs) = 0 Then Exit Sub

    associated = LoadProjectDirectCodes(projectTitle)

    For i = 1 To VariantLen(directs)
        code = CStr(VariantItem(directs, i))
        lst.AddItem code

        isSelected = False
        For j = 1 To VariantLen(associated)
            If StrComp(code, CStr(VariantItem(associated, j)), vbTextCompare) = 0 Then
                isSelected = True
                Exit For
            End If
        Next j
        lst.Selected(i - 1) = isSelected
    Next i
End Sub

Private Function SelectedProjectTitle() As String
    Dim lst As MSForms.ListBox

    Set lst = Me.Controls(CTRL_PROJECT_LIST)
    If lst.ListIndex < 0 Then
        SelectedProjectTitle = vbNullString
    Else
        SelectedProjectTitle = CStr(lst.List(lst.ListIndex))
    End If
End Function

Private Function SelectedDirectCode() As String
    Dim lst As MSForms.ListBox

    Set lst = Me.Controls(CTRL_DIRECT_LIST)
    If lst.ListIndex < 0 Then
        SelectedDirectCode = vbNullString
    Else
        SelectedDirectCode = CStr(lst.List(lst.ListIndex))
    End If
End Function

'------------------------------------------------------------------------------
' Public so clsHookCommand / clsHookListBox can CallByName them.
'------------------------------------------------------------------------------
Public Sub OnAddDirect()
    Dim txt As MSForms.TextBox
    Dim code As String

    On Error GoTo Fail
    Set txt = Me.Controls(CTRL_DIRECT_TEXT)
    code = Trim$(CStr(txt.Value))

    If Not AddDirectCode(code) Then
        MsgBox "That direct charge code already exists.", vbInformation, "Project Codes"
        Exit Sub
    End If

    txt.Value = vbNullString
    RefreshAll
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Project Codes"
End Sub

Public Sub OnRemoveDirect()
    Dim code As String

    On Error GoTo Fail
    code = SelectedDirectCode()
    If Len(code) = 0 Then
        MsgBox "Select a direct charge code to remove.", vbInformation, "Project Codes"
        Exit Sub
    End If

    If MsgBox("Remove direct charge code '" & code & "' from the list and all projects?", _
        vbYesNo + vbQuestion, "Project Codes") <> vbYes Then Exit Sub

    RemoveDirectCode code
    RefreshAll
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Project Codes"
End Sub

Public Sub OnAddProject()
    Dim txt As MSForms.TextBox
    Dim title As String

    On Error GoTo Fail
    Set txt = Me.Controls(CTRL_PROJECT_TITLE)
    title = Trim$(CStr(txt.Value))

    If Not AddProject(title) Then
        MsgBox "That project title already exists.", vbInformation, "Project Codes"
        Exit Sub
    End If

    txt.Value = vbNullString
    RefreshDirectList
    RefreshProjectList title
    RefreshAssociationList
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Project Codes"
End Sub

Public Sub OnSaveProject()
    Dim projectTitle As String
    Dim lst As MSForms.ListBox
    Dim i As Long
    Dim kept As Long
    Dim selected() As String

    On Error GoTo Fail
    projectTitle = SelectedProjectTitle()
    If Len(projectTitle) = 0 Then
        MsgBox "Select or add a project first.", vbInformation, "Project Codes"
        Exit Sub
    End If

    Set lst = Me.Controls(CTRL_ASSOC_LIST)
    If lst.ListCount = 0 Then
        SaveProjectAssociations projectTitle, Array()
        MsgBox "Saved project '" & projectTitle & "' with no associated codes.", _
            vbInformation, "Project Codes"
        Exit Sub
    End If

    ReDim selected(1 To lst.ListCount)
    kept = 0
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            kept = kept + 1
            selected(kept) = CStr(lst.List(i))
        End If
    Next i

    If kept = 0 Then
        SaveProjectAssociations projectTitle, Array()
    Else
        ReDim Preserve selected(1 To kept)
        SaveProjectAssociations projectTitle, selected
    End If

    MsgBox "Saved associations for '" & projectTitle & "'.", vbInformation, "Project Codes"
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Project Codes"
End Sub

Public Sub OnDeleteProject()
    Dim projectTitle As String

    On Error GoTo Fail
    projectTitle = SelectedProjectTitle()
    If Len(projectTitle) = 0 Then
        MsgBox "Select a project to delete.", vbInformation, "Project Codes"
        Exit Sub
    End If

    If MsgBox("Delete project '" & projectTitle & "'?", vbYesNo + vbQuestion, "Project Codes") <> vbYes Then
        Exit Sub
    End If

    DeleteProject projectTitle
    RefreshAll
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Project Codes"
End Sub

Public Sub OnProjectSelected()
    Dim txt As MSForms.TextBox

    If mLoading Then Exit Sub
    Set txt = Me.Controls(CTRL_PROJECT_TITLE)
    txt.Value = SelectedProjectTitle()
    RefreshAssociationList
End Sub

Public Sub OnCloseForm()
    Unload Me
End Sub
