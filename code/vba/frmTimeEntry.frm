VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimeEntry
   Caption         =   "Enter Hours"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' Enter hours against a date, a project OR a direct charge code, and a
' start/end time. Writes that project title or charge code into matching
' Time Entry cells and blocks the save if any of those cells already have
' a value.
'
' Setup:
'   Insert an empty UserForm named frmTimeEntry and paste from Option Explicit
'   down, or import this .frm if supported. Also import modTimeEntryInput.bas.
'==============================================================================

Private Const CTRL_DATE As String = "cboDate"
Private Const CTRL_PROJECT As String = "cboProject"
Private Const CTRL_CHARGE As String = "cboChargeCode"
Private Const CTRL_START As String = "cboStartTime"
Private Const CTRL_END As String = "cboEndTime"

Private Const PROJECT_NONE As String = "(None)"
Private Const CHARGE_NONE As String = "(None)"

Private mLoading As Boolean
Private mHooks As Collection

'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set mHooks = New Collection
    EnsureUi
    LoadFormValues
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set mHooks = Nothing
End Sub

'------------------------------------------------------------------------------
Private Sub EnsureUi()
    Me.Caption = "Enter Hours"
    Me.Width = 420
    Me.Height = 360

    If ControlExists(CTRL_DATE) Then
        HookExistingControls
        Exit Sub
    End If

    AddLabel "lblDate", "Date", 24, 18, 360, 18
    AddCombo CTRL_DATE, 24, 40, 360, 24

    AddLabel "lblProject", "Project", 24, 78, 360, 18
    HookCombo AddCombo(CTRL_PROJECT, 24, 100, 360, 24), "OnProjectChanged"

    AddLabel "lblCharge", "Charge Code", 24, 138, 360, 18
    HookCombo AddCombo(CTRL_CHARGE, 24, 160, 360, 24), "OnChargeChanged"

    AddLabel "lblStart", "Start Time", 24, 198, 170, 18
    AddCombo CTRL_START, 24, 220, 170, 24

    AddLabel "lblEnd", "End Time", 214, 198, 170, 18
    AddCombo CTRL_END, 214, 220, 170, 24

    HookButton AddButton("cmdEnter", "Enter Hours", 24, 268, 170, 32), "OnEnterHours"
    HookButton AddButton("cmdClose", "Close", 214, 268, 170, 32), "OnCloseForm"
End Sub

Private Sub HookExistingControls()
    On Error Resume Next
    HookCombo Me.Controls(CTRL_PROJECT), "OnProjectChanged"
    HookCombo Me.Controls(CTRL_CHARGE), "OnChargeChanged"
    HookButton Me.Controls("cmdEnter"), "OnEnterHours"
    HookButton Me.Controls("cmdClose"), "OnCloseForm"
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

Private Function AddCombo(ByVal name As String, _
    ByVal leftPos As Single, ByVal topPos As Single, _
    ByVal widthPos As Single, ByVal heightPos As Single) As MSForms.ComboBox

    Dim ctl As MSForms.ComboBox
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", name, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    ctl.Style = fmStyleDropDownCombo
    Set AddCombo = ctl
End Function

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

Private Sub HookButton(ByVal btn As MSForms.CommandButton, ByVal procName As String)
    Dim hook As clsHookCommand
    If btn Is Nothing Then Exit Sub
    Set hook = New clsHookCommand
    hook.Init btn, Me, procName
    mHooks.Add hook
End Sub

Private Sub HookCombo(ByVal cbo As MSForms.ComboBox, ByVal procName As String)
    Dim hook As clsHookCombo
    If cbo Is Nothing Then Exit Sub
    Set hook = New clsHookCombo
    hook.Init cbo, Me, procName
    mHooks.Add hook
End Sub

' Designer-control fallbacks
Private Sub cboProject_Change()
    OnProjectChanged
End Sub
Private Sub cboChargeCode_Change()
    OnChargeChanged
End Sub
Private Sub cmdEnter_Click()
    OnEnterHours
End Sub
Private Sub cmdClose_Click()
    OnCloseForm
End Sub

'------------------------------------------------------------------------------
Private Sub LoadFormValues()
    mLoading = True
    LoadDates
    LoadProjects
    LoadChargeCodes
    LoadTimes
    mLoading = False
End Sub

Private Sub LoadDates()
    Dim cbo As MSForms.ComboBox
    Dim dates As Variant
    Dim i As Long
    Dim todayText As String
    Dim matched As Boolean

    Set cbo = Me.Controls(CTRL_DATE)
    cbo.Clear
    todayText = Format$(Date, "mm/dd/yyyy")
    matched = False

    On Error Resume Next
    dates = LoadTimeEntryDates()
    On Error GoTo 0

    For i = 1 To VariantLen(dates)
        cbo.AddItem Format$(CDate(VariantItem(dates, i)), "mm/dd/yyyy")
        If CDate(VariantItem(dates, i)) = Date Then
            cbo.ListIndex = i - 1
            matched = True
        End If
    Next i

    If Not matched Then
        cbo.Value = todayText
    End If
End Sub

Private Sub LoadProjects()
    Dim cbo As MSForms.ComboBox
    Dim titles As Variant
    Dim i As Long

    Set cbo = Me.Controls(CTRL_PROJECT)
    cbo.Clear
    cbo.AddItem PROJECT_NONE

    titles = LoadProjectTitles()
    For i = 1 To VariantLen(titles)
        cbo.AddItem CStr(VariantItem(titles, i))
    Next i

    cbo.ListIndex = 0
End Sub

Private Sub LoadChargeCodes()
    Dim cbo As MSForms.ComboBox
    Dim codes As Variant
    Dim i As Long

    Set cbo = Me.Controls(CTRL_CHARGE)
    cbo.Clear
    cbo.AddItem CHARGE_NONE

    codes = LoadDirectCodes()
    For i = 1 To VariantLen(codes)
        cbo.AddItem CStr(VariantItem(codes, i))
    Next i

    cbo.ListIndex = 0
End Sub

Private Sub LoadTimes()
    Dim startCbo As MSForms.ComboBox
    Dim endCbo As MSForms.ComboBox
    Dim minuteValue As Long
    Dim labelText As String

    Set startCbo = Me.Controls(CTRL_START)
    Set endCbo = Me.Controls(CTRL_END)
    startCbo.Clear
    endCbo.Clear

    ' Offer 15-minute steps across the visible work window; typing other times is allowed.
    For minuteValue = 330 To 1260 Step 15
        labelText = Format$(TimeSerial(0, 0, 0) + (minuteValue / 1440#), "h:mm AM/PM")
        startCbo.AddItem labelText
        endCbo.AddItem labelText
    Next minuteValue

    startCbo.Value = Format$(TimeSerial(8, 0, 0), "h:mm AM/PM")
    endCbo.Value = Format$(TimeSerial(17, 0, 0), "h:mm AM/PM")
End Sub

Private Function SelectedProjectTitle() As String
    Dim valueText As String

    valueText = Trim$(CStr(Me.Controls(CTRL_PROJECT).Value))
    If Len(valueText) = 0 Or StrComp(valueText, PROJECT_NONE, vbTextCompare) = 0 Then
        SelectedProjectTitle = vbNullString
    Else
        SelectedProjectTitle = valueText
    End If
End Function

Private Function SelectedChargeCode() As String
    Dim valueText As String

    valueText = Trim$(CStr(Me.Controls(CTRL_CHARGE).Value))
    If Len(valueText) = 0 Or StrComp(valueText, CHARGE_NONE, vbTextCompare) = 0 Then
        SelectedChargeCode = vbNullString
    Else
        SelectedChargeCode = valueText
    End If
End Function

'------------------------------------------------------------------------------
' Project and charge code are mutually exclusive: picking one clears the other.
'------------------------------------------------------------------------------
Public Sub OnProjectChanged()
    If mLoading Then Exit Sub
    If Len(SelectedProjectTitle()) = 0 Then Exit Sub

    mLoading = True
    Me.Controls(CTRL_CHARGE).ListIndex = 0
    mLoading = False
End Sub

Public Sub OnChargeChanged()
    If mLoading Then Exit Sub
    If Len(SelectedChargeCode()) = 0 Then Exit Sub

    mLoading = True
    Me.Controls(CTRL_PROJECT).ListIndex = 0
    mLoading = False
End Sub

Public Sub OnEnterHours()
    Dim entryDate As Date
    Dim entryValue As String
    Dim startTime As Date
    Dim endTime As Date
    Dim minutes As Long
    Dim projectTitle As String
    Dim chargeCode As String

    On Error GoTo Fail

    entryDate = ParseDateInput(CStr(Me.Controls(CTRL_DATE).Value))
    projectTitle = SelectedProjectTitle()
    chargeCode = SelectedChargeCode()
    startTime = ParseTimeInput(CStr(Me.Controls(CTRL_START).Value))
    endTime = ParseTimeInput(CStr(Me.Controls(CTRL_END).Value))

    If Len(projectTitle) > 0 Then
        entryValue = projectTitle
    ElseIf Len(chargeCode) > 0 Then
        entryValue = chargeCode
    Else
        Err.Raise vbObjectError + 4010, "OnEnterHours", _
            "Select a project or a charge code."
    End If

    EnterTimeRange entryDate, entryValue, startTime, endTime

    minutes = MinuteOfDay(endTime) - MinuteOfDay(startTime)
    MsgBox "Entered " & entryValue & " for " & Format$(entryDate, "mm/dd/yyyy") & _
        " (" & Format$(minutes / 60#, "0.##") & " hours).", _
        vbInformation, "Enter Hours"
    Exit Sub

Fail:
    MsgBox Err.Description, vbExclamation, "Enter Hours"
End Sub

Public Sub OnCloseForm()
    Unload Me
End Sub
