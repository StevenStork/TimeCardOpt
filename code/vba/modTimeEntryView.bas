Attribute VB_Name = "modTimeEntryView"
Option Explicit

'==============================================================================
' Time Entry View
'
' Hides Time Entry columns/rows that are outside the active work window:
'   - Date columns not in the current 2-week pay period
'   - Minute rows before 5:30 AM and from 9:00 PM onward
'
' Pay periods are exactly 14 days, anchored to Welcome!F3 (same date as
' Time Entry!B1 after PrepareTimeEntry runs).
'
' Call from the Time Entry sheet module:
'
'   Private Sub Worksheet_Activate()
'       ApplyTimeEntryView
'   End Sub
'==============================================================================

Private Const WELCOME_SHEET As String = "Welcome"
Private Const TIME_ENTRY_SHEET As String = "Time Entry"
Private Const WELCOME_DATE_CELL As String = "F3"
Private Const TIME_ENTRY_DATE_CELL As String = "B1"
Private Const DATE_HEADER_ROW As Long = 1
Private Const FIRST_DATE_COL As Long = 2          ' column B (start date in B1)
Private Const FOLLOW_ON_DATE_START_COL As Long = 3 ' column C (B1+1 ... B1+366)
Private Const DAYS_AFTER_START As Long = 366
Private Const PAY_PERIOD_DAYS As Long = 14

Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440
Private Const SHOW_FROM_MINUTE As Long = 330      ' 5:30 AM
Private Const HIDE_FROM_MINUTE As Long = 1260     ' 9:00 PM

'------------------------------------------------------------------------------
' Public entry point — safe to call from Worksheet_Activate.
'------------------------------------------------------------------------------
Public Sub ApplyTimeEntryView()
    Dim timeEntryWs As Worksheet
    Dim periodStart As Date
    Dim periodEnd As Date

    On Error GoTo CleanFail
    OptimizeExcel True

    Set timeEntryWs = ThisWorkbook.Worksheets(TIME_ENTRY_SHEET)

    GetCurrentPayPeriod periodStart, periodEnd
    ApplyPayPeriodColumns timeEntryWs, periodStart, periodEnd
    ApplyWorkHourRows timeEntryWs

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    MsgBox "ApplyTimeEntryView failed: " & Err.Description, vbExclamation, "Time Card"
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' Resolve the 14-day pay period that contains today's date.
' Period 0 starts on Welcome!F3; each later period advances by exactly 14 days.
'------------------------------------------------------------------------------
Private Sub GetCurrentPayPeriod(ByRef periodStart As Date, ByRef periodEnd As Date)
    Dim anchorDate As Date
    Dim todayDate As Date
    Dim daysSinceAnchor As Long
    Dim periodIndex As Long

    anchorDate = ReadAnchorDate()
    todayDate = Date

    daysSinceAnchor = CLng(todayDate - anchorDate)
    If daysSinceAnchor < 0 Then
        periodIndex = 0
    Else
        periodIndex = daysSinceAnchor \ PAY_PERIOD_DAYS
    End If

    periodStart = anchorDate + (periodIndex * PAY_PERIOD_DAYS)
    periodEnd = periodStart + (PAY_PERIOD_DAYS - 1)
End Sub

'------------------------------------------------------------------------------
' Prefer Time Entry!B1 (already synced); fall back to Welcome!F3.
'------------------------------------------------------------------------------
Private Function ReadAnchorDate() As Date
    Dim rawValue As Variant
    Dim timeEntryWs As Worksheet
    Dim welcomeWs As Worksheet

    Set timeEntryWs = ThisWorkbook.Worksheets(TIME_ENTRY_SHEET)
    rawValue = timeEntryWs.Range(TIME_ENTRY_DATE_CELL).Value

    If IsEmpty(rawValue) Or Not IsDate(rawValue) Then
        Set welcomeWs = ThisWorkbook.Worksheets(WELCOME_SHEET)
        rawValue = welcomeWs.Range(WELCOME_DATE_CELL).Value
    End If

    If IsEmpty(rawValue) Or Not IsDate(rawValue) Then
        Err.Raise vbObjectError + 2001, "ReadAnchorDate", _
            "Set a valid pay-period start date in Welcome!" & WELCOME_DATE_CELL & _
            " (or Time Entry!" & TIME_ENTRY_DATE_CELL & ")."
    End If

    ReadAnchorDate = CDate(rawValue)
End Function

'------------------------------------------------------------------------------
' Show only date columns whose header date falls in [periodStart, periodEnd].
' Column B uses B1; columns C onward use row-1 headers from PrepareTimeEntry.
'------------------------------------------------------------------------------
Private Sub ApplyPayPeriodColumns( _
    ByVal timeEntryWs As Worksheet, _
    ByVal periodStart As Date, _
    ByVal periodEnd As Date)

    Dim col As Long
    Dim lastDateCol As Long
    Dim colDate As Date
    Dim rawValue As Variant

    lastDateCol = FOLLOW_ON_DATE_START_COL + DAYS_AFTER_START - 1

    ' Reset visibility across the whole date block first.
    timeEntryWs.Range( _
        timeEntryWs.Columns(FIRST_DATE_COL), _
        timeEntryWs.Columns(lastDateCol)).Hidden = False

    For col = FIRST_DATE_COL To lastDateCol
        If col = FIRST_DATE_COL Then
            rawValue = timeEntryWs.Range(TIME_ENTRY_DATE_CELL).Value
        Else
            rawValue = timeEntryWs.Cells(DATE_HEADER_ROW, col).Value
        End If

        If IsEmpty(rawValue) Or Not IsDate(rawValue) Then
            timeEntryWs.Columns(col).Hidden = True
        Else
            colDate = CDate(rawValue)
            timeEntryWs.Columns(col).Hidden = _
                (colDate < periodStart Or colDate > periodEnd)
        End If
    Next col
End Sub

'------------------------------------------------------------------------------
' Hide minute rows outside the 5:30 AM – 9:00 PM work window.
' Visible rows are 5:30 AM through 8:59 PM inclusive.
'------------------------------------------------------------------------------
Private Sub ApplyWorkHourRows(ByVal timeEntryWs As Worksheet)
    Dim firstMinuteRow As Long
    Dim lastMinuteRow As Long
    Dim firstVisibleRow As Long
    Dim firstEveningRow As Long

    firstMinuteRow = MINUTE_START_ROW
    lastMinuteRow = MINUTE_START_ROW + MINUTES_PER_DAY - 1
    firstVisibleRow = MINUTE_START_ROW + SHOW_FROM_MINUTE
    firstEveningRow = MINUTE_START_ROW + HIDE_FROM_MINUTE

    ' Show the full minute block, then hide the two off-hour ranges.
    timeEntryWs.Range( _
        timeEntryWs.Rows(firstMinuteRow), _
        timeEntryWs.Rows(lastMinuteRow)).Hidden = False

    If firstVisibleRow > firstMinuteRow Then
        timeEntryWs.Range( _
            timeEntryWs.Rows(firstMinuteRow), _
            timeEntryWs.Rows(firstVisibleRow - 1)).Hidden = True
    End If

    If firstEveningRow <= lastMinuteRow Then
        timeEntryWs.Range( _
            timeEntryWs.Rows(firstEveningRow), _
            timeEntryWs.Rows(lastMinuteRow)).Hidden = True
    End If
End Sub
