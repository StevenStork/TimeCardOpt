Attribute VB_Name = "modWelcomeSummary"
Option Explicit

'==============================================================================
' Welcome Pay-Period Summary
'
' Rebuilds a project-hours summary on the Welcome sheet starting at B5.
' Call from the Welcome sheet module:
'
'   Private Sub Worksheet_Activate()
'       UpdateWelcomeSummary
'   End Sub
'
' Layout:
'   B5          = "Projects"
'   C5:P5       = the 14 dates in the pay period that contains Welcome!J3
'   B6...       = project titles charged in that pay period
'   C6:P...     = hours per project per day, rounded to the nearest hundredth
'
' Pay-period boundaries are anchored to Welcome!F3 (same rule as Time Entry).
'==============================================================================

Private Const WELCOME_SHEET As String = "Welcome"
Private Const PERIOD_DATE_CELL As String = "J3"
Private Const ANCHOR_DATE_CELL As String = "F3"
Private Const SUMMARY_TITLE_CELL As String = "B5"
Private Const SUMMARY_TITLE As String = "Projects"
Private Const SUMMARY_START_ROW As Long = 5
Private Const SUMMARY_DATA_START_ROW As Long = 6
Private Const SUMMARY_FIRST_DATE_COL As Long = 3   ' C
Private Const SUMMARY_LAST_DATE_COL As Long = 16   ' P
Private Const PAY_PERIOD_DAYS As Long = 14
Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440

'------------------------------------------------------------------------------
Public Sub UpdateWelcomeSummary()
    Dim welcomeWs As Worksheet
    Dim timeEntryWs As Worksheet
    Dim periodStart As Date
    Dim periodEnd As Date
    Dim periodDates() As Date
    Dim projectTitles As Variant
    Dim minuteCounts() As Long
    Dim projectCount As Long
    Dim dayIndex As Long
    Dim projectIndex As Long
    Dim totalMinutes As Long
    Dim outRow As Long
    Dim hoursValue As Double

    On Error GoTo CleanFail
    OptimizeExcel True

    Set welcomeWs = WelcomeSheet()
    Set timeEntryWs = TimeEntrySheet()

    GetPayPeriodForDate ReadWelcomeDate(welcomeWs, PERIOD_DATE_CELL), _
        ReadWelcomeDate(welcomeWs, ANCHOR_DATE_CELL), _
        periodStart, periodEnd

    BuildPeriodDates periodStart, periodDates
    projectTitles = LoadProjectTitles()
    projectCount = VariantLen(projectTitles)

    ClearSummaryArea welcomeWs

    welcomeWs.Range(SUMMARY_TITLE_CELL).Value = SUMMARY_TITLE
    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = _
            periodDates(dayIndex)
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_FIRST_DATE_COL + dayIndex - 1).NumberFormat = _
            "mm/dd/yyyy"
    Next dayIndex

    If projectCount = 0 Then GoTo CleanExit

    ReDim minuteCounts(1 To projectCount, 1 To PAY_PERIOD_DAYS)
    AccumulateProjectMinutes timeEntryWs, periodDates, projectTitles, minuteCounts

    outRow = SUMMARY_DATA_START_ROW
    For projectIndex = 1 To projectCount
        totalMinutes = 0
        For dayIndex = 1 To PAY_PERIOD_DAYS
            totalMinutes = totalMinutes + minuteCounts(projectIndex, dayIndex)
        Next dayIndex

        If totalMinutes > 0 Then
            welcomeWs.Cells(outRow, 2).Value = CStr(VariantItem(projectTitles, projectIndex))

            For dayIndex = 1 To PAY_PERIOD_DAYS
                If minuteCounts(projectIndex, dayIndex) > 0 Then
                    hoursValue = Application.WorksheetFunction.Round( _
                        minuteCounts(projectIndex, dayIndex) / 60#, 2)
                    welcomeWs.Cells(outRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = hoursValue
                    welcomeWs.Cells(outRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).NumberFormat = "0.00"
                End If
            Next dayIndex

            outRow = outRow + 1
        End If
    Next projectIndex

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    MsgBox "UpdateWelcomeSummary failed: " & Err.Description, vbExclamation, "Welcome Summary"
End Sub

'------------------------------------------------------------------------------
Private Function WelcomeSheet() As Worksheet
    On Error Resume Next
    Set WelcomeSheet = ThisWorkbook.Worksheets(WELCOME_SHEET)
    On Error GoTo 0

    If WelcomeSheet Is Nothing Then
        Err.Raise vbObjectError + 5001, "WelcomeSheet", _
            "Worksheet '" & WELCOME_SHEET & "' was not found."
    End If
End Function

'------------------------------------------------------------------------------
Private Function ReadWelcomeDate(ByVal welcomeWs As Worksheet, ByVal cellAddress As String) As Date
    Dim rawValue As Variant

    rawValue = welcomeWs.Range(cellAddress).Value
    If IsEmpty(rawValue) Or Not IsDate(rawValue) Then
        Err.Raise vbObjectError + 5002, "ReadWelcomeDate", _
            "Welcome!" & cellAddress & " must contain a valid date."
    End If

    ReadWelcomeDate = CDate(Int(CDate(rawValue)))
End Function

'------------------------------------------------------------------------------
Private Sub GetPayPeriodForDate( _
    ByVal referenceDate As Date, _
    ByVal anchorDate As Date, _
    ByRef periodStart As Date, _
    ByRef periodEnd As Date)

    Dim daysSinceAnchor As Long
    Dim periodIndex As Long

    daysSinceAnchor = CLng(referenceDate - anchorDate)
    If daysSinceAnchor < 0 Then
        periodIndex = 0
    Else
        periodIndex = daysSinceAnchor \ PAY_PERIOD_DAYS
    End If

    periodStart = anchorDate + (periodIndex * PAY_PERIOD_DAYS)
    periodEnd = periodStart + (PAY_PERIOD_DAYS - 1)
End Sub

'------------------------------------------------------------------------------
Private Sub BuildPeriodDates(ByVal periodStart As Date, ByRef periodDates() As Date)
    Dim i As Long

    ReDim periodDates(1 To PAY_PERIOD_DAYS)
    For i = 1 To PAY_PERIOD_DAYS
        periodDates(i) = periodStart + (i - 1)
    Next i
End Sub

'------------------------------------------------------------------------------
Private Sub ClearSummaryArea(ByVal welcomeWs As Worksheet)
    Dim lastRow As Long

    lastRow = welcomeWs.Cells(welcomeWs.Rows.Count, 2).End(xlUp).Row
    If lastRow < SUMMARY_START_ROW Then lastRow = SUMMARY_START_ROW

    ' Keep a generous clear window so shorter new summaries replace longer old ones.
    If lastRow < SUMMARY_START_ROW + 200 Then lastRow = SUMMARY_START_ROW + 200

    welcomeWs.Range( _
        welcomeWs.Cells(SUMMARY_START_ROW, 2), _
        welcomeWs.Cells(lastRow, SUMMARY_LAST_DATE_COL)).ClearContents
End Sub

'------------------------------------------------------------------------------
' Count one minute for each Time Entry cell whose value matches a project title.
'------------------------------------------------------------------------------
Private Sub AccumulateProjectMinutes( _
    ByVal timeEntryWs As Worksheet, _
    ByRef periodDates() As Date, _
    ByRef projectTitles As Variant, _
    ByRef minuteCounts() As Long)

    Dim dayIndex As Long
    Dim col As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim cellText As String
    Dim projectIndex As Long

    For dayIndex = 1 To PAY_PERIOD_DAYS
        col = FindDateColumn(timeEntryWs, periodDates(dayIndex))
        If col = 0 Then GoTo NextDay

        values = timeEntryWs.Range( _
            timeEntryWs.Cells(MINUTE_START_ROW, col), _
            timeEntryWs.Cells(MINUTE_START_ROW + MINUTES_PER_DAY - 1, col)).Value

        For rowIndex = 1 To MINUTES_PER_DAY
            cellText = Trim$(CStr(values(rowIndex, 1)))
            If Len(cellText) = 0 Then GoTo NextCell

            projectIndex = FindProjectIndex(projectTitles, cellText)
            If projectIndex > 0 Then
                minuteCounts(projectIndex, dayIndex) = minuteCounts(projectIndex, dayIndex) + 1
            End If
NextCell:
        Next rowIndex
NextDay:
    Next dayIndex
End Sub

'------------------------------------------------------------------------------
Private Function FindProjectIndex(ByRef projectTitles As Variant, ByVal cellText As String) As Long
    Dim i As Long

    For i = 1 To VariantLen(projectTitles)
        If StrComp(CStr(VariantItem(projectTitles, i)), cellText, vbTextCompare) = 0 Then
            FindProjectIndex = i
            Exit Function
        End If
    Next i

    FindProjectIndex = 0
End Function
