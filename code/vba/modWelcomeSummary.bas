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
' Each refresh:
'   1. Saves rows 1-3 (values + formatting)
'   2. Clears the Welcome sheet
'   3. Restores rows 1-3
'   4. Rebuilds the formatted project summary table
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
Private Const SUMMARY_TITLE As String = "Projects"
Private Const SUMMARY_START_ROW As Long = 5
Private Const SUMMARY_DATA_START_ROW As Long = 6
Private Const SUMMARY_LABEL_COL As Long = 2          ' B
Private Const SUMMARY_FIRST_DATE_COL As Long = 3     ' C
Private Const SUMMARY_LAST_DATE_COL As Long = 16     ' P
Private Const PAY_PERIOD_DAYS As Long = 14
Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440
Private Const PRESERVE_ROWS As Long = 3
Private Const TEMP_SHEET_NAME As String = "__WelcomeRowsTemp"

Private Const HEADER_FILL_RGB As Long = 4737096      ' RGB(72, 100, 120)
Private Const HEADER_FONT_RGB As Long = 16777215     ' white
Private Const ALT_ROW_FILL_RGB As Long = 15132390    ' RGB(230, 236, 240)
Private Const BORDER_RGB As Long = 11316396          ' RGB(140, 156, 172)

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
    Dim lastDataRow As Long

    On Error GoTo CleanFail
    OptimizeExcel True

    Set welcomeWs = WelcomeSheet()
    Set timeEntryWs = TimeEntrySheet()

    ' Capture period inputs before the sheet is cleared.
    GetPayPeriodForDate ReadWelcomeDate(welcomeWs, PERIOD_DATE_CELL), _
        ReadWelcomeDate(welcomeWs, ANCHOR_DATE_CELL), _
        periodStart, periodEnd

    BuildPeriodDates periodStart, periodDates
    projectTitles = LoadProjectTitles()
    projectCount = VariantLen(projectTitles)

    ResetWelcomeSheetPreservingTopRows welcomeWs

    welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LABEL_COL).Value = SUMMARY_TITLE
    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = _
            periodDates(dayIndex)
    Next dayIndex

    lastDataRow = SUMMARY_START_ROW

    If projectCount > 0 Then
        ReDim minuteCounts(1 To projectCount, 1 To PAY_PERIOD_DAYS)
        AccumulateProjectMinutes timeEntryWs, periodDates, projectTitles, minuteCounts

        outRow = SUMMARY_DATA_START_ROW
        For projectIndex = 1 To projectCount
            totalMinutes = 0
            For dayIndex = 1 To PAY_PERIOD_DAYS
                totalMinutes = totalMinutes + minuteCounts(projectIndex, dayIndex)
            Next dayIndex

            If totalMinutes > 0 Then
                welcomeWs.Cells(outRow, SUMMARY_LABEL_COL).Value = _
                    CStr(VariantItem(projectTitles, projectIndex))

                For dayIndex = 1 To PAY_PERIOD_DAYS
                    hoursValue = Application.WorksheetFunction.Round( _
                        minuteCounts(projectIndex, dayIndex) / 60#, 2)
                    welcomeWs.Cells(outRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = hoursValue
                Next dayIndex

                lastDataRow = outRow
                outRow = outRow + 1
            End If
        Next projectIndex
    End If

    FormatSummaryTable welcomeWs, lastDataRow

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    MsgBox "UpdateWelcomeSummary failed: " & Err.Description, vbExclamation, "Welcome Summary"
End Sub

'------------------------------------------------------------------------------
' Save rows 1-3 (all columns, values + formatting), clear the sheet, restore.
'------------------------------------------------------------------------------
Private Sub ResetWelcomeSheetPreservingTopRows(ByVal welcomeWs As Worksheet)
    Dim tempWs As Worksheet

    RemoveTempSheetIfPresent

    Set tempWs = ThisWorkbook.Worksheets.Add(After:=welcomeWs)
    tempWs.Name = TEMP_SHEET_NAME
    tempWs.Visible = xlSheetVeryHidden

    welcomeWs.Range("1:" & CStr(PRESERVE_ROWS)).Copy Destination:=tempWs.Range("A1")

    welcomeWs.Cells.Clear

    tempWs.Range("1:" & CStr(PRESERVE_ROWS)).Copy Destination:=welcomeWs.Range("A1")

    RemoveTempSheetIfPresent
End Sub

'------------------------------------------------------------------------------
Private Sub RemoveTempSheetIfPresent()
    Dim tempWs As Worksheet

    On Error Resume Next
    Set tempWs = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)
    On Error GoTo 0

    If Not tempWs Is Nothing Then
        tempWs.Visible = xlSheetVisible
        tempWs.Delete
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub FormatSummaryTable(ByVal welcomeWs As Worksheet, ByVal lastDataRow As Long)
    Dim headerRange As Range
    Dim tableRange As Range
    Dim dataRange As Range
    Dim hoursRange As Range
    Dim labelRange As Range
    Dim r As Long
    Dim dayIndex As Long

    Set headerRange = welcomeWs.Range( _
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LABEL_COL), _
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LAST_DATE_COL))

    With headerRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = HEADER_FONT_RGB
        .Interior.Color = HEADER_FILL_RGB
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LABEL_COL).HorizontalAlignment = xlLeft

    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_FIRST_DATE_COL + dayIndex - 1).NumberFormat = _
            "mmm d"
    Next dayIndex

    welcomeWs.Rows(SUMMARY_START_ROW).RowHeight = 30

    If lastDataRow >= SUMMARY_DATA_START_ROW Then
        Set dataRange = welcomeWs.Range( _
            welcomeWs.Cells(SUMMARY_DATA_START_ROW, SUMMARY_LABEL_COL), _
            welcomeWs.Cells(lastDataRow, SUMMARY_LAST_DATE_COL))
        Set hoursRange = welcomeWs.Range( _
            welcomeWs.Cells(SUMMARY_DATA_START_ROW, SUMMARY_FIRST_DATE_COL), _
            welcomeWs.Cells(lastDataRow, SUMMARY_LAST_DATE_COL))
        Set labelRange = welcomeWs.Range( _
            welcomeWs.Cells(SUMMARY_DATA_START_ROW, SUMMARY_LABEL_COL), _
            welcomeWs.Cells(lastDataRow, SUMMARY_LABEL_COL))

        With dataRange
            .Font.Name = "Calibri"
            .Font.Size = 10
            .VerticalAlignment = xlCenter
        End With

        labelRange.HorizontalAlignment = xlLeft
        labelRange.IndentLevel = 1

        With hoursRange
            .HorizontalAlignment = xlCenter
            .NumberFormat = "0.00"
        End With

        For r = SUMMARY_DATA_START_ROW To lastDataRow
            If ((r - SUMMARY_DATA_START_ROW) Mod 2) = 1 Then
                welcomeWs.Range( _
                    welcomeWs.Cells(r, SUMMARY_LABEL_COL), _
                    welcomeWs.Cells(r, SUMMARY_LAST_DATE_COL)).Interior.Color = ALT_ROW_FILL_RGB
            End If
            welcomeWs.Rows(r).RowHeight = 20
        Next r
    End If

    Set tableRange = welcomeWs.Range( _
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LABEL_COL), _
        welcomeWs.Cells(lastDataRow, SUMMARY_LAST_DATE_COL))

    With tableRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = BORDER_RGB
    End With

    welcomeWs.Columns(SUMMARY_LABEL_COL).ColumnWidth = 22
    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Columns(SUMMARY_FIRST_DATE_COL + dayIndex - 1).ColumnWidth = 8
    Next dayIndex
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
