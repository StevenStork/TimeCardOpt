Attribute VB_Name = "modTimeEntryInput"
Option Explicit

'==============================================================================
' Time Entry Input
'
' Writes a project title or direct charge code into the Time Entry minute
' grid for a date/time range.
'
'   ShowTimeEntryForm   launch the hours-entry UserForm
'   EnterTimeRange      validate and write the selected entry value
'
' Grid layout (from PrepareTimeEntry):
'   Time Entry!B1       = first date (column B)
'   Time Entry!C1...    = following dates
'   Time Entry!A2:A1441 = minutes 00:00 through 23:59
'
' Setup:
'   1. Import this module plus project-code helpers already in the workbook
'   2. Insert an empty UserForm named frmTimeEntry and paste code from
'      frmTimeEntry.frm (from Option Explicit down)
'   3. Run ShowTimeEntryForm
'==============================================================================

Private Const TIME_ENTRY_SHEET As String = "Time Entry"
Private Const TIME_ENTRY_DATE_CELL As String = "B1"
Private Const DATE_HEADER_ROW As Long = 1
Private Const FIRST_DATE_COL As Long = 2
Private Const FOLLOW_ON_DATE_START_COL As Long = 3
Private Const DAYS_AFTER_START As Long = 366
Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440

'------------------------------------------------------------------------------
Public Sub ShowTimeEntryForm()
    frmTimeEntry.Show vbModal
End Sub

'------------------------------------------------------------------------------
' Write entryValue (project title or charge code) into every minute cell for
' entryDate in [startTime, endTime).
' If allowOverwrite is False and cells are occupied, raises an error the form
' can catch; pass True after the user confirms overwrite.
'------------------------------------------------------------------------------
Public Sub EnterTimeRange( _
    ByVal entryDate As Date, _
    ByVal entryValue As String, _
    ByVal startTime As Date, _
    ByVal endTime As Date, _
    Optional ByVal allowOverwrite As Boolean = False)

    Dim ws As Worksheet
    Dim col As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim valueText As String
    Dim overlapAddress As String

    valueText = Trim$(entryValue)
    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 4001, "EnterTimeRange", _
            "Select a project or a charge code."
    End If

    ResolveTimeRangeBounds entryDate, startTime, endTime, ws, startRow, endRow, col

    If Not allowOverwrite Then
        overlapAddress = FindOccupiedCells(ws, startRow, endRow, col)
        If Len(overlapAddress) > 0 Then
            Err.Raise vbObjectError + 4004, "EnterTimeRange", _
                "Existing entries found in " & overlapAddress & "."
        End If
    End If

    WriteEntryValue ws, startRow, endRow, col, valueText
End Sub

'------------------------------------------------------------------------------
' Clear every minute cell for entryDate in [startTime, endTime).
'------------------------------------------------------------------------------
Public Sub ClearTimeRange( _
    ByVal entryDate As Date, _
    ByVal startTime As Date, _
    ByVal endTime As Date)

    Dim ws As Worksheet
    Dim col As Long
    Dim startRow As Long
    Dim endRow As Long

    ResolveTimeRangeBounds entryDate, startTime, endTime, ws, startRow, endRow, col
    ClearEntryRange ws, startRow, endRow, col
End Sub

'------------------------------------------------------------------------------
' Returns the address of occupied cells in the range, or vbNullString if empty.
'------------------------------------------------------------------------------
Public Function DescribeTimeRangeOverlap( _
    ByVal entryDate As Date, _
    ByVal startTime As Date, _
    ByVal endTime As Date) As String

    Dim ws As Worksheet
    Dim col As Long
    Dim startRow As Long
    Dim endRow As Long

    ResolveTimeRangeBounds entryDate, startTime, endTime, ws, startRow, endRow, col
    DescribeTimeRangeOverlap = FindOccupiedCells(ws, startRow, endRow, col)
End Function

'------------------------------------------------------------------------------
Private Sub ResolveTimeRangeBounds( _
    ByVal entryDate As Date, _
    ByVal startTime As Date, _
    ByVal endTime As Date, _
    ByRef ws As Worksheet, _
    ByRef startRow As Long, _
    ByRef endRow As Long, _
    ByRef col As Long)

    Dim startMinute As Long
    Dim endMinute As Long

    startMinute = MinuteOfDay(startTime)
    endMinute = MinuteOfDay(endTime)

    If endMinute <= startMinute Then
        Err.Raise vbObjectError + 4002, "ResolveTimeRangeBounds", _
            "End time must be after start time on the same day."
    End If

    Set ws = TimeEntrySheet()
    col = FindDateColumn(ws, entryDate)
    If col = 0 Then
        Err.Raise vbObjectError + 4003, "ResolveTimeRangeBounds", _
            "Date " & Format$(entryDate, "mm/dd/yyyy") & _
            " was not found on the Time Entry sheet. Run PrepareTimeEntry first."
    End If

    ' Rows cover minutes startMinute .. endMinute-1 inclusive.
    startRow = MINUTE_START_ROW + startMinute
    endRow = MINUTE_START_ROW + endMinute - 1
End Sub

'------------------------------------------------------------------------------
Public Function TimeEntrySheet() As Worksheet
    On Error Resume Next
    Set TimeEntrySheet = ThisWorkbook.Worksheets(TIME_ENTRY_SHEET)
    On Error GoTo 0

    If TimeEntrySheet Is Nothing Then
        Err.Raise vbObjectError + 4005, "TimeEntrySheet", _
            "Worksheet '" & TIME_ENTRY_SHEET & "' was not found."
    End If
End Function

'------------------------------------------------------------------------------
' Return available dates from the Time Entry header row (B1, then C1...).
'------------------------------------------------------------------------------
Public Function LoadTimeEntryDates() As Variant
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim count As Long
    Dim dates() As Date
    Dim rawValue As Variant

    Set ws = TimeEntrySheet()
    lastCol = FOLLOW_ON_DATE_START_COL + DAYS_AFTER_START - 1
    ReDim dates(1 To lastCol - FIRST_DATE_COL + 1)
    count = 0

    For col = FIRST_DATE_COL To lastCol
        If col = FIRST_DATE_COL Then
            rawValue = ws.Range(TIME_ENTRY_DATE_CELL).Value
        Else
            rawValue = ws.Cells(DATE_HEADER_ROW, col).Value
        End If

        If Not IsEmpty(rawValue) And IsDate(rawValue) Then
            count = count + 1
            dates(count) = CDate(rawValue)
        End If
    Next col

    If count = 0 Then
        LoadTimeEntryDates = Array()
    Else
        ReDim Preserve dates(1 To count)
        LoadTimeEntryDates = dates
    End If
End Function

'------------------------------------------------------------------------------
Public Function FindDateColumn(ByVal ws As Worksheet, ByVal entryDate As Date) As Long
    Dim col As Long
    Dim lastCol As Long
    Dim rawValue As Variant
    Dim colDate As Date

    lastCol = FOLLOW_ON_DATE_START_COL + DAYS_AFTER_START - 1

    For col = FIRST_DATE_COL To lastCol
        If col = FIRST_DATE_COL Then
            rawValue = ws.Range(TIME_ENTRY_DATE_CELL).Value
        Else
            rawValue = ws.Cells(DATE_HEADER_ROW, col).Value
        End If

        If Not IsEmpty(rawValue) And IsDate(rawValue) Then
            colDate = CDate(rawValue)
            If colDate = CDate(Int(entryDate)) Then
                FindDateColumn = col
                Exit Function
            End If
        End If
    Next col

    FindDateColumn = 0
End Function

'------------------------------------------------------------------------------
' Minutes from midnight for a time or datetime value (0-1439).
'------------------------------------------------------------------------------
Public Function MinuteOfDay(ByVal timeValue As Date) As Long
    Dim minutes As Long

    minutes = Int((CDbl(timeValue) - Int(CDbl(timeValue))) * MINUTES_PER_DAY + 0.5)
    If minutes >= MINUTES_PER_DAY Then minutes = 0
    If minutes < 0 Then minutes = 0
    MinuteOfDay = minutes
End Function

'------------------------------------------------------------------------------
' Robust time parsing for manual entry (8:30, 830, 8:30 AM, 17:30, etc.).
'------------------------------------------------------------------------------
Public Function ParseTimeInput(ByVal textValue As String) As Date
    Dim trimmed As String
    Dim hourPart As Long
    Dim minutePart As Long
    Dim isPM As Boolean
    Dim isAM As Boolean
    Dim colonPos As Long
    Dim leftPart As String
    Dim rightPart As String
    Dim digitsOnly As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim$(textValue)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 4006, "ParseTimeInput", "Enter a start and end time."
    End If

  ' Collapse internal spaces for suffix detection.
    Do While InStr(trimmed, "  ") > 0
        trimmed = Replace(trimmed, "  ", " ")
    Loop

    isPM = HasTimeSuffix(trimmed, "p")
    isAM = HasTimeSuffix(trimmed, "a")
    If isPM And isAM Then
        Err.Raise vbObjectError + 4011, "ParseTimeInput", _
            "'" & textValue & "' has conflicting AM/PM markers."
    End If

    colonPos = InStr(trimmed, ":")
    If colonPos > 0 Then
        leftPart = Trim$(Left$(trimmed, colonPos - 1))
        rightPart = Trim$(Mid$(trimmed, colonPos + 1))
        hourPart = ParseHourToken(leftPart, isAM, isPM)
        minutePart = ParseMinuteToken(rightPart, isAM, isPM)
    Else
        digitsOnly = ExtractTimeDigits(trimmed)
        If Len(digitsOnly) = 0 Then
            If IsDate(trimmed) Then
                ParseTimeInput = CDate(trimmed)
                Exit Function
            End If
            Err.Raise vbObjectError + 4007, "ParseTimeInput", _
                "'" & textValue & "' is not a valid time."
        End If

        If Len(digitsOnly) <= 2 Then
            hourPart = CLng(digitsOnly)
            minutePart = 0
        ElseIf Len(digitsOnly) = 3 Then
            hourPart = CLng(Left$(digitsOnly, 1))
            minutePart = CLng(Right$(digitsOnly, 2))
        ElseIf Len(digitsOnly) = 4 Then
            hourPart = CLng(Left$(digitsOnly, 2))
            minutePart = CLng(Right$(digitsOnly, 2))
        Else
            Err.Raise vbObjectError + 4007, "ParseTimeInput", _
                "'" & textValue & "' is not a valid time."
        End If

        hourPart = NormalizeHour(hourPart, isAM, isPM)
    End If

    If minutePart < 0 Or minutePart > 59 Then
        Err.Raise vbObjectError + 4012, "ParseTimeInput", _
            "Minutes must be between 0 and 59."
    End If

    If hourPart < 0 Or hourPart > 23 Then
        Err.Raise vbObjectError + 4013, "ParseTimeInput", _
            "Hours must be between 0 and 23."
    End If

    ParseTimeInput = TimeSerial(hourPart, minutePart, 0)
End Function

'------------------------------------------------------------------------------
' Return the latest minute-of-day with a charge on entryDate, or -1 if none.
'------------------------------------------------------------------------------
Public Function FindLatestChargedMinute(ByVal entryDate As Date) As Long
    Dim ws As Worksheet
    Dim col As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim minuteIndex As Long

    Set ws = TimeEntrySheet()
    col = FindDateColumn(ws, entryDate)
    If col = 0 Then
        FindLatestChargedMinute = -1
        Exit Function
    End If

    values = ws.Range( _
        ws.Cells(MINUTE_START_ROW, col), _
        ws.Cells(MINUTE_START_ROW + MINUTES_PER_DAY - 1, col)).Value

    For rowIndex = MINUTES_PER_DAY To 1 Step -1
        If Len(Trim$(CStr(values(rowIndex, 1)))) > 0 Then
            minuteIndex = rowIndex - 1
            FindLatestChargedMinute = minuteIndex
            Exit Function
        End If
    Next rowIndex

    FindLatestChargedMinute = -1
End Function

'------------------------------------------------------------------------------
Public Function DescribeLatestChargedTime(ByVal entryDate As Date) As String
    Dim minuteIndex As Long

    minuteIndex = FindLatestChargedMinute(entryDate)
    If minuteIndex < 0 Then
        DescribeLatestChargedTime = "None"
    Else
        DescribeLatestChargedTime = Format$( _
            TimeSerial(0, 0, 0) + (minuteIndex / MINUTES_PER_DAY), "h:mm AM/PM")
    End If
End Function

'------------------------------------------------------------------------------
Private Function ParseHourToken(ByVal token As String, ByVal isAM As Boolean, ByVal isPM As Boolean) As Long
    Dim hourPart As Long

    If Len(token) = 0 Then
        Err.Raise vbObjectError + 4014, "ParseHourToken", "Enter an hour value."
    End If

    If Not IsNumeric(token) Then
        Err.Raise vbObjectError + 4015, "ParseHourToken", _
            "'" & token & "' is not a valid hour."
    End If

    hourPart = CLng(token)
    ParseHourToken = NormalizeHour(hourPart, isAM, isPM)
End Function

'------------------------------------------------------------------------------
Private Function ParseMinuteToken(ByVal token As String, ByVal isAM As Boolean, ByVal isPM As Boolean) As Long
    Dim cleaned As String
    Dim i As Long
    Dim ch As String
    Dim minutePart As Long

    cleaned = ""
    For i = 1 To Len(token)
        ch = Mid$(token, i, 1)
        If ch >= "0" And ch <= "9" Then
            cleaned = cleaned & ch
        ElseIf ch = " " Then
            ' allow trailing AM/PM after minutes
        ElseIf ch = "a" Or ch = "A" Or ch = "p" Or ch = "P" Then
            Exit For
        Else
            Err.Raise vbObjectError + 4016, "ParseMinuteToken", _
                "'" & token & "' is not a valid minute value."
        End If
    Next i

    If Len(cleaned) = 0 Then
        Err.Raise vbObjectError + 4016, "ParseMinuteToken", _
            "'" & token & "' is not a valid minute value."
    End If

    minutePart = CLng(cleaned)
    ParseMinuteToken = minutePart
End Function

'------------------------------------------------------------------------------
Private Function NormalizeHour(ByVal hourPart As Long, ByVal isAM As Boolean, ByVal isPM As Boolean) As Long
    If isPM And isAM Then
        Err.Raise vbObjectError + 4017, "NormalizeHour", "Time has conflicting AM/PM markers."
    End If

    If isAM Or isPM Then
        If hourPart < 1 Or hourPart > 12 Then
            Err.Raise vbObjectError + 4018, "NormalizeHour", _
                "Hour must be between 1 and 12 when using AM or PM."
        End If
        If isPM And hourPart < 12 Then hourPart = hourPart + 12
        If isAM And hourPart = 12 Then hourPart = 0
    ElseIf hourPart > 23 Then
        Err.Raise vbObjectError + 4019, "NormalizeHour", _
            "Hour must be between 0 and 23."
    End If

    NormalizeHour = hourPart
End Function

'------------------------------------------------------------------------------
Private Function HasTimeSuffix(ByVal textValue As String, ByVal suffixLetter As String) As Boolean
    Dim lowered As String

    lowered = LCase$(Trim$(textValue))

    If suffixLetter = "p" Then
        If Right$(lowered, 2) = "pm" Then
            HasTimeSuffix = True
        ElseIf Right$(lowered, 1) = "p" And Right$(lowered, 2) <> "am" Then
            HasTimeSuffix = True
        End If
    ElseIf suffixLetter = "a" Then
        If Right$(lowered, 2) = "am" Then
            HasTimeSuffix = True
        ElseIf Right$(lowered, 1) = "a" Then
            HasTimeSuffix = True
        End If
    End If
End Function

'------------------------------------------------------------------------------
Private Function ExtractTimeDigits(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next i

    ExtractTimeDigits = digits
End Function

'------------------------------------------------------------------------------
Public Function ParseDateInput(ByVal textValue As String) As Date
    Dim trimmed As String

    trimmed = Trim$(textValue)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 4008, "ParseDateInput", "Enter a date."
    End If

    If Not IsDate(trimmed) Then
        Err.Raise vbObjectError + 4009, "ParseDateInput", _
            "'" & trimmed & "' is not a valid date."
    End If

    ParseDateInput = CDate(Int(CDate(trimmed)))
End Function

'------------------------------------------------------------------------------
Private Function FindOccupiedCells( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal col As Long) As String

    Dim values As Variant
    Dim i As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim found As Boolean

    values = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).Value
    found = False

    ' Single-cell ranges return a scalar, not an array.
    If startRow = endRow Then
        If Len(Trim$(CStr(values))) > 0 Then
            FindOccupiedCells = ws.Cells(startRow, col).Address(False, False)
        Else
            FindOccupiedCells = vbNullString
        End If
        Exit Function
    End If

    For i = 1 To UBound(values, 1)
        If Len(Trim$(CStr(values(i, 1)))) > 0 Then
            If Not found Then
                firstRow = startRow + i - 1
                found = True
            End If
            lastRow = startRow + i - 1
        End If
    Next i

    If Not found Then
        FindOccupiedCells = vbNullString
    ElseIf firstRow = lastRow Then
        FindOccupiedCells = ws.Cells(firstRow, col).Address(False, False)
    Else
        FindOccupiedCells = ws.Cells(firstRow, col).Address(False, False) & _
            ":" & ws.Cells(lastRow, col).Address(False, False)
    End If
End Function

'------------------------------------------------------------------------------
Private Sub WriteEntryValue( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal col As Long, _
    ByVal entryValue As String)

    Dim rowCount As Long
    Dim out() As Variant
    Dim i As Long

    rowCount = endRow - startRow + 1
    ReDim out(1 To rowCount, 1 To 1)

    For i = 1 To rowCount
        out(i, 1) = entryValue
    Next i

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Cells(startRow, col).Resize(rowCount, 1).Value = out

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
Private Sub ClearEntryRange( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal col As Long)

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).ClearContents

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
