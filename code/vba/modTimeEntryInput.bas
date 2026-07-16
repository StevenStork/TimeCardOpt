Attribute VB_Name = "modTimeEntryInput"
Option Explicit

'==============================================================================
' Time Entry Input
'
' Writes a charge code into the Time Entry minute grid for a date/time range.
'
'   ShowTimeEntryForm   launch the hours-entry UserForm
'   EnterTimeRange      validate and write the selected charge code
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
' Write chargeCode into every minute cell for entryDate in [startTime, endTime).
' Raises a clear error if any target cell already has a value.
'------------------------------------------------------------------------------
Public Sub EnterTimeRange( _
    ByVal entryDate As Date, _
    ByVal chargeCode As String, _
    ByVal startTime As Date, _
    ByVal endTime As Date)

    Dim ws As Worksheet
    Dim col As Long
    Dim startMinute As Long
    Dim endMinute As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim code As String
    Dim overlapAddress As String

    code = Trim$(chargeCode)
    If Len(code) = 0 Then
        Err.Raise vbObjectError + 4001, "EnterTimeRange", "Select a charge code."
    End If

    startMinute = MinuteOfDay(startTime)
    endMinute = MinuteOfDay(endTime)

    If endMinute <= startMinute Then
        Err.Raise vbObjectError + 4002, "EnterTimeRange", _
            "End time must be after start time on the same day."
    End If

    Set ws = TimeEntrySheet()
    col = FindDateColumn(ws, entryDate)
    If col = 0 Then
        Err.Raise vbObjectError + 4003, "EnterTimeRange", _
            "Date " & Format$(entryDate, "mm/dd/yyyy") & _
            " was not found on the Time Entry sheet. Run PrepareTimeEntry first."
    End If

    ' Rows cover minutes startMinute .. endMinute-1 inclusive.
    startRow = MINUTE_START_ROW + startMinute
    endRow = MINUTE_START_ROW + endMinute - 1

    overlapAddress = FindOccupiedCells(ws, startRow, endRow, col)
    If Len(overlapAddress) > 0 Then
        Err.Raise vbObjectError + 4004, "EnterTimeRange", _
            "Cannot overwrite existing entries in " & overlapAddress & "."
    End If

    WriteChargeCode ws, startRow, endRow, col, code
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
Public Function ParseTimeInput(ByVal textValue As String) As Date
    Dim trimmed As String

    trimmed = Trim$(textValue)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 4006, "ParseTimeInput", "Enter a start and end time."
    End If

    If Not IsDate(trimmed) Then
        Err.Raise vbObjectError + 4007, "ParseTimeInput", _
            "'" & trimmed & "' is not a valid time."
    End If

    ParseTimeInput = CDate(trimmed)
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
Private Sub WriteChargeCode( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal col As Long, _
    ByVal chargeCode As String)

    Dim rowCount As Long
    Dim out() As Variant
    Dim i As Long

    rowCount = endRow - startRow + 1
    ReDim out(1 To rowCount, 1 To 1)

    For i = 1 To rowCount
        out(i, 1) = chargeCode
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
