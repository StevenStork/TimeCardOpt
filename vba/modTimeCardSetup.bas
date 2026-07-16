Attribute VB_Name = "modTimeCardSetup"
Option Explicit

'==============================================================================
' Time Card Setup
'
' Builds the Time Entry sheet from the date chosen on the Welcome page.
'
'   PrepareTimeEntry
'     1. Copies Welcome!F1 into Time Entry!B1
'     2. Fills Time Entry!A2 downward with every minute of the day (00:00-23:59)
'
' Both sheets are expected to already exist in the workbook template.
'==============================================================================

Private Const WELCOME_SHEET As String = "Welcome"
Private Const TIME_ENTRY_SHEET As String = "Time Entry"
Private Const WELCOME_DATE_CELL As String = "F1"
Private Const TIME_ENTRY_DATE_CELL As String = "B1"
Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440   ' 24 * 60

'------------------------------------------------------------------------------
' Public entry point — run this from a button or the Macros dialog.
'------------------------------------------------------------------------------
Public Sub PrepareTimeEntry()
    Dim welcomeWs As Worksheet
    Dim timeEntryWs As Worksheet
    Dim entryDate As Date

    On Error GoTo CleanFail
    OptimizeExcel True

    Set welcomeWs = ThisWorkbook.Worksheets(WELCOME_SHEET)
    Set timeEntryWs = ThisWorkbook.Worksheets(TIME_ENTRY_SHEET)

    entryDate = ReadWelcomeDate(welcomeWs)
    timeEntryWs.Range(TIME_ENTRY_DATE_CELL).Value = entryDate
    timeEntryWs.Range(TIME_ENTRY_DATE_CELL).NumberFormat = "mm/dd/yyyy"

    FillMinuteColumn timeEntryWs

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    MsgBox "PrepareTimeEntry failed: " & Err.Description, vbExclamation, "Time Card"
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' Read and validate the date from Welcome!F1.
'------------------------------------------------------------------------------
Private Function ReadWelcomeDate(ByVal welcomeWs As Worksheet) As Date
    Dim rawValue As Variant

    rawValue = welcomeWs.Range(WELCOME_DATE_CELL).Value

    If IsEmpty(rawValue) Or Not IsDate(rawValue) Then
        Err.Raise vbObjectError + 1001, "ReadWelcomeDate", _
            "Welcome!" & WELCOME_DATE_CELL & " must contain a valid date."
    End If

    ReadWelcomeDate = CDate(rawValue)
End Function

'------------------------------------------------------------------------------
' Write 00:00 through 23:59 into column A starting at row 2.
' Uses a single array write so 1,440 cells stay responsive.
'------------------------------------------------------------------------------
Private Sub FillMinuteColumn(ByVal timeEntryWs As Worksheet)
    Dim minutes() As Variant
    Dim i As Long
    Dim targetRange As Range

    ReDim minutes(1 To MINUTES_PER_DAY, 1 To 1)

    For i = 1 To MINUTES_PER_DAY
        ' Excel stores time as a fraction of a day; i-1 minutes from midnight.
        minutes(i, 1) = (i - 1) / MINUTES_PER_DAY
    Next i

    Set targetRange = timeEntryWs.Range( _
        timeEntryWs.Cells(MINUTE_START_ROW, 1), _
        timeEntryWs.Cells(MINUTE_START_ROW + MINUTES_PER_DAY - 1, 1))

    ' Clear any leftover values below the minute list, then write the day.
    timeEntryWs.Range(timeEntryWs.Cells(MINUTE_START_ROW, 1), _
        timeEntryWs.Cells(timeEntryWs.Rows.Count, 1)).ClearContents

    targetRange.Value = minutes
    targetRange.NumberFormat = "h:mm"
End Sub
