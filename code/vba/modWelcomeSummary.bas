Attribute VB_Name = "modWelcomeSummary"
Option Explicit

'==============================================================================
' Welcome Pay-Period Summary
'
' Rebuilds project and charge-code hour summaries on the Welcome sheet.
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
'   4. Rebuilds the Projects table (hours to nearest hundredth)
'   5. Rebuilds the Codes table two rows below (hours to nearest tenth)
'   6. Writes Codes total hours to N3
'   7. Sizes/stacks REButton, ETButton, ECButton, and ResetButton in column R
'   8. Activates the Welcome sheet
'
' Projects table:
'   B5          = "Projects"
'   C5:P5       = the 14 dates in the pay period that contains Welcome!J3
'   B6...       = project titles charged in that pay period
'
' Codes table (starts 2 rows below the Projects table):
'   B{n}        = "Codes"
'   C{n}:P{n}   = the same 14 pay-period dates
'   Rows below  = direct charge codes used this period, plus every direct
'                 code belonging to a project used this period
'   Hours       = direct minutes for that code, plus an even share of each
'                 parent project's minutes that day
'
' Pay-period boundaries are anchored to Welcome!F3 (same rule as Time Entry).
'==============================================================================

Private Const WELCOME_SHEET As String = "Welcome"
Private Const PERIOD_DATE_CELL As String = "J3"
Private Const ANCHOR_DATE_CELL As String = "F3"
Private Const PROJECTS_TITLE As String = "Projects"
Private Const CODES_TITLE As String = "Codes"
Private Const SUMMARY_START_ROW As Long = 5
Private Const SUMMARY_LABEL_COL As Long = 2          ' B
Private Const SUMMARY_FIRST_DATE_COL As Long = 3     ' C
Private Const SUMMARY_LAST_DATE_COL As Long = 16     ' P
Private Const PAY_PERIOD_DAYS As Long = 14
Private Const MINUTE_START_ROW As Long = 2
Private Const MINUTES_PER_DAY As Long = 1440
Private Const PRESERVE_ROWS As Long = 3
Private Const TEMP_SHEET_NAME As String = "__WelcomeRowsTemp"
Private Const TABLE_GAP_ROWS As Long = 2
Private Const CODES_TOTAL_CELL As String = "N3"
Private Const BUTTON_COLUMN As String = "R"
Private Const BUTTON_WIDTH As Double = 110
Private Const BUTTON_HEIGHT As Double = 32
Private Const BUTTON_GAP As Double = 14
Private Const BUTTON_START_ROW As Long = 5

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
    Dim directCodes As Variant
    Dim projectCount As Long
    Dim projectMinutes() As Long
    Dim directMinutes() As Long
    Dim lastProjectRow As Long
    Dim codesHeaderRow As Long
    Dim codesTotalHours As Double

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
    directCodes = LoadDirectCodes()
    projectCount = VariantLen(projectTitles)

    ResetWelcomeSheetPreservingTopRows welcomeWs

    If projectCount > 0 Then
        ReDim projectMinutes(1 To projectCount, 1 To PAY_PERIOD_DAYS)
    End If
    If VariantLen(directCodes) > 0 Then
        ReDim directMinutes(1 To VariantLen(directCodes), 1 To PAY_PERIOD_DAYS)
    End If

    ScanPayPeriodEntries timeEntryWs, periodDates, projectTitles, directCodes, _
        projectMinutes, directMinutes

    lastProjectRow = WriteProjectsTable(welcomeWs, periodDates, projectTitles, projectMinutes)
    FormatSummaryTable welcomeWs, SUMMARY_START_ROW, lastProjectRow, 2

    codesHeaderRow = lastProjectRow + TABLE_GAP_ROWS
    codesTotalHours = WriteCodesTable(welcomeWs, codesHeaderRow, periodDates, projectTitles, _
        directCodes, projectMinutes, directMinutes)

    With welcomeWs.Range(CODES_TOTAL_CELL)
        .Value = codesTotalHours
        .NumberFormat = "0.0"
    End With

    LayoutWelcomeButtons welcomeWs

CleanExit:
    OptimizeExcel False
    On Error Resume Next
    welcomeWs.Activate
    On Error GoTo 0
    Exit Sub

CleanFail:
    OptimizeExcel False
    On Error Resume Next
    If Not welcomeWs Is Nothing Then welcomeWs.Activate
    On Error GoTo 0
    MsgBox "UpdateWelcomeSummary failed: " & Err.Description, vbExclamation, "Welcome Summary"
End Sub

'------------------------------------------------------------------------------
Private Function WriteProjectsTable( _
    ByVal welcomeWs As Worksheet, _
    ByRef periodDates() As Date, _
    ByRef projectTitles As Variant, _
    ByRef projectMinutes() As Long) As Long

    Dim dayIndex As Long
    Dim projectIndex As Long
    Dim totalMinutes As Long
    Dim outRow As Long
    Dim hoursValue As Double
    Dim lastDataRow As Long

    welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_LABEL_COL).Value = PROJECTS_TITLE
    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(SUMMARY_START_ROW, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = _
            periodDates(dayIndex)
    Next dayIndex

    lastDataRow = SUMMARY_START_ROW
    outRow = SUMMARY_START_ROW + 1

    For projectIndex = 1 To VariantLen(projectTitles)
        totalMinutes = 0
        For dayIndex = 1 To PAY_PERIOD_DAYS
            totalMinutes = totalMinutes + projectMinutes(projectIndex, dayIndex)
        Next dayIndex

        If totalMinutes > 0 Then
            welcomeWs.Cells(outRow, SUMMARY_LABEL_COL).Value = _
                CStr(VariantItem(projectTitles, projectIndex))

            For dayIndex = 1 To PAY_PERIOD_DAYS
                hoursValue = Application.WorksheetFunction.Round( _
                    projectMinutes(projectIndex, dayIndex) / 60#, 2)
                welcomeWs.Cells(outRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = hoursValue
            Next dayIndex

            lastDataRow = outRow
            outRow = outRow + 1
        End If
    Next projectIndex

    WriteProjectsTable = lastDataRow
End Function

'------------------------------------------------------------------------------
' Returns the sum of hours written into the Codes table (nearest tenth).
'------------------------------------------------------------------------------
Private Function WriteCodesTable( _
    ByVal welcomeWs As Worksheet, _
    ByVal headerRow As Long, _
    ByRef periodDates() As Date, _
    ByRef projectTitles As Variant, _
    ByRef directCodes As Variant, _
    ByRef projectMinutes() As Long, _
    ByRef directMinutes() As Long) As Double

    Dim dayIndex As Long
    Dim codeList() As String
    Dim codeCount As Long
    Dim codeMinutes() As Double
    Dim codeIndex As Long
    Dim totalMinutes As Double
    Dim outRow As Long
    Dim lastDataRow As Long
    Dim hoursValue As Double
    Dim tableTotalHours As Double

    BuildCodesListAndMinutes projectTitles, directCodes, projectMinutes, directMinutes, _
        codeList, codeCount, codeMinutes

    welcomeWs.Cells(headerRow, SUMMARY_LABEL_COL).Value = CODES_TITLE
    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(headerRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = _
            periodDates(dayIndex)
    Next dayIndex

    lastDataRow = headerRow
    outRow = headerRow + 1
    tableTotalHours = 0

    For codeIndex = 1 To codeCount
        totalMinutes = 0
        For dayIndex = 1 To PAY_PERIOD_DAYS
            totalMinutes = totalMinutes + codeMinutes(codeIndex, dayIndex)
        Next dayIndex

        If totalMinutes > 0.0000001 Then
            welcomeWs.Cells(outRow, SUMMARY_LABEL_COL).Value = codeList(codeIndex)

            For dayIndex = 1 To PAY_PERIOD_DAYS
                hoursValue = Application.WorksheetFunction.Round( _
                    codeMinutes(codeIndex, dayIndex) / 60#, 1)
                welcomeWs.Cells(outRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).Value = hoursValue
                tableTotalHours = tableTotalHours + hoursValue
            Next dayIndex

            lastDataRow = outRow
            outRow = outRow + 1
        End If
    Next codeIndex

    FormatSummaryTable welcomeWs, headerRow, lastDataRow, 1
    WriteCodesTable = Application.WorksheetFunction.Round(tableTotalHours, 1)
End Function

'------------------------------------------------------------------------------
' Make Welcome action buttons the same size and stack them in col R.
' Matches buttons by shape name (REButton, etc.) or, for Reset, by caption
' "Reset Template" when the shape was not renamed ResetButton.
'------------------------------------------------------------------------------
Public Sub LayoutWelcomePageButtons()
    LayoutWelcomeButtons WelcomeSheet()
End Sub

Private Sub LayoutWelcomeButtons(ByVal welcomeWs As Worksheet)
    Dim buttonNames As Variant
    Dim i As Long
    Dim leftPos As Double
    Dim topPos As Double
    Dim btnTop As Double

    buttonNames = Array("REButton", "ETButton", "ECButton", "ResetButton")
    leftPos = welcomeWs.Columns(BUTTON_COLUMN).Left
    topPos = welcomeWs.Rows(BUTTON_START_ROW).Top

    For i = LBound(buttonNames) To UBound(buttonNames)
        btnTop = topPos + ((i - LBound(buttonNames)) * (BUTTON_HEIGHT + BUTTON_GAP))
        PositionWelcomeButton welcomeWs, CStr(buttonNames(i)), leftPos, btnTop, _
            BUTTON_WIDTH, BUTTON_HEIGHT
    Next i
End Sub

'------------------------------------------------------------------------------
Private Sub PositionWelcomeButton( _
    ByVal welcomeWs As Worksheet, _
    ByVal buttonKey As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthPos As Double, _
    ByVal heightPos As Double)

    Dim shp As Shape
    Dim oleObj As OLEObject

    Set shp = FindWelcomeShape(welcomeWs, buttonKey)
    If Not shp Is Nothing Then
        With shp
            .LockAspectRatio = msoFalse
            .Placement = xlFreeFloating
            .Left = leftPos
            .Top = topPos
            .Width = widthPos
            .Height = heightPos
        End With
        AssignWelcomeButtonMacro buttonKey, shp
        Exit Sub
    End If

    Set oleObj = FindWelcomeOleObject(welcomeWs, buttonKey)
    If Not oleObj Is Nothing Then
        With oleObj
            .Left = leftPos
            .Top = topPos
            .Width = widthPos
            .Height = heightPos
        End With
    End If
End Sub

'------------------------------------------------------------------------------
Private Function FindWelcomeShape(ByVal welcomeWs As Worksheet, ByVal buttonKey As String) As Shape
    Dim shp As Shape

    On Error Resume Next
    Set FindWelcomeShape = welcomeWs.Shapes(buttonKey)
    On Error GoTo 0
    If Not FindWelcomeShape Is Nothing Then Exit Function

    For Each shp In welcomeWs.Shapes
        If ShapeMatchesButtonKey(shp, buttonKey) Then
            Set FindWelcomeShape = shp
            Exit Function
        End If
    Next shp
End Function

'------------------------------------------------------------------------------
Private Function FindWelcomeOleObject(ByVal welcomeWs As Worksheet, ByVal buttonKey As String) As OLEObject
    Dim oleObj As OLEObject

    On Error Resume Next
    Set FindWelcomeOleObject = welcomeWs.OLEObjects(buttonKey)
    On Error GoTo 0
    If Not FindWelcomeOleObject Is Nothing Then Exit Function

    For Each oleObj In welcomeWs.OLEObjects
        If OleMatchesButtonKey(oleObj, buttonKey) Then
            Set FindWelcomeOleObject = oleObj
            Exit Function
        End If
    Next oleObj
End Function

'------------------------------------------------------------------------------
Private Function ShapeMatchesButtonKey(ByVal shp As Shape, ByVal buttonKey As String) As Boolean
    If StrComp(shp.Name, buttonKey, vbTextCompare) = 0 Then
        ShapeMatchesButtonKey = True
        Exit Function
    End If

    If shp.Type <> msoFormControl Then Exit Function
    If shp.FormControlType <> xlButtonControl Then Exit Function

    Select Case buttonKey
        Case "ResetButton"
            ShapeMatchesButtonKey = CaptionMatchesReset(GetShapeCaption(shp))
    End Select
End Function

'------------------------------------------------------------------------------
Private Function OleMatchesButtonKey(ByVal oleObj As OLEObject, ByVal buttonKey As String) As Boolean
    Dim caption As String

    If StrComp(oleObj.Name, buttonKey, vbTextCompare) = 0 Then
        OleMatchesButtonKey = True
        Exit Function
    End If

    On Error Resume Next
    caption = CStr(oleObj.Object.Caption)
    On Error GoTo 0

    Select Case buttonKey
        Case "ResetButton"
            OleMatchesButtonKey = CaptionMatchesReset(caption)
    End Select
End Function

'------------------------------------------------------------------------------
Private Function GetShapeCaption(ByVal shp As Shape) As String
    On Error Resume Next
    GetShapeCaption = shp.ControlFormat.Caption
    If Len(GetShapeCaption) = 0 Then
        GetShapeCaption = shp.TextFrame.Characters.Text
    End If
    On Error GoTo 0
    GetShapeCaption = Trim$(GetShapeCaption)
End Function

'------------------------------------------------------------------------------
Private Function CaptionMatchesReset(ByVal caption As String) As Boolean
    CaptionMatchesReset = (StrComp(Trim$(caption), "Reset Template", vbTextCompare) = 0) _
        Or (InStr(1, caption, "Reset Template", vbTextCompare) > 0)
End Function

'------------------------------------------------------------------------------
Private Sub AssignWelcomeButtonMacro(ByVal buttonKey As String, ByVal btn As Object)
    If StrComp(buttonKey, "ResetButton", vbTextCompare) <> 0 Then Exit Sub
    On Error Resume Next
    btn.OnAction = "ResetTemplate"
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Codes = directs used this period + every direct that belongs to a used project.
' Minutes = direct minutes + even share of each parent project's minutes/day.
'------------------------------------------------------------------------------
Private Sub BuildCodesListAndMinutes( _
    ByRef projectTitles As Variant, _
    ByRef directCodes As Variant, _
    ByRef projectMinutes() As Long, _
    ByRef directMinutes() As Long, _
    ByRef codeList() As String, _
    ByRef codeCount As Long, _
    ByRef codeMinutes() As Double)

    Dim projectIndex As Long
    Dim dayIndex As Long
    Dim codeIndex As Long
    Dim assoc As Variant
    Dim assocCount As Long
    Dim a As Long
    Dim share As Double
    Dim usedProject() As Boolean
    Dim includeCode() As Boolean
    Dim directCount As Long
    Dim codeName As String
    Dim mappedIndex As Long

    directCount = VariantLen(directCodes)
    codeCount = 0

    If VariantLen(projectTitles) > 0 Then
        ReDim usedProject(1 To VariantLen(projectTitles))
        For projectIndex = 1 To VariantLen(projectTitles)
            For dayIndex = 1 To PAY_PERIOD_DAYS
                If projectMinutes(projectIndex, dayIndex) > 0 Then
                    usedProject(projectIndex) = True
                    Exit For
                End If
            Next dayIndex
        Next projectIndex
    End If

    If directCount > 0 Then
        ReDim includeCode(1 To directCount)
        For codeIndex = 1 To directCount
            For dayIndex = 1 To PAY_PERIOD_DAYS
                If directMinutes(codeIndex, dayIndex) > 0 Then
                    includeCode(codeIndex) = True
                    Exit For
                End If
            Next dayIndex
        Next codeIndex
    End If

    For projectIndex = 1 To VariantLen(projectTitles)
        If usedProject(projectIndex) Then
            assoc = LoadProjectDirectCodes(CStr(VariantItem(projectTitles, projectIndex)))
            For a = 1 To VariantLen(assoc)
                codeName = CStr(VariantItem(assoc, a))
                mappedIndex = FindCodeIndex(directCodes, codeName)
                If mappedIndex > 0 Then
                    includeCode(mappedIndex) = True
                End If
            Next a
        End If
    Next projectIndex

    ' Prefer the master direct-code order, then any project-only codes not in it.
    For codeIndex = 1 To directCount
        If includeCode(codeIndex) Then
            AppendUniqueCode codeList, codeCount, CStr(VariantItem(directCodes, codeIndex))
        End If
    Next codeIndex

    For projectIndex = 1 To VariantLen(projectTitles)
        If usedProject(projectIndex) Then
            assoc = LoadProjectDirectCodes(CStr(VariantItem(projectTitles, projectIndex)))
            For a = 1 To VariantLen(assoc)
                AppendUniqueCode codeList, codeCount, CStr(VariantItem(assoc, a))
            Next a
        End If
    Next projectIndex

    If codeCount = 0 Then
        ReDim codeList(1 To 1)
        ReDim codeMinutes(1 To 1, 1 To PAY_PERIOD_DAYS)
        codeCount = 0
        Exit Sub
    End If

    ReDim codeMinutes(1 To codeCount, 1 To PAY_PERIOD_DAYS)

    ' Direct minutes.
    For codeIndex = 1 To directCount
        If includeCode(codeIndex) Then
            mappedIndex = FindCodeIndexFromList(codeList, codeCount, _
                CStr(VariantItem(directCodes, codeIndex)))
            If mappedIndex > 0 Then
                For dayIndex = 1 To PAY_PERIOD_DAYS
                    codeMinutes(mappedIndex, dayIndex) = _
                        codeMinutes(mappedIndex, dayIndex) + directMinutes(codeIndex, dayIndex)
                Next dayIndex
            End If
        End If
    Next codeIndex

    ' Evenly distribute each used project's minutes across its direct codes.
    For projectIndex = 1 To VariantLen(projectTitles)
        If Not usedProject(projectIndex) Then GoTo NextProject

        assoc = LoadProjectDirectCodes(CStr(VariantItem(projectTitles, projectIndex)))
        assocCount = VariantLen(assoc)
        If assocCount = 0 Then GoTo NextProject

        For dayIndex = 1 To PAY_PERIOD_DAYS
            If projectMinutes(projectIndex, dayIndex) > 0 Then
                share = projectMinutes(projectIndex, dayIndex) / assocCount
                For a = 1 To assocCount
                    mappedIndex = FindCodeIndexFromList(codeList, codeCount, _
                        CStr(VariantItem(assoc, a)))
                    If mappedIndex > 0 Then
                        codeMinutes(mappedIndex, dayIndex) = _
                            codeMinutes(mappedIndex, dayIndex) + share
                    End If
                Next a
            End If
        Next dayIndex
NextProject:
    Next projectIndex
End Sub

'------------------------------------------------------------------------------
Private Sub AppendUniqueCode( _
    ByRef codeList() As String, _
    ByRef codeCount As Long, _
    ByVal codeName As String)

    Dim trimmed As String

    trimmed = Trim$(codeName)
    If Len(trimmed) = 0 Then Exit Sub
    If FindCodeIndexFromList(codeList, codeCount, trimmed) > 0 Then Exit Sub

    codeCount = codeCount + 1
    If codeCount = 1 Then
        ReDim codeList(1 To 1)
    Else
        ReDim Preserve codeList(1 To codeCount)
    End If
    codeList(codeCount) = trimmed
End Sub

'------------------------------------------------------------------------------
Private Function FindCodeIndexFromList( _
    ByRef codeList() As String, _
    ByVal codeCount As Long, _
    ByVal codeName As String) As Long

    Dim i As Long

    For i = 1 To codeCount
        If StrComp(codeList(i), codeName, vbTextCompare) = 0 Then
            FindCodeIndexFromList = i
            Exit Function
        End If
    Next i

    FindCodeIndexFromList = 0
End Function

'------------------------------------------------------------------------------
' One pass over the pay-period Time Entry columns.
'------------------------------------------------------------------------------
Private Sub ScanPayPeriodEntries( _
    ByVal timeEntryWs As Worksheet, _
    ByRef periodDates() As Date, _
    ByRef projectTitles As Variant, _
    ByRef directCodes As Variant, _
    ByRef projectMinutes() As Long, _
    ByRef directMinutes() As Long)

    Dim dayIndex As Long
    Dim col As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim cellText As String
    Dim projectIndex As Long
    Dim codeIndex As Long

    For dayIndex = 1 To PAY_PERIOD_DAYS
        col = FindDateColumn(timeEntryWs, periodDates(dayIndex))
        If col = 0 Then GoTo NextDay

        values = timeEntryWs.Range( _
            timeEntryWs.Cells(MINUTE_START_ROW, col), _
            timeEntryWs.Cells(MINUTE_START_ROW + MINUTES_PER_DAY - 1, col)).Value

        For rowIndex = 1 To MINUTES_PER_DAY
            cellText = Trim$(CStr(values(rowIndex, 1)))
            If Len(cellText) = 0 Then GoTo NextCell

            projectIndex = FindNameIndex(projectTitles, cellText)
            If projectIndex > 0 Then
                projectMinutes(projectIndex, dayIndex) = projectMinutes(projectIndex, dayIndex) + 1
            Else
                codeIndex = FindNameIndex(directCodes, cellText)
                If codeIndex > 0 Then
                    directMinutes(codeIndex, dayIndex) = directMinutes(codeIndex, dayIndex) + 1
                End If
            End If
NextCell:
        Next rowIndex
NextDay:
    Next dayIndex
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
' hourDecimals: 2 = nearest hundredth (Projects), 1 = nearest tenth (Codes)
'------------------------------------------------------------------------------
Private Sub FormatSummaryTable( _
    ByVal welcomeWs As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastDataRow As Long, _
    ByVal hourDecimals As Long)

    Dim headerRange As Range
    Dim tableRange As Range
    Dim dataRange As Range
    Dim hoursRange As Range
    Dim labelRange As Range
    Dim r As Long
    Dim dayIndex As Long
    Dim numberFormat As String
    Dim firstDataRow As Long

    firstDataRow = headerRow + 1
    If hourDecimals <= 1 Then
        numberFormat = "0.0"
    Else
        numberFormat = "0.00"
    End If

    Set headerRange = welcomeWs.Range( _
        welcomeWs.Cells(headerRow, SUMMARY_LABEL_COL), _
        welcomeWs.Cells(headerRow, SUMMARY_LAST_DATE_COL))

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

    welcomeWs.Cells(headerRow, SUMMARY_LABEL_COL).HorizontalAlignment = xlLeft

    For dayIndex = 1 To PAY_PERIOD_DAYS
        welcomeWs.Cells(headerRow, SUMMARY_FIRST_DATE_COL + dayIndex - 1).NumberFormat = "mmm d"
    Next dayIndex

    welcomeWs.Rows(headerRow).RowHeight = 30

    If lastDataRow >= firstDataRow Then
        Set dataRange = welcomeWs.Range( _
            welcomeWs.Cells(firstDataRow, SUMMARY_LABEL_COL), _
            welcomeWs.Cells(lastDataRow, SUMMARY_LAST_DATE_COL))
        Set hoursRange = welcomeWs.Range( _
            welcomeWs.Cells(firstDataRow, SUMMARY_FIRST_DATE_COL), _
            welcomeWs.Cells(lastDataRow, SUMMARY_LAST_DATE_COL))
        Set labelRange = welcomeWs.Range( _
            welcomeWs.Cells(firstDataRow, SUMMARY_LABEL_COL), _
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
            .NumberFormat = numberFormat
        End With

        For r = firstDataRow To lastDataRow
            If ((r - firstDataRow) Mod 2) = 1 Then
                welcomeWs.Range( _
                    welcomeWs.Cells(r, SUMMARY_LABEL_COL), _
                    welcomeWs.Cells(r, SUMMARY_LAST_DATE_COL)).Interior.Color = ALT_ROW_FILL_RGB
            End If
            welcomeWs.Rows(r).RowHeight = 20
        Next r
    End If

    Set tableRange = welcomeWs.Range( _
        welcomeWs.Cells(headerRow, SUMMARY_LABEL_COL), _
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
Private Function FindNameIndex(ByRef names As Variant, ByVal cellText As String) As Long
    Dim i As Long

    For i = 1 To VariantLen(names)
        If StrComp(CStr(VariantItem(names, i)), cellText, vbTextCompare) = 0 Then
            FindNameIndex = i
            Exit Function
        End If
    Next i

    FindNameIndex = 0
End Function

'------------------------------------------------------------------------------
Private Function FindCodeIndex(ByRef directCodes As Variant, ByVal codeName As String) As Long
    FindCodeIndex = FindNameIndex(directCodes, codeName)
End Function
