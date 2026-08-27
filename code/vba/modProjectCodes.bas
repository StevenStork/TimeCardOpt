Attribute VB_Name = "modProjectCodes"
Option Explicit

'==============================================================================
' Project Codes
'
' Sheet layout on "Project Codes" (sheet is expected to already exist):
'   Column B      - direct charge codes (B1 header, codes from B2 down)
'   Column C      - spacer
'   Columns D...  - one project per column
'                   row 1 = project title
'                   row 2 down = associated direct charge codes
'
'   ShowProjectCodesForm   launch the editor UserForm
'
' Setup in the workbook:
'   1. Import modProjectCodes.bas, clsHookCommand.cls, clsHookListBox.cls
'   2. Insert an empty UserForm and rename it to frmProjectCodes
'   3. Paste the code from frmProjectCodes.frm (from Option Explicit down)
'   4. Confirm a sheet named "Project Codes" already exists
'   5. Run ShowProjectCodesForm (or hook it to a Welcome-page button)
'==============================================================================

Public Const PROJECT_CODES_SHEET As String = "Project Codes"
Public Const DIRECT_CODE_COL As Long = 2          ' B
Public Const DIRECT_CODE_HEADER_ROW As Long = 1
Public Const DIRECT_CODE_START_ROW As Long = 2
Public Const PROJECT_START_COL As Long = 4        ' D
Public Const PROJECT_TITLE_ROW As Long = 1
Public Const PROJECT_CODE_START_ROW As Long = 2

'------------------------------------------------------------------------------
Public Sub ShowProjectCodesForm()
    EnsureProjectCodesLayout
    frmProjectCodes.Show vbModal
End Sub

'------------------------------------------------------------------------------
' Make sure headers exist; does not create the worksheet.
'------------------------------------------------------------------------------
Public Sub EnsureProjectCodesLayout()
    Dim ws As Worksheet

    Set ws = ProjectCodesSheet()
    If Len(Trim$(CStr(ws.Cells(DIRECT_CODE_HEADER_ROW, DIRECT_CODE_COL).Value))) = 0 Then
        ws.Cells(DIRECT_CODE_HEADER_ROW, DIRECT_CODE_COL).Value = "Direct Charge Codes"
    End If
End Sub

'------------------------------------------------------------------------------
Public Function ProjectCodesSheet() As Worksheet
    On Error Resume Next
    Set ProjectCodesSheet = ThisWorkbook.Worksheets(PROJECT_CODES_SHEET)
    On Error GoTo 0

    If ProjectCodesSheet Is Nothing Then
        Err.Raise vbObjectError + 3001, "ProjectCodesSheet", _
            "Worksheet '" & PROJECT_CODES_SHEET & "' was not found."
    End If
End Function

'------------------------------------------------------------------------------
Public Function LoadDirectCodes() As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim count As Long
    Dim r As Long
    Dim values() As String
    Dim cellValue As String

    Set ws = ProjectCodesSheet()
    lastRow = ws.Cells(ws.Rows.Count, DIRECT_CODE_COL).End(xlUp).Row

    If lastRow < DIRECT_CODE_START_ROW Then
        LoadDirectCodes = Array()
        Exit Function
    End If

    ReDim values(1 To lastRow - DIRECT_CODE_START_ROW + 1)
    count = 0

    For r = DIRECT_CODE_START_ROW To lastRow
        cellValue = Trim$(CStr(ws.Cells(r, DIRECT_CODE_COL).Value))
        If Len(cellValue) > 0 Then
            count = count + 1
            values(count) = cellValue
        End If
    Next r

    If count = 0 Then
        LoadDirectCodes = Array()
    Else
        ReDim Preserve values(1 To count)
        LoadDirectCodes = values
    End If
End Function

'------------------------------------------------------------------------------
Public Sub SaveDirectCodes(ByRef codes As Variant)
    Dim ws As Worksheet
    Dim i As Long
    Dim n As Long
    Dim out() As Variant

    Set ws = ProjectCodesSheet()
    n = VariantLen(codes)

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Range(ws.Cells(DIRECT_CODE_START_ROW, DIRECT_CODE_COL), _
        ws.Cells(ws.Rows.Count, DIRECT_CODE_COL)).ClearContents

    If n > 0 Then
        ReDim out(1 To n, 1 To 1)
        For i = 1 To n
            out(i, 1) = CStr(VariantItem(codes, i))
        Next i
        ws.Cells(DIRECT_CODE_START_ROW, DIRECT_CODE_COL).Resize(n, 1).Value = out
    End If

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
' Remove every direct charge code and project from the Project Codes sheet.
' Keeps the column B header; clears project columns from D onward.
'------------------------------------------------------------------------------
Public Sub ClearAllProjectCodes()
    Dim ws As Worksheet
    Dim lastCol As Long

    Set ws = ProjectCodesSheet()

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Range(ws.Cells(DIRECT_CODE_START_ROW, DIRECT_CODE_COL), _
        ws.Cells(ws.Rows.Count, DIRECT_CODE_COL)).ClearContents

    lastCol = ws.Cells(PROJECT_TITLE_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol >= PROJECT_START_COL Then
        ws.Range(ws.Cells(1, PROJECT_START_COL), _
            ws.Cells(ws.Rows.Count, lastCol)).ClearContents
    End If

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
Public Function LoadProjectTitles() As Variant
    Dim ws As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim count As Long
    Dim titles() As String
    Dim title As String

    Set ws = ProjectCodesSheet()
    lastCol = ws.Cells(PROJECT_TITLE_ROW, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < PROJECT_START_COL Then
        LoadProjectTitles = Array()
        Exit Function
    End If

    ReDim titles(1 To lastCol - PROJECT_START_COL + 1)
    count = 0

    For col = PROJECT_START_COL To lastCol
        title = Trim$(CStr(ws.Cells(PROJECT_TITLE_ROW, col).Value))
        If Len(title) > 0 Then
            count = count + 1
            titles(count) = title
        End If
    Next col

    If count = 0 Then
        LoadProjectTitles = Array()
    Else
        ReDim Preserve titles(1 To count)
        LoadProjectTitles = titles
    End If
End Function

'------------------------------------------------------------------------------
' Returns 1-based String array of direct codes under the project title.
'------------------------------------------------------------------------------
Public Function LoadProjectDirectCodes(ByVal projectTitle As String) As Variant
    Dim ws As Worksheet
    Dim col As Long
    Dim lastRow As Long
    Dim r As Long
    Dim count As Long
    Dim values() As String
    Dim cellValue As String

    Set ws = ProjectCodesSheet()
    col = FindProjectColumn(projectTitle)

    If col = 0 Then
        LoadProjectDirectCodes = Array()
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < PROJECT_CODE_START_ROW Then
        LoadProjectDirectCodes = Array()
        Exit Function
    End If

    ReDim values(1 To lastRow - PROJECT_CODE_START_ROW + 1)
    count = 0

    For r = PROJECT_CODE_START_ROW To lastRow
        cellValue = Trim$(CStr(ws.Cells(r, col).Value))
        If Len(cellValue) > 0 Then
            count = count + 1
            values(count) = cellValue
        End If
    Next r

    If count = 0 Then
        LoadProjectDirectCodes = Array()
    Else
        ReDim Preserve values(1 To count)
        LoadProjectDirectCodes = values
    End If
End Function

'------------------------------------------------------------------------------
Public Function AddDirectCode(ByVal code As String) As Boolean
    Dim codes As Variant
    Dim trimmed As String
    Dim i As Long
    Dim n As Long
    Dim updated() As String

    trimmed = Trim$(code)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 3002, "AddDirectCode", "Enter a direct charge code."
    End If

    codes = LoadDirectCodes()
    n = VariantLen(codes)

    For i = 1 To n
        If StrComp(CStr(VariantItem(codes, i)), trimmed, vbTextCompare) = 0 Then
            AddDirectCode = False
            Exit Function
        End If
    Next i

    ReDim updated(1 To n + 1)
    For i = 1 To n
        updated(i) = CStr(VariantItem(codes, i))
    Next i
    updated(n + 1) = trimmed

    SaveDirectCodes updated
    AddDirectCode = True
End Function

'------------------------------------------------------------------------------
Public Sub RemoveDirectCode(ByVal code As String)
    Dim codes As Variant
    Dim i As Long
    Dim n As Long
    Dim kept As Long
    Dim updated() As String
    Dim trimmed As String

    trimmed = Trim$(code)
    codes = LoadDirectCodes()
    n = VariantLen(codes)
    If n = 0 Then Exit Sub

    ReDim updated(1 To n)
    kept = 0

    For i = 1 To n
        If StrComp(CStr(VariantItem(codes, i)), trimmed, vbTextCompare) <> 0 Then
            kept = kept + 1
            updated(kept) = CStr(VariantItem(codes, i))
        End If
    Next i

    If kept = 0 Then
        SaveDirectCodes Array()
    Else
        ReDim Preserve updated(1 To kept)
        SaveDirectCodes updated
    End If

    RemoveDirectCodeFromAllProjects trimmed
End Sub

'------------------------------------------------------------------------------
Public Function AddProject(ByVal projectTitle As String) As Boolean
    Dim ws As Worksheet
    Dim trimmed As String
    Dim col As Long

    trimmed = Trim$(projectTitle)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 3003, "AddProject", "Enter a project title."
    End If

    If FindProjectColumn(trimmed) > 0 Then
        AddProject = False
        Exit Function
    End If

    Set ws = ProjectCodesSheet()
    col = NextProjectColumn(ws)
    ws.Cells(PROJECT_TITLE_ROW, col).Value = trimmed
    AddProject = True
End Function

'------------------------------------------------------------------------------
Public Sub SaveProjectAssociations(ByVal projectTitle As String, ByRef selectedCodes As Variant)
    Dim ws As Worksheet
    Dim col As Long
    Dim n As Long
    Dim i As Long
    Dim out() As Variant

    col = FindProjectColumn(projectTitle)
    If col = 0 Then
        Err.Raise vbObjectError + 3004, "SaveProjectAssociations", _
            "Project '" & projectTitle & "' was not found."
    End If

    Set ws = ProjectCodesSheet()
    n = VariantLen(selectedCodes)

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Range(ws.Cells(PROJECT_CODE_START_ROW, col), _
        ws.Cells(ws.Rows.Count, col)).ClearContents

    If n > 0 Then
        ReDim out(1 To n, 1 To 1)
        For i = 1 To n
            out(i, 1) = CStr(VariantItem(selectedCodes, i))
        Next i
        ws.Cells(PROJECT_CODE_START_ROW, col).Resize(n, 1).Value = out
    End If

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
Public Sub DeleteProject(ByVal projectTitle As String)
    Dim ws As Worksheet
    Dim col As Long

    col = FindProjectColumn(projectTitle)
    If col = 0 Then Exit Sub

    Set ws = ProjectCodesSheet()

    OptimizeExcel True
    On Error GoTo CleanFail

    ' Deleting the column shifts later projects left so titles stay contiguous from D.
    ws.Columns(col).Delete Shift:=xlToLeft

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
Public Function FindProjectColumn(ByVal projectTitle As String) As Long
    Dim ws As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim title As String
    Dim trimmed As String

    trimmed = Trim$(projectTitle)
    Set ws = ProjectCodesSheet()
    lastCol = ws.Cells(PROJECT_TITLE_ROW, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < PROJECT_START_COL Then
        FindProjectColumn = 0
        Exit Function
    End If

    For col = PROJECT_START_COL To lastCol
        title = Trim$(CStr(ws.Cells(PROJECT_TITLE_ROW, col).Value))
        If StrComp(title, trimmed, vbTextCompare) = 0 Then
            FindProjectColumn = col
            Exit Function
        End If
    Next col

    FindProjectColumn = 0
End Function

'------------------------------------------------------------------------------
Private Function NextProjectColumn(ByVal ws As Worksheet) As Long
    Dim lastCol As Long
    Dim col As Long

    lastCol = ws.Cells(PROJECT_TITLE_ROW, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < PROJECT_START_COL Then
        NextProjectColumn = PROJECT_START_COL
        Exit Function
    End If

    For col = PROJECT_START_COL To lastCol
        If Len(Trim$(CStr(ws.Cells(PROJECT_TITLE_ROW, col).Value))) = 0 Then
            NextProjectColumn = col
            Exit Function
        End If
    Next col

    NextProjectColumn = lastCol + 1
End Function

'------------------------------------------------------------------------------
' Move a direct charge code up or down in column B.
'------------------------------------------------------------------------------
Public Sub MoveDirectCode(ByVal code As String, ByVal moveUp As Boolean)
    Dim codes As Variant
    Dim updated() As String
    Dim n As Long
    Dim idx As Long
    Dim i As Long

    Dim trimmed As String
    trimmed = Trim$(code)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 3010, "MoveDirectCode", "Select a direct charge code to move."
    End If

    codes = LoadDirectCodes()
    n = VariantLen(codes)
    If n < 2 Then Exit Sub

    idx = 0
    For i = 1 To n
        If StrComp(CStr(VariantItem(codes, i)), trimmed, vbTextCompare) = 0 Then
            idx = i
            Exit For
        End If
    Next i

    If idx = 0 Then
        Err.Raise vbObjectError + 3011, "MoveDirectCode", _
            "Direct charge code '" & trimmed & "' was not found."
    End If

    If moveUp And idx = 1 Then Exit Sub
    If Not moveUp And idx = n Then Exit Sub

    ReDim updated(1 To n)
    For i = 1 To n
        updated(i) = CStr(VariantItem(codes, i))
    Next i

    If moveUp Then
        SwapStrings updated(idx), updated(idx - 1)
    Else
        SwapStrings updated(idx), updated(idx + 1)
    End If

    SaveDirectCodes updated
End Sub

'------------------------------------------------------------------------------
' Move a project column left (up) or right (down) among project columns.
'------------------------------------------------------------------------------
Public Sub MoveProject(ByVal projectTitle As String, ByVal moveUp As Boolean)
    Dim ws As Worksheet
    Dim col As Long
    Dim lastCol As Long
    Dim trimmed As String

    trimmed = Trim$(projectTitle)
    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 3012, "MoveProject", "Select a project to move."
    End If

    col = FindProjectColumn(trimmed)
    If col = 0 Then
        Err.Raise vbObjectError + 3013, "MoveProject", _
            "Project '" & trimmed & "' was not found."
    End If

    Set ws = ProjectCodesSheet()
    lastCol = ws.Cells(PROJECT_TITLE_ROW, ws.Columns.Count).End(xlToLeft).Column

    If moveUp Then
        If col <= PROJECT_START_COL Then Exit Sub
        SwapProjectColumns ws, col, col - 1
    Else
        If col >= lastCol Then Exit Sub
        SwapProjectColumns ws, col, col + 1
    End If
End Sub

'------------------------------------------------------------------------------
Private Sub SwapProjectColumns(ByVal ws As Worksheet, ByVal colA As Long, ByVal colB As Long)
    Dim lastRow As Long
    Dim valuesA As Variant
    Dim valuesB As Variant

    lastRow = ws.Rows.Count
    valuesA = ws.Range(ws.Cells(1, colA), ws.Cells(lastRow, colA)).Value
    valuesB = ws.Range(ws.Cells(1, colB), ws.Cells(lastRow, colB)).Value

    OptimizeExcel True
    On Error GoTo CleanFail

    ws.Range(ws.Cells(1, colA), ws.Cells(lastRow, colA)).Value = valuesB
    ws.Range(ws.Cells(1, colB), ws.Cells(lastRow, colB)).Value = valuesA

CleanExit:
    OptimizeExcel False
    Exit Sub

CleanFail:
    OptimizeExcel False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'------------------------------------------------------------------------------
Private Sub SwapStrings(ByRef firstValue As String, ByRef secondValue As String)
    Dim temp As String
    temp = firstValue
    firstValue = secondValue
    secondValue = temp
End Sub

'------------------------------------------------------------------------------
Private Sub RemoveDirectCodeFromAllProjects(ByVal code As String)
    Dim titles As Variant
    Dim i As Long
    Dim associated As Variant
    Dim j As Long
    Dim kept As Long
    Dim updated() As String
    Dim n As Long

    titles = LoadProjectTitles()
    For i = 1 To VariantLen(titles)
        associated = LoadProjectDirectCodes(CStr(VariantItem(titles, i)))
        n = VariantLen(associated)
        If n = 0 Then GoTo NextProject

        ReDim updated(1 To n)
        kept = 0
        For j = 1 To n
            If StrComp(CStr(VariantItem(associated, j)), code, vbTextCompare) <> 0 Then
                kept = kept + 1
                updated(kept) = CStr(VariantItem(associated, j))
            End If
        Next j

        If kept = 0 Then
            SaveProjectAssociations CStr(VariantItem(titles, i)), Array()
        ElseIf kept < n Then
            ReDim Preserve updated(1 To kept)
            SaveProjectAssociations CStr(VariantItem(titles, i)), updated
        End If
NextProject:
    Next i
End Sub

'------------------------------------------------------------------------------
' Length helper for 1-based arrays or empty Array() results.
'------------------------------------------------------------------------------
Public Function VariantLen(ByRef values As Variant) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    If IsEmpty(values) Then
        VariantLen = 0
        Exit Function
    End If

    On Error Resume Next
    lowerBound = LBound(values)
    upperBound = UBound(values)
    If Err.Number <> 0 Then
        Err.Clear
        VariantLen = 0
        Exit Function
    End If
    On Error GoTo 0

    VariantLen = upperBound - lowerBound + 1
    If VariantLen < 0 Then VariantLen = 0
End Function

'------------------------------------------------------------------------------
' Return item i using either 0-based or 1-based source arrays.
'------------------------------------------------------------------------------
Public Function VariantItem(ByRef values As Variant, ByVal oneBasedIndex As Long) As Variant
    Dim lowerBound As Long

    lowerBound = LBound(values)
    VariantItem = values(lowerBound + oneBasedIndex - 1)
End Function
